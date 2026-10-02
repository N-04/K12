"""验证 Windows 文件句柄、Office 适配器和任务执行边界。"""
import hashlib
import json
import ctypes
import sqlite3
import subprocess
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch, MagicMock

from k12 import converters, local_client
from k12.store import AppStore
from k12.processor import TaskProcessor


class WindowsCompatibilityTests(unittest.TestCase):
    """无需真实 Office 的跨平台回归测试。"""

    def test_windows_reserved_filenames_and_archive_segments_are_safe(self):
        """设备名称、尾部点和路径分隔符不会生成无效 Windows 文件路径。"""
        cases = {'CON.docx': '_CON.docx', 'nul.txt': '_nul.txt', 'COM1.pdf': '_COM1.pdf', 'LPT³.xlsx': '_LPT³.xlsx', 'normal.docx. ': 'normal.docx', r'C:\folder\课件.docx': '课件.docx'}
        with tempfile.TemporaryDirectory() as tmp:
            for original, expected in cases.items():
                safe = TaskProcessor._safe_file_name(original)
                self.assertEqual(safe, expected)
                (Path(tmp) / safe).write_bytes(b'fixture')
        self.assertEqual(TaskProcessor._safe_relative_path('目录/CON.docx'), '目录/_CON.docx')

    def test_sqlite_context_closes_connection_after_commit_and_rollback(self):
        """成功与异常事务都关闭连接，提交保存且异常回滚。"""
        with tempfile.TemporaryDirectory() as tmp:
            store = AppStore(tmp)
            with store.connect() as conn:
                conn.execute("CREATE TABLE connection_probe(value INTEGER)")
                conn.execute("INSERT INTO connection_probe VALUES (1)")
            with self.assertRaises(sqlite3.ProgrammingError):
                conn.execute("SELECT 1")
            with self.assertRaises(ValueError):
                with store.connect() as failed:
                    failed.execute("INSERT INTO connection_probe VALUES (2)")
                    raise ValueError("模拟事务失败")
            with self.assertRaises(sqlite3.ProgrammingError):
                failed.execute("SELECT 1")
            with store.connect() as check:
                self.assertEqual([row[0] for row in check.execute("SELECT value FROM connection_probe")], [1])
            store.db_path.unlink()

    def test_windows_normalizer_passes_paths_as_arguments_and_verifies_output(self):
        """中文与特殊字符路径保持独立参数，并检查生成产物的哈希。"""
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "中文 '特殊字符' $文件.doc"
            target = Path(tmp) / "中文 输出.docx"
            source.write_bytes(b"fixture")

            def run(command, **kwargs):
                """模拟 COM 执行并写出可验证文档。"""
                self.assertEqual(command[-4:], ['-SourcePath', str(source.resolve()), '-TargetPath', str(target.resolve())])
                self.assertNotIn(str(source), Path(command[command.index('-File') + 1]).read_text(encoding='utf-8'))
                converters.build_docx([{'text': 'Windows 中文测试'}], target)
                return SimpleNamespace(returncode=0, stdout=b'', stderr=b'')

            with patch.object(local_client, 'windows_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=run):
                result = local_client.normalize_windows_office_document(source, target, 'word')
            self.assertEqual(result['sha256'], hashlib.sha256(target.read_bytes()).hexdigest())
            self.assertEqual(result['size'], target.stat().st_size)

    def test_windows_normalizer_removes_partial_output_on_timeout(self):
        """超时后不遗留可被误认为成功的部分文件。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / 'input.doc', Path(tmp) / 'output.docx'
            source.write_bytes(b'fixture')

            def run(command, **kwargs):
                """模拟先写入部分产物再超时。"""
                target.write_bytes(b'partial')
                raise subprocess.TimeoutExpired(command, 1)

            with patch.object(local_client, 'windows_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=run):
                with self.assertRaises(local_client.LocalClientError):
                    local_client.normalize_windows_office_document(source, target, 'word')
            self.assertFalse(target.exists())

    def test_windows_normalizer_rejects_invalid_output_and_existing_target(self):
        """无效 OOXML 不能登记成功，也不能覆盖已有文件。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / 'input.ppt', Path(tmp) / 'output.pptx'
            source.write_bytes(b'fixture')

            def run(command, **kwargs):
                """模拟 Office 返回成功但产物损坏。"""
                target.write_bytes(b'not a ZIP')
                return SimpleNamespace(returncode=0)

            with patch.object(local_client, 'windows_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=run) as runner:
                with self.assertRaises(local_client.LocalClientError):
                    local_client.normalize_windows_office_document(source, target, 'powerpoint')
                self.assertFalse(target.exists())
                target.write_bytes(b'original')
                with self.assertRaises(local_client.LocalClientError):
                    local_client.normalize_windows_office_document(source, target, 'powerpoint')
                self.assertEqual(target.read_bytes(), b'original')
                self.assertEqual(runner.call_count, 1)

    def test_windows_task_requires_authorization_and_matching_platform(self):
        """实际执行前验证显式授权和当前平台。"""
        with self.assertRaises(local_client.LocalClientError):
            local_client.execute_windows_office_task({})
        with patch.object(local_client.platform, 'system', return_value='Darwin'):
            with self.assertRaises(local_client.LocalClientError):
                local_client.execute_windows_office_task({}, allow_native_execution=True)

    def test_owned_process_cleanup_checks_identity_and_waits(self):
        """退出、强制清理及 PID 复用均只操作已核对身份的句柄。"""
        for wait_codes, created, image, should_kill, raises in [
            ([0], 123, 'WINWORD.EXE', False, False),
            ([258, 0], 123, 'WINWORD.EXE', True, False),
            ([258], 999, 'WINWORD.EXE', False, True),
            ([258], 123, 'EXCEL.EXE', False, True),
            ([258, 258], 123, 'WINWORD.EXE', True, True),
        ]:
            with self.subTest(wait_codes=wait_codes, created=created, image=image):
                kernel = MagicMock()
                kernel.OpenProcess.return_value = 42
                kernel.WaitForSingleObject.side_effect = wait_codes
                kernel.TerminateProcess.return_value = 1
                def times(handle, creation, *unused):
                    creation._obj.dwLowDateTime = created
                    return 1
                def name(handle, flags, buffer, length):
                    buffer.value = image
                    return 1
                kernel.GetProcessTimes.side_effect = times
                kernel.QueryFullProcessImageNameW.side_effect = name
                with patch.object(ctypes, 'WinDLL', return_value=kernel, create=True):
                    record = {'application': 'word', 'pid': 20, 'creation_time': '123'}
                    if raises:
                        with self.assertRaises(local_client.LocalClientError):
                            local_client._cleanup_windows_office_process(record)
                    else:
                        local_client._cleanup_windows_office_process(record)
                self.assertEqual(kernel.TerminateProcess.called, should_kill)
                kernel.CloseHandle.assert_called_once_with(42)

    def test_script_supervisor_cleans_registered_instances_on_all_exits(self):
        """成功、失败和超时均清理登记实例，无法归属的初始化明确报错。"""
        for mode in ['success', 'failed', 'timeout', 'unknown']:
            with self.subTest(mode=mode):
                def run(command, **kwargs):
                    root = Path(command[command.index('-OwnershipDirectory') + 1])
                    if mode == 'unknown':
                        (root / 'word.starting').write_text('word')
                    else:
                        (root / 'word.json').write_text(json.dumps({'application': 'word', 'pid': 20, 'creation_time': '123'}))
                    if mode in {'timeout', 'unknown'}:
                        raise subprocess.TimeoutExpired(command, 1)
                    return SimpleNamespace(returncode=1 if mode == 'failed' else 0)
                with patch.object(local_client.subprocess, 'run', side_effect=run), patch.object(local_client, '_cleanup_windows_office_process') as cleanup:
                    if mode in {'timeout', 'unknown'}:
                        with self.assertRaisesRegex(local_client.LocalClientError, '无法确认进程归属' if mode == 'unknown' else '已清理'):
                            local_client.run_windows_office_script(Path('fixture.ps1'), [])
                    else:
                        result = local_client.run_windows_office_script(Path('fixture.ps1'), [])
                        self.assertEqual(result.returncode, mode == 'failed')
                    self.assertEqual(cleanup.call_count, 0 if mode == 'unknown' else 1)

    def test_mixed_batches_select_pending_ids_on_both_platforms(self):
        """混合批次仅转换待处理文件，去重后正确计数且保留既有产物。"""
        for system, executor in [('Windows', local_client.execute_windows_office_task), ('Darwin', local_client.execute_macos_office_task)]:
            for task_type, legacy, modern in [('word_to_ppt', '.doc', '.docx'), ('ppt_to_word', '.ppt', '.pptx')]:
                with self.subTest(system=system, task_type=task_type), tempfile.TemporaryDirectory() as tmp:
                    root = Path(tmp)
                    existing = root / 'existing-output.docx'
                    existing.write_bytes(b'original')
                    files = [{'id': 'old', 'extension': legacy}, {'id': 'modern', 'extension': modern}]
                    payload = {'task': {'task_type': task_type}, 'files': files,
                               'local_actions': [{'type': 'office_conversion', 'pending_files': [{'file_id': 'old'}, {'file_id': 'old'}]}],
                               'desktop_execution_plan': {'actions': [{'type': 'office_conversion', 'gate_status': 'ready', 'output_contract': {'output_directory': str(root)}}]}}
                    with patch.object(local_client.platform, 'system', return_value=system), patch.object(local_client, '_execute_macos_office_file', return_value={'file_id': 'old', 'output_type': 'pptx'}) as convert:
                        result = executor(payload, True)
                    self.assertEqual(result['status'], 'success')
                    convert.assert_called_once()
                    self.assertEqual(convert.call_args.args[0]['id'], 'old')
                    self.assertEqual(existing.read_bytes(), b'original')

    def test_pending_list_empty_missing_unknown_and_rollback(self):
        """空清单跳过，错误清单在启动前拒绝，失败回滚仅影响本次产物。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            existing = root / 'existing.pptx'
            existing.write_bytes(b'original')
            payload = {'task': {'task_type': 'word_to_ppt'}, 'files': [{'id': 'a'}, {'id': 'b'}],
                       'local_actions': [{'type': 'office_conversion', 'pending_files': []}],
                       'desktop_execution_plan': {'actions': [{'type': 'office_conversion', 'gate_status': 'ready', 'output_contract': {'output_directory': str(root)}}]}}
            with patch.object(local_client.platform, 'system', return_value='Windows'), patch.object(local_client, '_execute_macos_office_file') as convert:
                result = local_client.execute_windows_office_task(payload, True)
                self.assertEqual(result['status'], 'success')
                self.assertFalse(result['native_execution_performed'])
                for pending in [None, [{'file_id': 'missing'}], [{}]]:
                    payload['local_actions'][0]['pending_files'] = pending
                    with self.assertRaises(local_client.LocalClientError):
                        local_client.execute_windows_office_task(payload, True)
                convert.assert_not_called()
                payload['local_actions'][0]['pending_files'] = [{'file_id': 'a'}, {'file_id': 'b'}]
                new_output = root / 'new.pptx'
                new_output.write_bytes(b'new')
                convert.side_effect = [{'file_id': 'a', 'path': str(new_output), 'output_type': 'pptx'}, local_client.LocalClientError('模拟失败')]
                result = local_client.execute_windows_office_task(payload, True)
                self.assertEqual(result['status'], 'failed')
                self.assertEqual(result['rollback_count'], 1)
                self.assertFalse(new_output.exists())
                self.assertEqual(existing.read_bytes(), b'original')

    def test_empty_pending_run_does_not_sync_a_running_report(self):
        """无原生动作时直接返回，不发送可能降低任务状态的同步请求。"""
        requests = []
        def request(origin, path, **kwargs):
            requests.append(path)
            if path.endswith('heartbeat'):
                return {}
            if path.endswith('manifest'):
                return {}
            if path.endswith('local-payload'):
                return {'task': {'id': 'done-task'}, 'local_actions': []}
            return {}
        empty = {'schema_version': 'k12.windowsOfficeTaskExecution.v1', 'status': 'success',
                 'native_execution_performed': False, 'outputs': [], 'failure_count': 0}
        with patch.object(local_client, 'request_json', side_effect=request), patch.object(local_client, 'build_heartbeat', return_value={}), patch.object(local_client, 'execute_windows_office_task', return_value=empty), patch.object(local_client.platform, 'system', return_value='Windows'):
            result = local_client.run_once('http://127.0.0.1:8765', token='fixture-token', task_id='done-task', allow_native_execution=True, execute_native_office=True)
        self.assertEqual(result['native_office_execution']['output_count'], 0)
        self.assertFalse(any(path.endswith('local-sync') for path in requests))

    def test_mixed_batches_sync_counts_and_downloads(self):
        """跨模块校验混合任务报告计数、既有产物及实际 HTTP 下载。"""
        import threading
        from urllib.request import Request, urlopen
        from k12.server import K12Server, K12RequestHandler
        for task_type, legacy_suffix, modern_suffix in [('word_to_ppt', 'doc', 'docx'), ('ppt_to_word', 'ppt', 'pptx')]:
            with self.subTest(task_type=task_type), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                server = K12Server(('127.0.0.1', 0), K12RequestHandler, root / 'data')
                store, processor = server.store, server.processor
                token = 'mixed-regression-token'
                store.update_settings({'localSecurityToken': token, 'localClientPlatform': 'Windows', 'allowWebLaunchLocalClient': True, 'allowTaskStatusCloudSync': True})
                processor.record_local_client_heartbeat({'client_id': 'mixed-regression', 'status': 'online', 'platform': 'Windows',
                    'capabilities': {'officeAutomation': True}, 'preflight': {'platform': 'Windows', 'components': {
                    'word': {'available': True, 'status': 'available'}, 'powerpoint': {'available': True, 'status': 'available'}},
                    'capabilities': {'officeAutomation': True}, 'executes_native_documents': True}})
                legacy = bytearray(512)
                legacy[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
                legacy[24:26] = (0x003E).to_bytes(2, 'little')
                legacy[26:28] = (3).to_bytes(2, 'little')
                legacy[28:30] = bytes.fromhex('feff')
                legacy[30:32] = (9).to_bytes(2, 'little')
                legacy[32:34] = (6).to_bytes(2, 'little')
                modern = root / ('modern.' + modern_suffix)
                if modern_suffix == 'docx':
                    converters.build_docx([{'text': '混合任务现代文档'}], modern)
                else:
                    converters.build_pptx([{'title': '混合任务现代演示', 'body': ['教学内容']}], modern)
                old_file = processor.create_uploaded_file('legacy.' + legacy_suffix, bytes(legacy))[0]
                new_file = processor.create_uploaded_file(modern.name, modern.read_bytes())[0]
                task = processor.create_task({'task_type': task_type, 'file_ids': [old_file['id'], new_file['id']]})
                payload = processor.local_task_payload(task['id'])
                report = processor._latest_report_for_task(task['id'])
                existing = next(item for item in report['analysis']['artifacts'] if item['file_id'] == new_file['id'])
                original = Path(existing['path']).read_bytes()
                def normalize(source, target, application, timeout_seconds=120):
                    if application == 'word':
                        converters.build_docx([{'text': '模拟旧文档规范化'}], target)
                    else:
                        converters.build_pptx([{'title': '模拟旧演示规范化', 'body': ['教学内容']}], target)
                    return {'native_execution_performed': True}
                with patch.object(local_client.platform, 'system', return_value='Windows'), patch.object(local_client, 'normalize_windows_office_document', side_effect=normalize) as normalizer:
                    execution = local_client.execute_windows_office_task(payload, True)
                self.assertEqual(normalizer.call_count, 1)
                sync_payload = local_client.build_native_report_sync_payload(payload, execution, platform_name='Windows', client_id='mixed-regression')
                sync_payload['outputs'] = execution['outputs']
                synced = processor.sync_local_task_status(task['id'], sync_payload)
                self.assertEqual(synced['status'], '成功')
                self.assertEqual(synced['success_count'], 2)
                self.assertEqual(synced['pending_count'], 0)
                self.assertEqual(Path(existing['path']).read_bytes(), original)
                report = processor._latest_report_for_task(task['id'])
                self.assertEqual(report['artifact_count'], 2)
                self.assertEqual(report['artifact_pending_count'], 0)
                thread = threading.Thread(target=server.serve_forever, daemon=True)
                thread.start()
                try:
                    for artifact in report['analysis']['artifacts']:
                        url = f'http://127.0.0.1:{server.server_port}' + artifact['url']
                        with urlopen(Request(url, headers={'X-K12-Token': token}), timeout=10) as response:
                            data = response.read()
                        self.assertEqual(hashlib.sha256(data).hexdigest(), artifact['sha256'])
                finally:
                    server.shutdown()
                    server.server_close()
                    thread.join(timeout=5)

    def test_powershell_cleanup_continues_after_document_close_failure(self):
        """真实 PowerShell 验证文档关闭异常仍继续调用应用退出。"""
        import shutil
        if not shutil.which('powershell.exe'):
            self.skipTest('当前设备没有 Windows PowerShell')
        script = Path(local_client.__file__).with_name('windows_office_lifecycle.ps1')
        code = "& { param($helper); . $helper; $global:quitCalled=$false; $doc=New-Object PSObject; $doc | Add-Member ScriptMethod Close { throw '关闭失败' }; $app=New-Object PSObject; $app | Add-Member ScriptMethod Quit { $global:quitCalled=$true }; Close-K12OfficeObjects $doc $app 'word'; if(-not $global:quitCalled){exit 3}; function Get-Process { return @{ProcessName='WINWORD'} }; try { New-K12OfficeApplication 'word' $env:TEMP; exit 4 } catch { exit 0 } }"
        result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', code, str(script)], capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors='replace'))


if __name__ == '__main__':
    unittest.main()
