"""验证 Windows 文件句柄、Office 适配器和任务执行边界。"""
import hashlib
import json
import ctypes
import sqlite3
import subprocess
import tempfile
import unittest
import zipfile
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

    def test_macos_normalizer_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / 'input.doc', Path(tmp) / 'output.docx'
            source.write_bytes(b'fixture')
            target.write_bytes(b'original')
            with patch.object(local_client, 'macos_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run') as runner:
                with self.assertRaisesRegex(local_client.LocalClientError, '拒绝覆盖'):
                    local_client.normalize_macos_office_document(source, target, 'word')
            runner.assert_not_called()
            self.assertEqual(target.read_bytes(), b'original')

    def test_macos_normalizer_cleans_failed_outputs_and_hides_office_paths(self):
        for mode in ['timeout', 'failed', 'empty']:
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as tmp:
                source, target = Path(tmp) / 'input.doc', Path(tmp) / 'output.docx'
                source.write_bytes(b'fixture')

                def run(command, **kwargs):
                    target.write_bytes(b'' if mode == 'empty' else b'partial')
                    if mode == 'timeout':
                        raise subprocess.TimeoutExpired(command, 1)
                    return SimpleNamespace(returncode=1 if mode == 'failed' else 0, stderr=str(source), stdout='')

                with patch.object(local_client, 'macos_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=run):
                    with self.assertRaises(local_client.LocalClientError) as caught:
                        local_client.normalize_macos_office_document(source, target, 'word')
                self.assertFalse(target.exists())
                self.assertNotIn(str(source), str(caught.exception))

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

    def test_native_conversions_apply_pending_settings_to_real_outputs(self):
        """旧 Office 二段转换的参数必须影响实际 OOXML 内容。"""
        for system, executor in [('Windows', local_client.execute_windows_office_task), ('Darwin', local_client.execute_macos_office_task)]:
            for task_type in ['word_to_ppt', 'ppt_to_word']:
                for enabled in [False, True]:
                    with self.subTest(system=system, task_type=task_type, enabled=enabled), tempfile.TemporaryDirectory() as tmp:
                        root = Path(tmp)
                        (root / '.inputs').mkdir()
                        suffix = '.doc' if task_type == 'word_to_ppt' else '.ppt'
                        source = root / '.inputs' / ('lesson' + suffix)
                        source.write_bytes(b'legacy fixture')
                        file = {'id': 'legacy', 'file_name': source.name, 'extension': suffix,
                                'input_path': str(source), 'file_size': source.stat().st_size,
                                'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest()}
                        settings = {'word_generate_toc': enabled, 'word_max_chars_per_slide': 40,
                                    'word_auto_pagination': enabled, 'ppt_to_word_generate_toc': enabled,
                                    'ppt_extract_notes': enabled, 'ppt_to_word_mode': '备注优先模式',
                                    'ppt_to_word_template': '教学模板'}
                        payload = {'task': {'task_type': task_type}, 'files': [file],
                                   'local_actions': [{'type': 'office_conversion', 'pending_files': [
                                       {'file_id': 'legacy', 'conversion_settings': settings}]}],
                                   'desktop_execution_plan': {'actions': [{'type': 'office_conversion',
                                       'gate_status': 'ready', 'output_contract': {'output_directory': str(root)}}]}}

                        def normalize(source, target, application):
                            if application == 'word':
                                converters.build_docx([{'text': '课程标题', 'style': 'Heading1'},
                                    {'text': '教学内容' * 20}, {'text': '练习内容' * 20}], target)
                            else:
                                converters.build_pptx([{'title': '课程标题', 'body': ['正文']}], target)

                        adapter = 'normalize_windows_office_document' if system == 'Windows' else 'normalize_macos_office_document'
                        with patch.object(local_client.platform, 'system', return_value=system), patch.object(local_client, adapter, side_effect=normalize), patch.object(local_client, 'extract_pptx_slides', return_value=[
                                {'title': '课程标题', 'body': ['正文'], 'notes': ['教师备注']} ]):
                            result = executor(payload, True)
                        self.assertEqual(result['status'], 'success')
                        output = Path(result['outputs'][0]['path'])
                        if task_type == 'word_to_ppt':
                            slides = converters.extract_pptx_slides(output)
                            self.assertEqual(slides[0]['title'] == '目录', enabled)
                            if enabled:
                                self.assertGreater(len(slides), 2)
                            else:
                                self.assertEqual(len(slides), 1)
                        else:
                            text = ' '.join(block['text'] for block in converters.extract_docx_blocks(output))
                            self.assertEqual('目录' in text, enabled)
                            self.assertEqual('教师备注' in text, enabled)
                            self.assertIn('教学模板', text)
                            self.assertIn('备注优先模式', text)
                            if enabled:
                                self.assertLess(text.index('教师备注'), text.index('正文'))
                        self.assertFalse(list(root.glob('.native-*')))

    def test_word_pagination_preserves_long_text_titles_and_full_toc(self):
        """长段落、长标题和超过一页的目录都不能截断。"""
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'lesson.pptx'
            title = '长标题' * 40
            text = '正文字符' * 103
            blocks = [{'text': title, 'level': 1}, {'text': text, 'level': 0}]
            converters.build_pptx_from_docx(blocks, target, max_chars=60)
            slides = converters.extract_pptx_slides(target)
            self.assertTrue(all(slide['title'] == title for slide in slides))
            bodies = [item for slide in slides for item in slide['body']]
            self.assertEqual(''.join(bodies), text)
            self.assertTrue(all(sum(len(item) for item in slide['body']) <= 60 for slide in slides))
            converters.build_pptx_from_docx(blocks, target, max_chars=60, auto_pagination=False)
            self.assertEqual(len(converters.extract_pptx_slides(target)), 1)
            chapters = [{'text': f'章节{i:02d}', 'level': 1} for i in range(25)]
            converters.build_pptx_from_docx(chapters, target, generate_toc=True)
            slides = converters.extract_pptx_slides(target)
            toc = [slide for slide in slides if slide['title'].startswith('目录')]
            self.assertEqual(len(toc), 2)
            self.assertEqual(' '.join(item for slide in toc for item in slide['body']),
                             ' '.join(block['text'] for block in chapters))
            with self.assertRaisesRegex(ValueError, '必须大于零'):
                converters.build_pptx_from_docx(blocks, target, max_chars=-1)

    def test_excel_exports_keep_all_selected_cells(self):
        """超过预览条数的数据必须完整进入三种转换产物。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sheets = [{'name': 'Data', 'index': 1, 'cells': [
                {'ref': f'A{index}', 'value': f'CELL_{index:03d}_END'} for index in range(1, 76)]}]
            docx, pptx, pdf = root / 'data.docx', root / 'data.pptx', root / 'data.pdf'
            converters.build_docx_from_xlsx(sheets, docx, 'Data')
            converters.build_pptx_from_xlsx(sheets, pptx, 'Data')
            converters.build_pdf_from_xlsx(sheets, pdf, 'Data')
            word_text = ' '.join(block['text'] for block in converters.extract_docx_blocks(docx))
            slides = converters.extract_pptx_slides(pptx)
            ppt_text = ' '.join(item for slide in slides for item in slide['body'])
            pdf_text = ' '.join(block['text'] for block in converters.extract_pdf_text_blocks(pdf))
            self.assertGreater(len(slides), 1)
            for index in range(1, 76):
                marker = f'CELL_{index:03d}_END'
                for text in [word_text, ppt_text, pdf_text]:
                    self.assertEqual(text.count(marker), 1)
            outline = root / 'outline.docx'
            converters.build_docx_from_slides([{'title': 'Data', 'body': [f'ITEM_{i}' for i in range(20)]}], outline, mode='大纲模式')
            outline_text = ' '.join(block['text'] for block in converters.extract_docx_blocks(outline))
            self.assertIn('ITEM_19', outline_text)

    def test_ppt_word_table_retention_option_applies_in_outline_and_handout(self):
        """表格保留开关在讲义和大纲模式均影响实际产物。"""
        import zipfile
        import xml.etree.ElementTree as ET
        for mode in ['大纲模式', '逐页讲义模式']:
            for retain in [True, False]:
                with self.subTest(mode=mode, retain=retain), tempfile.TemporaryDirectory() as tmp:
                    target = Path(tmp) / 'table.docx'
                    converters.build_docx_from_slides([{'title': '页标题', 'body': ['正文'], 'tables': [[['唯一表格内容']]]}], target, mode=mode, retain_tables=retain)
                    with zipfile.ZipFile(target) as archive:
                        root = ET.fromstring(archive.read('word/document.xml'))
                    self.assertEqual(root.find('.//{*}tbl') is not None, retain)
                    self.assertEqual('唯一表格内容' in ''.join(root.itertext()), retain)

    def test_excel_word_tables_preserve_sparse_coordinates_without_expansion(self):
        """稀疏单元格保留坐标，避免按最大行列生成巨表。"""
        import zipfile
        import xml.etree.ElementTree as ET
        cells = [{'ref': 'B2', 'value': '姓名'}, {'ref': 'D2', 'value': '成绩'},
                 {'ref': 'B9', 'value': '张三'}, {'ref': 'D9', 'value': '95'},
                 {'ref': 'XFD1048576', 'value': '末尾'}]
        rows = converters._xlsx_word_tables(cells)[0]
        self.assertEqual(rows[0], ['行号', 'B', 'D', 'XFD'])
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[2], ['9', 'B9: 张三', 'D9: 95', ''])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'sparse.docx'
            converters.build_docx_from_xlsx([{'name': '成绩', 'cells': cells}], target, '成绩')
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read('word/document.xml'))
            self.assertIsNotNone(root.find('.//{*}tbl'))
            text = ''.join(root.itertext())
            for cell in cells:
                self.assertEqual(text.count(f'{cell["ref"]}: {cell["value"]}'), 1)

    def test_ppt_tables_export_as_editable_word_cells(self):
        """PPT 表格行列内容进入 Word 原生表格。"""
        import zipfile
        import xml.etree.ElementTree as ET
        table_xml = '<root xmlns:a="urn:a"><a:tbl><a:tr><a:tc><a:p><a:r><a:t>姓名</a:t></a:r></a:p></a:tc><a:tc><a:p><a:r><a:t>成绩</a:t></a:r></a:p></a:tc></a:tr><a:tr><a:tc><a:p><a:r><a:t>张三</a:t></a:r></a:p></a:tc><a:tc><a:p><a:r><a:t>95</a:t></a:r></a:p></a:tc></a:tr></a:tbl></root>'
        tables = converters._pptx_tables(table_xml)
        self.assertEqual(tables, [[['姓名', '成绩'], ['张三', '95']]])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'table.docx'
            converters.build_docx_from_slides([{'title': '成绩表', 'body': [], 'tables': tables}], target)
            with zipfile.ZipFile(target) as archive:
                root = ET.fromstring(archive.read('word/document.xml'))
            table = root.find('.//{*}tbl')
            rows = table.findall('{*}tr')
            self.assertEqual(len(rows), 2)
            self.assertEqual([[cell.findtext('.//{*}t') for cell in row.findall('{*}tc')] for row in rows], tables[0])

    def test_ppt_notes_exclude_metadata_and_join_text_runs(self):
        """备注排除日期页码，文字片段按段落合并。"""
        import zipfile
        self.assertEqual(converters._pptx_texts('<s xmlns:a="urn:test"><a:p><a:r><a:t>数学</a:t></a:r><a:r><a:t>课堂</a:t></a:r></a:p></s>'), ['数学课堂'])
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'notes.zip'
            shapes = []
            for kind, text in [('body', '教师备注'), ('sldNum', '1'), ('dt', '2026'), ('ftr', '页脚')]:
                shapes.append(f'<p:sp><p:nvSpPr><p:nvPr><p:ph type="{kind}"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')
            with zipfile.ZipFile(target, 'w') as archive:
                archive.writestr('ppt/notesSlides/notesSlide1.xml', '<p:notes xmlns:p="urn:p" xmlns:a="urn:a">' + ''.join(shapes) + '</p:notes>')
            with zipfile.ZipFile(target) as archive:
                notes = converters._pptx_notes(archive, ['ppt/notesSlides/notesSlide1.xml'])
            self.assertEqual(notes, ['教师备注'])

    def test_word_object_counts_exclude_containers_and_properties(self):
        """公式容器、段落和表格属性不能增加实际对象数量。"""
        import zipfile
        xml = '<d:document xmlns:d="urn:word" xmlns:q="urn:math"><d:p><d:pPr/><q:oMathPara><q:oMath><q:r/></q:oMath></q:oMathPara></d:p><d:tbl><d:tblPr/></d:tbl></d:document>'
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'formula.docx'
            with zipfile.ZipFile(source, 'w') as archive:
                archive.writestr('word/document.xml', xml)
                archive.writestr('[Content_Types].xml', '<Types/>')
            store = AppStore(Path(tmp) / 'data')
            file = TaskProcessor(store).create_uploaded_file(source.name, source.read_bytes())[0]
            self.assertTrue(file['has_omml'])
            self.assertEqual(file['content_summary']['ommlFormulas'], 1)
            self.assertEqual(file['content_summary']['paragraphs'], 1)
            self.assertEqual(file['content_summary']['tables'], 1)
            summary = converters.extract_docx_object_summary(source)
            self.assertEqual(summary['paragraphs'], 1)
            self.assertEqual(summary['tables'], 1)
            self.assertEqual(summary['omml_formulas'], 1)
            self.assertEqual(summary['formulas'], 1)
            self.assertEqual(converters.count_xml_elements('<m:oMathPara/><m:oMathExtra/>', 'oMath'), 0)

    def test_pdf_chinese_text_is_preserved(self):
        """中文 PDF 使用 Unicode 文本，不再替换为问号。"""
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'chinese.pdf'
            converters.build_text_pdf(['中文导出测试', 'Hello 123'], target)
            content = target.read_bytes()
            self.assertIn(b'/UniGB-UCS2-H', content)
            text = ' '.join(block['text'] for block in converters.extract_pdf_text_blocks(target))
            self.assertIn('中文导出测试', text)
            self.assertIn('Hello 123', text)
            self.assertNotIn('?', text)

    def test_generated_word_styles_and_controls_round_trip(self):
        """DOCX 样式实际登记，换行与制表符往返不丢失。"""
        import zipfile
        import xml.etree.ElementTree as ET
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'styles.docx'
            converters.build_docx([{'text': '标题', 'style': 'Heading1'}, {'text': '小节', 'style': 'Heading2'}, {'text': ' 空格\n换行\t制表符 '}], target)
            blocks = converters.extract_docx_blocks(target)
            self.assertEqual([block['level'] for block in blocks], [1, 2, 0])
            self.assertEqual(blocks[2]['text'], ' 空格\n换行\t制表符 ')
            with zipfile.ZipFile(target) as archive:
                styles = ET.fromstring(archive.read('word/styles.xml'))
                self.assertEqual(len(styles), 3)
                self.assertIn(b'styles.xml', archive.read('word/_rels/document.xml.rels'))
                self.assertIn(b'word/styles.xml', archive.read('[Content_Types].xml'))

    def test_generated_ppt_textboxes_have_layout_and_complete_nonvisual_properties(self):
        """标题和正文有完整可编辑形状结构，行内容分别保留。"""
        import zipfile
        import xml.etree.ElementTree as ET
        with tempfile.TemporaryDirectory() as tmp:
            ppt = Path(tmp) / 'layout.pptx'
            converters.build_pptx([{'title': '中文标题', 'body': ['第一行', '第二行 & <测试>']}], ppt)
            with zipfile.ZipFile(ppt) as archive:
                root = ET.fromstring(archive.read('ppt/slides/slide1.xml'))
            ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            shapes = root.findall('.//p:sp', ns)
            self.assertEqual(len(shapes), 2)
            for shape in shapes:
                self.assertIsNotNone(shape.find('p:nvSpPr/p:cNvSpPr', ns))
                self.assertIsNotNone(shape.find('p:nvSpPr/p:nvPr', ns))
                size = shape.find('p:spPr/a:xfrm/a:ext', ns)
                self.assertGreater(int(size.get('cx')), 0)
                self.assertGreater(int(size.get('cy')), 0)
            self.assertEqual([node.text for node in shapes[1].findall('.//a:t', ns)], ['第一行', '第二行 & <测试>'])

    def test_word_secondary_headings_and_text_controls_are_preserved(self):
        """二级标题创建内容页，替代前缀不影响换行与多段文字。"""
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            word = root / 'lesson.docx'
            with zipfile.ZipFile(word, 'w') as archive:
                archive.writestr('word/document.xml', '<d:document xmlns:d="urn:test"><d:body><d:p><d:pPr><d:pStyle d:val="Heading1"/></d:pPr><d:r><d:t>章节</d:t></d:r></d:p><d:p><d:pPr><d:pStyle d:val="Heading2"/></d:pPr><d:r><d:t>小节</d:t></d:r></d:p><d:p><d:r><d:t>正文一</d:t><d:br/><d:t>正文二</d:t><d:tab/><d:t>正文三</d:t></d:r></d:p></d:body></d:document>')
            blocks = converters.extract_docx_blocks(word)
            self.assertEqual([block['level'] for block in blocks], [1, 2, 0])
            self.assertEqual(blocks[2]['text'], '正文一\n正文二\t正文三')
            ppt = root / 'lesson.pptx'
            converters.build_pptx_from_docx(blocks, ppt)
            slides = converters.extract_pptx_slides(ppt)
            self.assertEqual([slide['title'] for slide in slides], ['章节', '小节'])
            self.assertIn('正文三', ' '.join(slides[1]['body']))

    def test_reordered_ppt_preserves_page_content_and_notes(self):
        """重排后的页面与对应备注一起进入 Word 输出。"""
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            ppt = root / 'ordered.pptx'
            converters.build_pptx([{'title': '第一页', 'body': ['内容一']}, {'title': '第二页', 'body': ['内容二']}], ppt)
            with zipfile.ZipFile(ppt) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts['ppt/presentation.xml'] = b'<presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sldIdLst><sldId r:id="rId2"/><sldId r:id="rId1"/></sldIdLst></presentation>'
            for index in [1, 2]:
                parts[f'ppt/slides/_rels/slide{index}.xml.rels'] = f'<Relationships><Relationship Target="../notesSlides/notesSlide{index}.xml"/></Relationships>'.encode()
                parts[f'ppt/notesSlides/notesSlide{index}.xml'] = f'<notes xmlns:a="urn:test"><a:t>备注{index}</a:t></notes>'.encode()
            with zipfile.ZipFile(ppt, 'w') as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            slides = converters.extract_pptx_slides(ppt)
            self.assertEqual([slide['title'] for slide in slides], ['第二页', '第一页'])
            self.assertEqual([slide['notes'] for slide in slides], [['备注2'], ['备注1']])
            word = root / 'ordered.docx'
            converters.build_docx_from_slides(slides, word, generate_toc=False)
            text = ' '.join(block['text'] for block in converters.extract_docx_blocks(word))
            self.assertLess(text.index('内容二'), text.index('内容一'))
            self.assertLess(text.index('备注2'), text.index('备注1'))

    def test_excel_rich_text_and_namespaced_cells_preserve_content(self):
        """多个富文本片段、前缀和 XML 实体只解码一次。"""
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'rich.xlsx'
            with zipfile.ZipFile(source, 'w') as archive:
                archive.writestr('xl/workbook.xml', '<workbook><sheets><sheet name="Data"/></sheets></workbook>')
                archive.writestr('xl/sharedStrings.xml', '<s:sst xmlns:s="urn:test"><s:si><s:r><s:t>保留 </s:t></s:r><s:r><s:t>&amp;lt;文字</s:t></s:r></s:si></s:sst>')
                archive.writestr('xl/worksheets/sheet1.xml', '<s:worksheet xmlns:s="urn:test"><s:c r="A1" t="inlineStr"><s:is><s:r><s:t>第一段 </s:t></s:r><s:r><s:t>第二段</s:t></s:r></s:is></s:c><s:c r="B1" t="s"><s:v>0</s:v></s:c><s:c r="C1"><s:f>1-1</s:f><s:v>0</s:v></s:c></s:worksheet>')
            sheet = converters.extract_xlsx_sheets(source)[0]
            self.assertEqual([cell['value'] for cell in sheet['cells']], ['第一段 第二段', '保留 &lt;文字', '0 (=1-1)'])
            self.assertEqual(sheet['formula_count'], 1)
            with self.assertRaisesRegex(ValueError, '索引无效'):
                converters._xlsx_cells('<worksheet><c t="s"><v>99</v></c></worksheet>', ['文本'])

    def test_excel_relationships_preserve_workbook_order(self):
        """工作表文件编号与显示顺序不同仍必须正确匹配名称和内容。"""
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / 'ordered.xlsx'
            for external in [False, True]:
                with zipfile.ZipFile(target, 'w') as archive:
                    archive.writestr('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="先看" r:id="r2"/><sheet name="后看" r:id="r1"/></sheets></workbook>')
                    mode = ' TargetMode="External"' if external else ''
                    archive.writestr('xl/_rels/workbook.xml.rels', f'<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="worksheets/sheet2.xml"{mode}/></Relationships>')
                    for index in [1, 2]:
                        archive.writestr(f'xl/worksheets/sheet{index}.xml', f'<worksheet><c r="A1"><v>{index}</v></c></worksheet>')
                if external:
                    with self.assertRaisesRegex(ValueError, '外部资源'):
                        converters.extract_xlsx_sheets(target)
                else:
                    sheets = converters.extract_xlsx_sheets(target)
                    self.assertEqual([sheet['name'] for sheet in sheets], ['先看', '后看'])
                    self.assertEqual([sheet['cells'][0]['value'] for sheet in sheets], ['2', '1'])

    def test_excel_named_sheet_selection_applies_range_and_rejects_missing_sheet(self):
        """指定工作表与选区组合生效，缺失工作表不能静默替换。"""
        from k12.excel import select_excel_sheets
        sheets = [{'name': name, 'index': index, 'formula_count': 2, 'cells': [
            {'ref': 'A1', 'value': '=1-1', 'formula': '1-1', 'result': '0'},
            {'ref': 'B1', 'value': '=2', 'formula': '2', 'result': '2'}]}
            for index, name in enumerate(['第一张', '成绩'], 1)]
        selected = select_excel_sheets(sheets, {'sheetNames': ['成绩'], 'conversionRange': '选区',
                                              'cellRefs': ['B1']}, {})
        self.assertEqual([sheet['name'] for sheet in selected], ['成绩'])
        self.assertEqual([cell['ref'] for cell in selected[0]['cells']], ['B1'])
        self.assertEqual(selected[0]['formula_count'], 1)
        self.assertEqual(len(sheets[1]['cells']), 2)
        for options in [{'sheetNames': ['不存在']}, {'sheetIndexes': [99]}]:
            with self.subTest(options=options), self.assertRaisesRegex(ValueError, '工作表不存在'):
                select_excel_sheets(sheets, options, {})

    def test_excel_selection_keeps_zero_results_and_filters_cells(self):
        """计算结果零不能当作缺失，选区不携带其他单元格。"""
        from k12.excel import select_excel_sheets
        sheets = [{'name': '成绩', 'index': 1, 'cells': [
            {'ref': 'A1', 'formula': '1-1', 'result': 0, 'value': '=1-1'},
            {'ref': 'B1', 'value': '未选择'}]}]
        selected = select_excel_sheets(sheets, {'conversionRange': '选区', 'cellRefs': ['A1'], 'retainFormulas': False}, {})
        self.assertEqual(len(selected[0]['cells']), 1)
        self.assertEqual(selected[0]['cells'][0]['value'], 0)
        self.assertEqual(sheets[0]['cells'][0]['value'], '=1-1')

    def test_macos_excel_normalization_uses_isolated_output(self):
        """Excel 规范化使用参数化脚本与独立输出，并拒绝错误源格式。"""
        from test_processor import make_xlsx_bytes
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / 'legacy.xls', Path(tmp) / 'modern.xlsx'
            source.write_bytes(b'legacy-fixture')
            def normalize(command, **kwargs):
                """模拟 Office 保存产物，不启动真实 Excel。"""
                target.write_bytes(make_xlsx_bytes())
                self.assertIn('editable true', command[2])
                self.assertIn('do not update links', command[2])
                self.assertIn('close workbookRef saving no', command[2])
                return SimpleNamespace(returncode=0)
            with patch.object(local_client, 'macos_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=normalize) as runner:
                result = local_client.normalize_macos_office_document(source, target, 'excel')
            self.assertEqual(result['output_type'], 'xlsx')
            self.assertTrue(result['native_execution_performed'])
            self.assertEqual(runner.call_count, 1)
            wrong = Path(tmp) / 'wrong.doc'
            wrong.write_bytes(b'wrong-format')
            with patch.object(local_client, 'macos_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run') as runner:
                with self.assertRaisesRegex(local_client.LocalClientError, '源文件必须'):
                    local_client.normalize_macos_office_document(wrong, Path(tmp) / 'wrong.xlsx', 'excel')
            runner.assert_not_called()

    def test_macos_excel_rejects_broken_sheet_and_removes_partial_output(self):
        """有工作簿标识但缺少工作表的 ZIP 不能当成有效 Excel 输出。"""
        with tempfile.TemporaryDirectory() as tmp:
            source, target = Path(tmp) / 'legacy.xls', Path(tmp) / 'broken.xlsx'
            source.write_bytes(b'legacy-fixture')
            def normalize(command, **kwargs):
                """模拟 Office 写入缺少关系部件的半成品。"""
                with zipfile.ZipFile(target, 'w') as archive:
                    archive.writestr('[Content_Types].xml', '<Types/>')
                    archive.writestr('xl/workbook.xml', '<workbook><sheets><sheet name="缺少关系"/></sheets></workbook>')
                return SimpleNamespace(returncode=0)
            with patch.object(local_client, 'macos_office_adapter_available', return_value=True), patch.object(local_client.subprocess, 'run', side_effect=normalize):
                with self.assertRaisesRegex(local_client.LocalClientError, '工作表结构或关系无效'):
                    local_client.normalize_macos_office_document(source, target, 'excel')
            self.assertFalse(target.exists())
            self.assertEqual(source.read_bytes(), b'legacy-fixture')

    def test_mixed_batches_sync_counts_and_downloads(self):
        """跨模块校验混合任务报告计数、既有产物及实际 HTTP 下载。"""
        import threading
        from urllib.request import Request, urlopen
        from k12.server import K12Server, K12RequestHandler
        from test_processor import make_xlsx_bytes, make_xlsx_object_files
        cases = [('word_to_ppt', 'doc', 'docx'), ('ppt_to_word', 'ppt', 'pptx')]
        cases += [(task, 'xls', 'xlsx') for task in ['excel_to_pdf', 'excel_to_word', 'excel_to_ppt']]
        cases = [(platform_name, *case) for platform_name in ["Windows", "macOS"] for case in cases]
        for platform_name, task_type, legacy_suffix, modern_suffix in cases:
            with self.subTest(task_type=task_type, platform=platform_name), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                server = K12Server(('127.0.0.1', 0), K12RequestHandler, root / 'data')
                store, processor = server.store, server.processor
                token = 'mixed-regression-token'
                store.update_settings({'localSecurityToken': token, 'localClientPlatform': platform_name, 'allowWebLaunchLocalClient': True, 'allowTaskStatusCloudSync': True})
                processor.record_local_client_heartbeat({'client_id': 'mixed-regression', 'status': 'online', 'platform': platform_name,
                    'capabilities': {'officeAutomation': True}, 'preflight': {'platform': platform_name, 'components': {
                    'word': {'available': True, 'status': 'available'}, 'powerpoint': {'available': True, 'status': 'available'},
                    'excel': {'available': True, 'status': 'available'}},
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
                elif modern_suffix == 'xlsx':
                    modern.write_bytes(make_xlsx_bytes(make_xlsx_object_files()))
                else:
                    converters.build_pptx([{'title': '混合任务现代演示', 'body': ['教学内容']}], modern)
                old_file = processor.create_uploaded_file('legacy.' + legacy_suffix, bytes(legacy))[0]
                new_file = processor.create_uploaded_file(modern.name, modern.read_bytes())[0]
                task = processor.create_task({'task_type': task_type, 'file_ids': [old_file['id'], new_file['id']],
                    'options': {'word': {'generateToc': True, 'maxCharsPerSlide': 80},
                                'ppt': {'extractNotes': False, 'generateToc': False},
                                'excel': {'retainComments': True}}})
                store.update_settings({'pptToWordTemplate': '交接后修改的模板', 'retainComments': False})
                payload = processor.local_task_payload(task['id'])
                office_action = next(action for action in payload['local_actions'] if action['type'] == 'office_conversion')
                pending_settings = office_action['pending_files'][0]['conversion_settings']
                if task_type.startswith('excel_'):
                    self.assertTrue(pending_settings['excel_retain_comments'])
                self.assertTrue(pending_settings['word_generate_toc'])
                self.assertEqual(pending_settings['word_max_chars_per_slide'], 80)
                self.assertFalse(pending_settings['ppt_extract_notes'])
                self.assertFalse(pending_settings['ppt_to_word_generate_toc'])
                self.assertNotEqual(pending_settings['ppt_to_word_template'], '交接后修改的模板')
                report = processor._latest_report_for_task(task['id'])
                existing = next(item for item in report['analysis']['artifacts'] if item['file_id'] == new_file['id'])
                original = Path(existing['path']).read_bytes()
                def normalize(source, target, application, timeout_seconds=120):
                    if application == 'word':
                        converters.build_docx([{'text': '模拟旧文档规范化'}], target)
                    elif application == 'excel':
                        target.write_bytes(make_xlsx_bytes(make_xlsx_object_files()))
                    else:
                        converters.build_pptx([{'title': '模拟旧演示规范化', 'body': ['教学内容']}], target)
                    return {'native_execution_performed': True}
                normalizer_name = 'normalize_windows_office_document' if platform_name == 'Windows' else 'normalize_macos_office_document'
                runner = local_client.execute_windows_office_task if platform_name == 'Windows' else local_client.execute_macos_office_task
                with patch.object(local_client.platform, 'system', return_value=platform_name), patch.object(local_client, normalizer_name, side_effect=normalize) as normalizer:
                    execution = runner(payload, True)
                self.assertEqual(normalizer.call_count, 1)
                sync_payload = local_client.build_native_report_sync_payload(payload, execution, platform_name=platform_name, client_id='mixed-regression')
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
                        if task_type.startswith('excel_'):
                            self.assertIn('图表部件缺失或格式无效', artifact['message'])
                        if task_type.startswith('excel_'):
                            if artifact['output_type'] == 'pdf':
                                path = root / 'download-check.pdf'
                                path.write_bytes(data)
                                text = ''.join(item['text'] for item in converters.extract_pdf_text_blocks(path))
                            else:
                                from io import BytesIO
                                with zipfile.ZipFile(BytesIO(data)) as archive:
                                    text = ''.join(archive.read(name).decode() for name in archive.namelist()
                                                   if name.endswith('.xml') and (name.startswith('word/') or name.startswith('ppt/slides/')))
                            self.assertIn('复核', text)

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
