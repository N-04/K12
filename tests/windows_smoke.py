"""在 Windows 上用真实 Office 验证上传、转换、同步和产物下载。"""
import hashlib
import json
import pathlib
import secrets
import subprocess
import tempfile
import threading
import zipfile
import io
import argparse
from urllib.request import Request, urlopen

from k12.local_client import build_heartbeat, request_json, run_once, run_windows_office_script
from k12.server import K12RequestHandler, K12Server
from k12.converters import build_docx, build_pptx


def main(web_only=False):
    """创建独立测试数据和 Office 样本，结束时关闭服务与清理临时目录。"""
    results = []
    with tempfile.TemporaryDirectory(prefix='k12-windows-native-') as tmp:
        root = pathlib.Path(tmp)
        if web_only:
            build_docx([{'text': 'Windows 中文课件测试'}], root / 'legacy.docx')
            build_pptx([{'title': 'Windows 中文课件测试', 'body': ['教学内容']}], root / 'legacy.pptx')
        else:
            result = run_windows_office_script(pathlib.Path(__file__).with_name('windows_fixtures.ps1'), ['-FixtureDirectory', str(root)])
            if result.returncode:
                raise RuntimeError('真实 Office 测试样本生成失败，请检查 Office 初始化和授权')
        server = K12Server(('127.0.0.1', 0), K12RequestHandler, root / 'data')
        token = secrets.token_urlsafe(24)
        server.store.update_settings({'localSecurityToken': token, 'localClientPlatform': 'Windows'})
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        origin = f'http://127.0.0.1:{server.server_port}'
        headers = {'X-K12-Token': token}
        try:
            with urlopen(origin + '/', timeout=10) as response:
                assert response.status == 200 and b'app.js' in response.read()
            request_json(origin, '/api/local-client/heartbeat', method='POST', token=token, payload=build_heartbeat())
            for extension, task_type, required_part in [('doc', 'word_to_ppt', 'ppt/presentation.xml'), ('ppt', 'ppt_to_word', 'word/document.xml')]:
                if web_only:
                    extension += 'x'
                boundary = 'k12-smoke-boundary'
                content = (root / ('legacy.' + extension)).read_bytes()
                body = (f'--{boundary}\r\nContent-Disposition: form-data; name="files"; filename="legacy.{extension}"\r\nContent-Type: application/octet-stream\r\n\r\n'.encode() + content + f'\r\n--{boundary}--\r\n'.encode())
                request = Request(origin + '/api/uploads', data=body, headers={**headers, 'Content-Type': 'multipart/form-data; boundary=' + boundary})
                with urlopen(request, timeout=30) as response:
                    file = json.load(response)['files'][0]
                task = request_json(origin, '/api/tasks', method='POST', token=token, payload={'task_type': task_type, 'file_ids': [file['id']]})['task']
                if web_only:
                    assert task['status'] == '成功', task
                else:
                    assert task['status'] == '待处理', task
                    execution = run_once(origin, token=token, task_id=task['id'], allow_native_execution=True, execute_native_office=True, timeout=30)
                    assert execution['native_office_execution']['status'] == 'success', execution
                    assert execution['sync']['task']['status'] == '成功', execution
                report = server.processor._latest_report_for_task(task['id'])
                artifact = report['analysis']['artifacts'][0]
                with urlopen(Request(origin + '/api/tasks/' + task['id'] + '/download', headers=headers), timeout=30) as response:
                    bundle = response.read()
                with zipfile.ZipFile(io.BytesIO(bundle)) as archive:
                    assert archive.testzip() is None
                with urlopen(Request(origin + artifact['url'], headers=headers), timeout=30) as response:
                    output_data = response.read()
                assert hashlib.sha256(output_data).hexdigest() == artifact['sha256']
                with zipfile.ZipFile(io.BytesIO(output_data)) as archive:
                    assert required_part in archive.namelist() and archive.testzip() is None
                report_data = request_json(origin, '/api/reports/' + report['id'] + '/download?format=json', token=token)
                assert report_data['id'] == report['id']
                results.append({'task_type': task_type, 'status': '成功', 'bundle_bytes': len(bundle), 'bundle_sha256': hashlib.sha256(bundle).hexdigest()})
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)
    print(json.dumps({'windows_web_smoke' if web_only else 'windows_native_smoke': results}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='验证 Windows 网页功能或真实 Office 转换')
    parser.add_argument('--web-only', action='store_true', help='仅测试标准库网页功能，不启动 Office')
    main(parser.parse_args().web_only)
