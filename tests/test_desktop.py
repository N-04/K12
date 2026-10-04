"""验证双平台桌面入口与数据持久化位置。"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

from k12 import desktop


class DesktopTests(unittest.TestCase):
    """检查打包程序不会将运行数据写入安装目录。"""

    def test_platform_data_directories(self):
        """每个平台使用用户可写的固定数据目录。"""
        with patch.object(desktop.Path, 'home', return_value=Path('/user')):
            with patch.object(sys, 'platform', 'darwin'):
                self.assertEqual(desktop.data_directory(), Path('/user/Library/Application Support/K12'))
            with patch.object(sys, 'platform', 'win32'), patch.dict(desktop.os.environ, {'LOCALAPPDATA': '/local'}):
                self.assertEqual(desktop.data_directory(), Path('/local/K12'))

    def test_server_defaults_and_explicit_directory(self):
        """默认目录与浏览器参数只在用户未指定时添加。"""
        for arguments in [[], ['--data-dir', '/custom'], ['--data-dir=/custom']]:
            with self.subTest(arguments=arguments), patch.object(sys, 'argv', ['K12', *arguments]), patch.object(desktop, 'data_directory', return_value=Path('/default')), patch.object(desktop.server, 'main') as run:
                desktop.main()
                run.assert_called_once_with()
                self.assertEqual(sys.argv.count('--open-browser'), 1)
                if not arguments:
                    self.assertEqual(sys.argv[-3:], ['--data-dir', '/default', '--open-browser'])
                else:
                    self.assertNotIn('/default', sys.argv)

    def test_background_start_does_not_open_browser(self):
        """包验证与后台服务可以显式关闭浏览器启动。"""
        with patch.object(sys, 'argv', ['K12', '--no-open-browser']), patch.object(desktop.server, 'main') as run:
            desktop.main()
            run.assert_called_once_with()
            self.assertNotIn('--open-browser', sys.argv)
            self.assertNotIn('--no-open-browser', sys.argv)

    def test_startup_errors_have_chinese_messages_and_nonzero_exit(self):
        """桌面错误可见且返回失败状态，不泄露原始路径。"""
        import errno
        for code, expected in [(errno.EADDRINUSE, '端口已被占用'), (errno.EACCES, '权限'), (errno.ENOSPC, '磁盘空间不足')]:
            with self.subTest(code=code), patch.object(sys, 'argv', ['K12']), patch.object(desktop.server, 'main', side_effect=OSError(code, '/private/input')), patch.object(desktop, 'show_startup_error') as alert:
                with self.assertRaises(SystemExit) as caught:
                    desktop.main()
                self.assertEqual(caught.exception.code, 1)
                self.assertIn(expected, alert.call_args.args[0])
                self.assertNotIn('/private/input', alert.call_args.args[0])

    def test_failed_store_initialization_releases_server_socket(self):
        """数据库初始化失败后仍必须释放已监听的端口。"""
        from k12.server import K12Server, K12RequestHandler
        with patch.object(K12Server, 'server_bind'), patch.object(K12Server, 'server_activate'), patch.object(K12Server, 'server_close', autospec=True, side_effect=K12Server.server_close) as close, patch('k12.server.AppStore', side_effect=OSError('不可写')):
            with self.assertRaises(OSError):
                K12Server(('127.0.0.1', 0), K12RequestHandler, Path('/data'))
            close.assert_called_once()

    def test_client_mode_preserves_arguments(self):
        """客户端模式不附加服务端参数。"""
        with patch.object(sys, 'argv', ['K12', '--local-client', '--help']), patch.object(desktop.local_client, 'main') as client, patch.object(desktop.server, 'main') as server:
            desktop.main()
            self.assertEqual(sys.argv, ['K12', '--help'])
            client.assert_called_once_with()
            server.assert_not_called()
