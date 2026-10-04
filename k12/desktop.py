"""提供 macOS 与 Windows 桌面包的统一入口。"""

import errno
import os
import subprocess
import sys
from pathlib import Path

from . import local_client, server


def data_directory() -> Path:
    """将桌面数据放在用户可写目录，避免写入安装包。"""
    if sys.platform == "win32":
        root = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return root / "K12"


def show_startup_error(message: str) -> None:
    """通过平台对话框显示错误，失败时保留终端提示。"""
    print(message, file=sys.stderr)
    if sys.platform == "darwin":
        script = 'on run argv\ndisplay alert "K12 启动失败" message (item 1 of argv) as critical\nend run'
        try:
            subprocess.run(["/usr/bin/osascript", "-e", script, "--", message], timeout=30, check=False)
        except (OSError, subprocess.TimeoutExpired):
            # 系统对话框不可用时，保留已输出的终端错误。
            pass
    elif sys.platform == "win32":
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, "K12 启动失败", 0x10)


def main() -> None:
    """双击打开工作台，也支持调用打包后的本地客户端。"""
    if "--local-client" in sys.argv[1:]:
        sys.argv.remove("--local-client")
        local_client.main()
        return
    if not any(arg == "--data-dir" or arg.startswith("--data-dir=") for arg in sys.argv[1:]):
        sys.argv.extend(["--data-dir", str(data_directory())])
    background = "--no-open-browser" in sys.argv
    if background:
        sys.argv.remove("--no-open-browser")
    elif "--open-browser" not in sys.argv:
        sys.argv.append("--open-browser")
    try:
        server.main()
    except OSError as exc:
        if exc.errno == errno.EADDRINUSE:
            message = "端口已被占用。请关闭已运行的 K12，或使用 --port 指定其他端口。"
        elif exc.errno in {errno.EACCES, errno.EPERM}:
            message = "无法访问数据目录或监听端口。请检查权限，或使用 --data-dir 指定可写目录。"
        elif exc.errno == errno.ENOSPC:
            message = "磁盘空间不足，请清理空间后重新启动。"
        else:
            message = "本地服务启动失败，请检查数据目录与系统资源后重试。"
        if background:
            print(message, file=sys.stderr)
        else:
            show_startup_error(message)
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
