"""在目标系统构建独立桌面版本，包含工作台与 Office 脚本。"""

import subprocess
import sys
from pathlib import Path


def main() -> None:
    """由 PyInstaller 收集 Python 运行时与项目资源。"""
    root = Path(__file__).resolve().parent.parent
    command = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir",
               "--name", "K12", "--paths", str(root),
               "--distpath", str(root / "dist"), "--workpath", str(root / "build"),
               "--specpath", str(root / "build"),
               "--add-data", f"{root / 'static'}:static",
               "--add-data", f"{root / 'k12/pptx_template.pptx'}:k12",
               "--add-data", f"{root / 'k12/windows_office.ps1'}:k12",
               "--add-data", f"{root / 'k12/windows_office_lifecycle.ps1'}:k12"]
    if sys.platform == "darwin":
        command += ["--windowed", "--osx-bundle-identifier", "com.k12.workbench"]
    elif sys.platform != "win32":
        raise SystemExit("请在 macOS 或 Windows 上构建对应桌面版本。")
    subprocess.run([*command, str(root / "tools/desktop_entry.py")], cwd=root, check=True)
    # macOS 签名由交付工具在临时目录执行，避免文件提供器附加元数据。
    print("构建完成，请运行 python tools/package_desktop.py 验证交付包。")


if __name__ == "__main__":
    main()
