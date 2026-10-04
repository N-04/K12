"""生成并验证双平台交付压缩包，打印 SHA-256。"""

import hashlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def main() -> None:
    """在临时目录打包，验证实际压缩包解压后的应用。"""
    root = Path(__file__).resolve().parent.parent
    if sys.platform not in {"darwin", "win32"}:
        raise SystemExit("请在 macOS 或 Windows 上打包对应桌面版本。")
    label = "macOS" if sys.platform == "darwin" else "Windows"
    target = root / f"dist/K12-{label}.zip"
    with tempfile.TemporaryDirectory(prefix="k12-release-") as temporary:
        stage = Path(temporary)
        archive = stage / target.name
        unpacked = stage / "verified"
        if sys.platform == "darwin":
            bundle = stage / "K12.app"
            # 避开工作目录文件提供器重新附加的 Finder 元数据。
            subprocess.run(["ditto", str(root / "dist/K12.app"), str(bundle)], check=True)
            subprocess.run(["xattr", "-cr", str(bundle)], check=True)
            subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(bundle)], check=True)
            subprocess.run(["codesign", "--verify", "--deep", "--strict", str(bundle)], check=True)
            subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(bundle), str(archive)], check=True)
            subprocess.run(["ditto", "-x", "-k", str(archive), str(unpacked)], check=True)
            subprocess.run(["codesign", "--verify", "--deep", "--strict", str(unpacked / "K12.app")], check=True)
            executable = unpacked / "K12.app/Contents/MacOS/K12"
        else:
            shutil.make_archive(str(archive.with_suffix("")), "zip", root / "dist", "K12")
            shutil.unpack_archive(archive, unpacked)
            executable = unpacked / "K12/K12.exe"
        subprocess.run([sys.executable, str(root / "tools/smoke_desktop.py"), str(executable)], cwd=root, check=True)
        # 仅在交付包验证成功后替换已有产物。
        shutil.copyfile(archive, target)
    with target.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256").hexdigest()
    print(f"{target}\nSHA-256：{digest}")


if __name__ == "__main__":
    main()
