@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys; raise SystemExit(sys.version_info < (3,11))" >nul 2>&1
    if not errorlevel 1 (
        ".venv\Scripts\python.exe" -m k12 --open-browser %*
        goto done
    )
)
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys; raise SystemExit(sys.version_info < (3,11))" >nul 2>&1
    if not errorlevel 1 (
        py -3 -m k12 --open-browser %*
        goto done
    )
)
where python >nul 2>&1
if not errorlevel 1 (
    python -c "import sys; raise SystemExit(sys.version_info < (3,11))" >nul 2>&1
    if not errorlevel 1 (
        python -m k12 --open-browser %*
        goto done
    )
)
echo 未找到可用的 Python 3.11 或更高版本。请安装 Python，并在 PyCharm 中配置项目解释器。
:done
if errorlevel 1 pause
endlocal
