#!/bin/bash
# 固定到项目目录，双击启动时也能正确找到资源。
cd "$(dirname "$0")" || exit 1
export PYTHONUTF8=1

# 优先使用项目解释器，再检查常见 Python 安装位置。
for k12_python in "$PWD/.venv/bin/python" /opt/homebrew/bin/python3 /usr/local/bin/python3 "$(command -v python3)"; do
    if [ -x "$k12_python" ] && "$k12_python" -c 'import sys; raise SystemExit(sys.version_info < (3, 11))' >/dev/null 2>&1; then
        "$k12_python" -m k12 --open-browser "$@"
        k12_status=$?
        if [ "$k12_status" -ne 0 ]; then
            read -r -p '启动失败，按回车关闭窗口。' k12_reply
        fi
        exit "$k12_status"
    fi
done
printf '%s\n' '未找到 Python 3.11 或更高版本，请安装后重新启动。'
read -r -p '按回车关闭窗口。' k12_reply
exit 1
