#!/usr/bin/env bash
# Fill Form backend packaging script for macOS / Linux.
# Mirrors build_backend.ps1 so CI and local macOS builds behave the same way.
# PaddleOCR runtime is bundled, but downloaded OCR models are not.
#
# Usage:
#   bash build_backend.sh

set -euo pipefail

script_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
project_root="$(dirname "$script_root")"

# 优先复用已有的虚拟环境，否则使用 PATH 上的 python3。
if [[ -x "$project_root/.venv/bin/python" ]]; then
    python_bin="$project_root/.venv/bin/python"
elif [[ -x "$script_root/.venv/bin/python" ]]; then
    python_bin="$script_root/.venv/bin/python"
else
    python_bin="$(command -v python3 || command -v python)"
fi

echo "Using Python: $python_bin"
"$python_bin" --version

echo "Checking PyInstaller..."
# PyInstaller 探测缺失模块时会返回非零退出码，这里用 if 包住避免被 set -e 中断。
if ! "$python_bin" -m PyInstaller --version >/dev/null 2>&1; then
    echo "Installing PyInstaller..."
    "$python_bin" -m pip install PyInstaller
fi

cd "$script_root"
echo "Building backend. Downloaded OCR models are not included..."
# 注意：macOS 上必须使用 --noupx，UPX 会压坏 Mach-O 产物。
"$python_bin" -m PyInstaller \
    --noconfirm \
    --clean \
    --onedir \
    --noupx \
    --name fill_form_backend \
    --collect-all paddle \
    --collect-all paddleocr \
    --collect-all paddlex \
    --collect-all fitz \
    --collect-all aistudio_sdk \
    --collect-all chardet \
    --collect-all colorlog \
    --collect-all filelock \
    --collect-all huggingface_hub \
    --collect-all modelscope \
    --collect-all pandas \
    --collect-all prettytable \
    --collect-all cpuinfo \
    --collect-all pydantic \
    --collect-all yaml \
    --collect-all requests \
    --collect-all ruamel \
    --collect-all typing_extensions \
    --collect-all ujson \
    --collect-all imagesize \
    --collect-all cv2 \
    --collect-all pyclipper \
    --collect-all pypdfium2 \
    --collect-all bidi \
    --collect-all shapely \
    --hidden-import uvicorn.logging \
    --hidden-import uvicorn.loops.auto \
    --hidden-import uvicorn.protocols.http.auto \
    --hidden-import uvicorn.protocols.websockets.auto \
    main.py

echo ""
echo "Packaging complete. Backend directory:"
echo "  $script_root/dist/fill_form_backend"
echo "This directory will be included as an Electron application resource."
