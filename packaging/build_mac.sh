#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WORK_DIR="$PROJECT_ROOT/.tmp/installer-mac"
APP_STAGE_DIR="$WORK_DIR/app"
CERT_STAGE_DIR="$WORK_DIR/certificate"
PYINSTALLER_DIR="$WORK_DIR/pyinstaller"
PKG_SCRIPTS_DIR="$WORK_DIR/pkg-scripts"
OUTPUT_DIR="$PROJECT_ROOT/dist"
MANIFEST_PATH="$PROJECT_ROOT/apps/excel-addin/manifest.xml"

APP_VERSION="$(python3 -c 'import json; print(json.load(open("'"$PROJECT_ROOT"'/package.json"))["version"])')"
ARCH="$(uname -m)"
INSTALLER_NAME="Excel-Bro-Setup-$APP_VERSION-$ARCH.pkg"
INSTALLER_PATH="$OUTPUT_DIR/$INSTALLER_NAME"

python3 - "$MANIFEST_PATH" <<'PY'
import sys
import xml.etree.ElementTree as ET

manifest = ET.parse(sys.argv[1]).getroot()
version = [int(part) for part in manifest.findtext("Version").split(".")]
if version < [1, 0, 0, 0]:
    raise SystemExit(f"Office 加载项 manifest 版本不能低于 1.0.0.0：{manifest.findtext('Version')}")
for tag in ("IconUrl", "HighResolutionIconUrl"):
    node = manifest.find(tag)
    if node is None or not (node.get("DefaultValue") or "").strip():
        raise SystemExit("Office 加载项 manifest 必须包含 IconUrl 和 HighResolutionIconUrl。")
PY

rm -rf "$WORK_DIR"
mkdir -p "$APP_STAGE_DIR" "$CERT_STAGE_DIR" "$PKG_SCRIPTS_DIR" "$OUTPUT_DIR"

cd "$PROJECT_ROOT"
npm ci
npm run build:addin
python3 packaging/generate_certificate.py "$CERT_STAGE_DIR"
python3 -m PyInstaller \
    --noconfirm \
    --clean \
    --onedir \
    --name ExcelBro \
    --distpath "$PYINSTALLER_DIR/dist" \
    --workpath "$PYINSTALLER_DIR/work" \
    --specpath "$PYINSTALLER_DIR/spec" \
    --add-data "$PROJECT_ROOT/server:server" \
    --add-data "$PROJECT_ROOT/config:config" \
    --hidden-import tkinter \
    --hidden-import tkinter.filedialog \
    --exclude-module cv2 \
    --exclude-module IPython \
    --exclude-module matplotlib \
    --exclude-module numba \
    --exclude-module pyarrow \
    --exclude-module PyQt5 \
    --exclude-module pytest \
    --exclude-module scipy \
    --exclude-module torch \
    --exclude-module yt_dlp \
    packaging/runtime.py

INSTALL_ROOT="$APP_STAGE_DIR/Applications/Excel Bro"
mkdir -p "$INSTALL_ROOT/frontend" "$INSTALL_ROOT/catalog"
cp -R "$PYINSTALLER_DIR/dist/ExcelBro/"* "$INSTALL_ROOT/"
cp -R "$PROJECT_ROOT/apps/excel-addin/dist/"* "$INSTALL_ROOT/frontend/"
cp "$MANIFEST_PATH" "$INSTALL_ROOT/manifest.xml"
cp "$MANIFEST_PATH" "$INSTALL_ROOT/catalog/manifest.xml"
cp "$CERT_STAGE_DIR"/localhost.* "$INSTALL_ROOT/"
cp "$PROJECT_ROOT/packaging/mac/uninstall.sh" "$INSTALL_ROOT/uninstall.sh"
chmod +x "$INSTALL_ROOT/ExcelBro" "$INSTALL_ROOT/uninstall.sh"

cp "$PROJECT_ROOT/packaging/mac/pkg-scripts/"* "$PKG_SCRIPTS_DIR/"
chmod +x "$PKG_SCRIPTS_DIR/"*

rm -f "$INSTALLER_PATH"
pkgbuild \
    --root "$APP_STAGE_DIR" \
    --scripts "$PKG_SCRIPTS_DIR" \
    --identifier com.excelbro.app \
    --version "$APP_VERSION" \
    --install-location / \
    "$WORK_DIR/excel-bro-component.pkg"
productbuild \
    --package "$WORK_DIR/excel-bro-component.pkg" \
    "$INSTALLER_PATH"

echo "已生成：$INSTALLER_PATH"
