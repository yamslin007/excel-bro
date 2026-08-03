#!/bin/bash
# 卸载 Excel Bro（需要 sudo：删除系统钥匙串证书和 /Applications 下的程序）。
# 不会删除 ~/Library/Application Support/Excel Bro 中的模型配置。
set -euo pipefail

if [ "$(id -u)" -ne 0 ]; then
    echo "请使用 sudo 运行：sudo bash uninstall.sh" >&2
    exit 1
fi

INSTALL_DIR="/Applications/Excel Bro"
PLIST_LABEL="com.excelbro.runtime"
CERT_SUBJECT="Excel Bro localhost"
SYSTEM_KEYCHAIN="/Library/Keychains/System.keychain"

CONSOLE_USER="$(stat -f %Su /dev/console 2>/dev/null || true)"
if [ -n "$CONSOLE_USER" ] && [ "$CONSOLE_USER" != "root" ]; then
    CONSOLE_UID="$(id -u "$CONSOLE_USER")"
    launchctl bootout "gui/$CONSOLE_UID/$PLIST_LABEL" 2>/dev/null || true
    rm -f "/Users/$CONSOLE_USER/Library/LaunchAgents/$PLIST_LABEL.plist"
    rm -f "/Users/$CONSOLE_USER/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml"
fi

pkill -x ExcelBro 2>/dev/null || true

while IFS= read -r hash; do
    [ -n "$hash" ] && security delete-certificate -Z "$hash" "$SYSTEM_KEYCHAIN" || true
done < <(security find-certificate -c "$CERT_SUBJECT" -a -Z "$SYSTEM_KEYCHAIN" 2>/dev/null |
    sed -n 's/^SHA-256 hash: //p')

rm -rf "$INSTALL_DIR"
echo "Excel Bro 已卸载。模型配置保留在 ~/Library/Application Support/Excel Bro。"
