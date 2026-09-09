#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../.." && pwd)"

APP_NAME="xlsx-a11y"
VERSION="0.1.0"
DIST_DIR="${REPO_ROOT}/dist"
APP_DIR="${DIST_DIR}/AppDir"
OUT_DIR="${DIST_DIR}/linux"

mkdir -p "${OUT_DIR}"
rm -rf "${APP_DIR}"
mkdir -p "${APP_DIR}/usr/bin" \
         "${APP_DIR}/usr/share/applications" \
         "${APP_DIR}/usr/share/icons/hicolor/256x256/apps" \
         "${APP_DIR}/usr/lib"

echo "==> Staging Linux AppDir files..."
cp -r "${DIST_DIR}/xlsx-a11y-gui"/* "${APP_DIR}/usr/bin/"

# Ensure main launcher symlink
ln -sf "xlsx-a11y" "${APP_DIR}/AppRun"
chmod +x "${APP_DIR}/AppRun"

# Copy icons
cp "${REPO_ROOT}/packaging/icons/xlsx-a11y.png" "${APP_DIR}/usr/share/icons/hicolor/256x256/apps/"
cp "${REPO_ROOT}/packaging/icons/xlsx-a11y.png" "${APP_DIR}/"

# Desktop entry
cat > "${APP_DIR}/xlsx-a11y.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Excel Accessibility Auditor
Exec=xlsx-a11y
Icon=xlsx-a11y
Comment=Audit and remediate Microsoft Excel .xlsx workbooks against WCAG 2.1 AA
Categories=Office;Accessibility;Utility;
Terminal=false
EOF

cp "${APP_DIR}/xlsx-a11y.desktop" "${APP_DIR}/usr/share/applications/"

echo "==> Building Linux AppImage..."
if command -v appimagetool >/dev/null 2>&1; then
  export APPIMAGE_EXTRACT_AND_RUN=1
  appimagetool "${APP_DIR}" "${OUT_DIR}/${APP_NAME}-v${VERSION}-x86_64.AppImage"
  echo "==> AppImage built successfully: ${OUT_DIR}/${APP_NAME}-v${VERSION}-x86_64.AppImage"
else
  echo "Warning: appimagetool not found. AppDir staged at: ${APP_DIR}"
fi
