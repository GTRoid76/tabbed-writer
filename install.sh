#!/usr/bin/env bash
# Installs Tabbed Writer for the current user on Linux.
# It adds the program to your Menu, gives it an icon, and makes
# .twp files open in Tabbed Writer when you double-click them.
#
# Run it from this folder with:   bash install.sh

set -e

HERE="$(cd "$(dirname "$0")" && pwd)"
APP_DIR="$HOME/.local/share/tabbed-writer"
ICON_DIR="$HOME/.local/share/icons/hicolor/scalable/apps"
MIME_DIR="$HOME/.local/share/mime"
MENU_DIR="$HOME/.local/share/applications"

echo "Installing Tabbed Writer to $APP_DIR"
mkdir -p "$APP_DIR" "$ICON_DIR" "$MIME_DIR/packages" "$MENU_DIR"
cp "$HERE/main.py" "$HERE/tabbed-writer.svg" "$APP_DIR/"

echo "Setting up Python and the Qt toolkit (this download can take a few minutes)..."
if ! python3 -m venv "$APP_DIR/.venv"; then
    echo
    echo "Python couldn't create its private folder. On Mint, Ubuntu or Debian, run:"
    echo "    sudo apt install python3-venv"
    echo "and then run this installer again."
    exit 1
fi
"$APP_DIR/.venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/.venv/bin/pip" install --quiet -r "$HERE/requirements.txt"

echo "Adding the icon, the Menu entry and the .twp file type..."
cp "$HERE/tabbed-writer.svg" "$ICON_DIR/"

cat > "$MIME_DIR/packages/tabbed-writer.xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/x-tabbed-writer">
    <comment>Tabbed Writer document</comment>
    <glob pattern="*.twp" weight="60"/>
    <generic-icon name="x-office-document"/>
  </mime-type>
</mime-info>
EOF

cat > "$MENU_DIR/tabbed-writer.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Tabbed Writer
Comment=Write with tabs and chapters
Exec="$APP_DIR/.venv/bin/python" "$APP_DIR/main.py" %f
Icon=tabbed-writer
Terminal=false
Categories=Office;WordProcessor;
StartupWMClass=main.py
MimeType=application/x-tabbed-writer;
EOF

update-mime-database "$MIME_DIR" >/dev/null 2>&1 || true
update-desktop-database "$MENU_DIR" >/dev/null 2>&1 || true
gtk-update-icon-cache -f -t "$HOME/.local/share/icons/hicolor" >/dev/null 2>&1 || true
xdg-mime default tabbed-writer.desktop application/x-tabbed-writer >/dev/null 2>&1 || true

echo
echo "Done! Open Tabbed Writer from your Menu (search for \"Tabbed\")."
echo "Your writing is kept in Documents/Tabbed Writer."
