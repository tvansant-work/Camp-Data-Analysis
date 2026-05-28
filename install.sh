#!/bin/bash

# Configuration - REPLACE THESE
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Data-Analysis"
APP_NAME="Camp_Analysis"

# 1. Setup Folders
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
mkdir -p "$BASE_DIR"
cd "$BASE_DIR"

echo "Step 1/3: Creating secure Python environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

# 2. Create the "Launch & Update" Wrapper
# This script runs EVERY TIME they click the icon
cat << EOF > launcher.sh
#!/bin/bash
cd "$BASE_DIR"

# A. Update Code & Requirements from GitHub
curl -s -o camp_report_app.py https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/camp_report_app.py
curl -s -o requirements.txt https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/requirements.txt

# B. Sync Dependencies (Only installs if requirements.txt changed)
source venv/bin/activate
pip install -r requirements.txt --quiet

# C. Update Icon if missing
if [ ! -f "app_icon.png" ]; then
    curl -s -o app_icon.png https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png
fi

# D. RUN THE APP
python3 camp_report_app.py
EOF

chmod +x launcher.sh

echo "Step 2/3: Creating Desktop Shortcut..."
SHORTCUT="$HOME/Desktop/Camp Analysis.command"
echo "\"$BASE_DIR/launcher.sh\"" > "$SHORTCUT"
chmod +x "$SHORTCUT"

echo "Step 3/3: Applying Custom Icon..."
# Download icon for the setter
curl -s -o app_icon.png https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png

# Advanced Python-Cocoa bridge to set the icon programmatically
/usr/bin/python3 - << 'PYEOF'
import Cocoa
import os
import sys

icon_path = os.path.expanduser("~/Library/Application Support/Camp_Analysis/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Analysis.command")

if os.path.exists(icon_path) and os.path.exists(file_path):
    icon_image = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if icon_image:
        Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(icon_image, file_path, 0)
PYEOF

echo "------------------------------------------------"
echo "INSTALLATION COMPLETE!"
echo "You can now close this window and use the app on your Desktop."