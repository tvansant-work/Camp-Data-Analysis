#!/bin/bash

# Configuration
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Data-Analysis"
APP_NAME="Camp_Analysis"
SHORTCUT="$HOME/Desktop/Camp Analysis.command"

# 1. Setup Folders
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
mkdir -p "$BASE_DIR"
cd "$BASE_DIR"

echo "Step 1/3: Preparing Python environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

# 2. Create the "Launch & Update" Wrapper
# We use 'EOF' (with quotes) to prevent the local shell from evaluating variables early
cat << 'EOF' > launcher.sh
#!/bin/bash
# Re-establish variables inside the launcher context
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Data-Analysis"
APP_NAME="Camp_Analysis"
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
SHORTCUT="$HOME/Desktop/Camp Analysis.command"

cd "$BASE_DIR"

echo "=================================================="
echo " 🔄 Checking for updates from GitHub..."
echo "=================================================="

# A. Force Update Code & Requirements (Cache-Busting applied)
curl -s -L -o camp_report_app.py "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/camp_report_app.py?t=$(date +%s)"
curl -s -L -o requirements.txt "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/requirements.txt?t=$(date +%s)"

# B. Sync Dependencies
source venv/bin/activate
pip install -r requirements.txt --quiet

# C. Force Update Icon
echo " 🖼️  Updating App Icon..."
curl -s -L -o app_icon.png "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png?t=$(date +%s)"

# D. Re-apply Icon to Desktop Shortcut (Ensures icon updates instantly)
./venv/bin/python3 - << 'PYEOF'
import Cocoa
import os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Analysis/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Analysis.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    icon_image = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if icon_image:
        Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(icon_image, file_path, 0)
PYEOF

echo " ✅ Everything is up to date. Launching..."
echo "=================================================="

# E. RUN THE APP
python3 camp_report_app.py
EOF

chmod +x launcher.sh

# 3. Finalize Desktop Shortcut
echo "Step 2/3: Creating Desktop Shortcut..."
echo "\"$BASE_DIR/launcher.sh\"" > "$SHORTCUT"
chmod +x "$SHORTCUT"

echo "Step 3/3: Applying Initial Icon..."
curl -s -L -o app_icon.png "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png?t=$(date +%s)"
./venv/bin/pip install pyobjc-framework-Cocoa --quiet

./venv/bin/python3 - << 'PYEOF'
import Cocoa
import os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Analysis/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Analysis.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    img = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if img: Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(img, file_path, 0)
PYEOF

echo "------------------------------------------------"
echo "INSTALLATION COMPLETE!"
echo "The app will now check for ALL updates (including icons) every time it starts."