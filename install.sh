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
cat << 'EOF' > launcher.sh
#!/bin/bash
GITHUB_USER="tvansant-work"
GITHUB_REPO="Camp-Data-Analysis"
APP_NAME="Camp_Analysis"
BASE_DIR="$HOME/Library/Application Support/$APP_NAME"
SHORTCUT="$HOME/Desktop/Camp Analysis.command"
BRANCH="main"

cd "$BASE_DIR"

echo "=================================================="
echo " Checking for updates from GitHub..."
echo "=================================================="

# A. Use the GitHub API to get the exact latest commit SHA.
#    Fetching by SHA bypasses GitHub's CDN cache completely — the
#    ?t= trick does NOT reliably bust GitHub's raw content cache.
API_BASE="https://api.github.com/repos/$GITHUB_USER/$GITHUB_REPO"
echo " Fetching latest commit SHA..."
LATEST_SHA=$(curl -s -f \
  -H "Accept: application/vnd.github+json" \
  "$API_BASE/commits/$BRANCH" \
  | python3 -c "import sys,json; print(json.load(sys.stdin)['sha'])" 2>/dev/null)

if [ -n "$LATEST_SHA" ]; then
  echo " Latest commit: ${LATEST_SHA:0:7}"
  RAW_BASE="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$LATEST_SHA"
else
  echo " WARNING: GitHub API unavailable, falling back to branch URL"
  RAW_BASE="https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/$BRANCH"
fi

# B. Download files at the exact commit SHA (cache-proof)
echo " Downloading camp_report_app.py..."
curl -s -L -H "Cache-Control: no-cache" -H "Pragma: no-cache" \
  -o camp_report_app.py "$RAW_BASE/camp_report_app.py"

curl -s -L -H "Cache-Control: no-cache" -H "Pragma: no-cache" \
  -o requirements.txt "$RAW_BASE/requirements.txt"

# C. Print file info so you can confirm it's the right version
APP_BYTES=$(wc -c < camp_report_app.py | tr -d ' ')
APP_PREVIEW=$(sed -n '3p' camp_report_app.py)
echo " camp_report_app.py: ${APP_BYTES} bytes"
echo " Preview line 3:     $APP_PREVIEW"

# D. Sync Dependencies
source venv/bin/activate
pip install -r requirements.txt --quiet

# E. Update Icon
echo " Updating App Icon..."
curl -s -L -H "Cache-Control: no-cache" -H "Pragma: no-cache" \
  -o app_icon.png "$RAW_BASE/app_icon.png"

# F. Re-apply Icon to Desktop Shortcut
./venv/bin/python3 - << 'PYEOF'
import Cocoa, os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Analysis/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Analysis.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    icon_image = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if icon_image:
        Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(icon_image, file_path, 0)
PYEOF

echo " All up to date. Launching..."
echo "=================================================="

# G. RUN THE APP
python3 camp_report_app.py
EOF

chmod +x launcher.sh

# 3. Finalize Desktop Shortcut
echo "Step 2/3: Creating Desktop Shortcut..."
echo "\"$BASE_DIR/launcher.sh\"" > "$SHORTCUT"
chmod +x "$SHORTCUT"

echo "Step 3/3: Applying Initial Icon..."
curl -s -L -o app_icon.png "https://raw.githubusercontent.com/$GITHUB_USER/$GITHUB_REPO/main/app_icon.png"
./venv/bin/pip install pyobjc-framework-Cocoa --quiet

./venv/bin/python3 - << 'PYEOF'
import Cocoa, os
icon_path = os.path.expanduser("~/Library/Application Support/Camp_Analysis/app_icon.png")
file_path = os.path.expanduser("~/Desktop/Camp Analysis.command")
if os.path.exists(icon_path) and os.path.exists(file_path):
    img = Cocoa.NSImage.alloc().initWithContentsOfFile_(icon_path)
    if img: Cocoa.NSWorkspace.sharedWorkspace().setIcon_forFile_options_(img, file_path, 0)
PYEOF

echo "------------------------------------------------"
echo "INSTALLATION COMPLETE!"
echo "The app will now verify its version from GitHub every time it starts."