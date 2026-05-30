ST. ANNE MISSION HOSPITAL — ICT COMMAND CENTRE
===============================================
Desktop Application — Your Excel IS the database

HOW IT WORKS
------------
Every action you take (add, edit, delete, replace, ink update)
saves IMMEDIATELY and PERMANENTLY to ICT_MASTER.xlsx.

No internet needed. No database server. No login.
The Excel file IS the system. Copy it to USB — take it anywhere.


FOLDER CONTENTS
---------------
app.py               — The application (do not move this)
ICT_MASTER.xlsx      — YOUR MASTER DATA FILE (all assets live here)
INK_MASTER.xlsx      — Original ink reference
RUN_ON_WINDOWS.bat   — Double-click this to launch on Windows
RUN_ON_MAC.sh        — Double-click this to launch on Mac
backups/             — Auto-created. A backup is made before every change.
README.txt           — This file


FIRST TIME SETUP (Windows)
--------------------------
1. Make sure Python is installed
   - Open Command Prompt, type: python --version
   - If not found: go to https://www.python.org/downloads/
   - Download Python 3.11 or newer
   - During install: CHECK the box "Add Python to PATH"

2. Double-click RUN_ON_WINDOWS.bat
   - It installs openpyxl automatically
   - The app opens

3. Done. From now on just double-click RUN_ON_WINDOWS.bat


FIRST TIME SETUP (Mac)
-----------------------
1. Open Terminal
2. Type: xcode-select --install   (if prompted)
3. Type: pip3 install openpyxl
4. Double-click RUN_ON_MAC.sh
   OR in Terminal: cd to this folder, then type: python3 app.py


WHAT SAVES WHERE IN ICT_MASTER.xlsx
-------------------------------------
Sheet: Inventory      — All active assets. New items added here in order.
Sheet: History        — Every repair, upgrade, reassignment logged here.
Sheet: Replacements   — Full replacement records with dates and notes.
Sheet: Ink            — Every ink use/restock with timestamp.
Sheet: Audit Log      — Every single action with timestamp.

You can open ICT_MASTER.xlsx in Excel ANY TIME and see live data.
You can also click "Open Excel" in the sidebar to open it directly.


BACKUPS
-------
Before EVERY change (add, edit, delete, replace) the system
automatically saves a backup to the backups/ folder named:
  ICT_MASTER_backup_YYYYMMDD_HHMMSS.xlsx

If anything ever goes wrong, go to backups/ and restore the latest one.


PORTABLE USE
------------
To use on another computer:
1. Copy the ENTIRE folder (app.py + ICT_MASTER.xlsx + .bat file)
2. Put it on a USB stick or shared folder
3. Run RUN_ON_WINDOWS.bat on the other computer
4. All your data is there — it's all in ICT_MASTER.xlsx


SHARING / MULTIPLE USERS
--------------------------
This app is designed for ONE user at a time editing the Excel.
If two people open it simultaneously they should NOT both edit —
last save wins and the other's changes will be lost.

For multi-user access (everyone edits at the same time from
different computers), upgrade to the Supabase cloud version.
Ask your ICT developer to set that up.


ADDING A NEW ASSET
------------------
- Click "Add Asset" button or "➕ Add" in the toolbar
- Fill in the form
- Click "Save to Excel"
- The asset appears in the table AND in ICT_MASTER.xlsx immediately
- Open the Excel file — your new asset is in the Inventory sheet


RECORDING A REPLACEMENT
------------------------
1. Select the asset in the Inventory tab
2. Click "🔄 Replace" button
3. Fill in: date, new item details, replacement note
4. Click "Confirm Replacement"

What happens:
- Old asset → condition set to "Replaced", note added to remarks
- New asset → added as a new row in Inventory
- Replacements sheet → full record with date, old item, new item, note
- History sheet → event logged
- Audit Log → logged with timestamp


LOGGING A REPAIR / UPGRADE / EVENT
-------------------------------------
1. Select the asset
2. Click "📝 Log Event"
3. Choose event type (repaired, upgraded, reassigned, serviced...)
4. Write the full description
5. Save

The event is stored permanently in the History sheet.
If event is "repaired" and condition was "Needs Repair" → auto-set to "Good"


INK STATION
-----------
- Use the − button to log ink usage (deducts from store count)
- Use the ＋ button to restock
- Set the spinner to the quantity before pressing +/−
- Every action saves to the Ink sheet with timestamp
- Big ICT / Small ICT installed counts are editable


SEARCH
------
Type anything in the search box — it searches across ALL fields:
type, brand, serial number, department, assigned staff, IP address,
MAC address, OS, CPU, RAM, disk, remarks, date.


TROUBLESHOOTING
---------------
Problem: "Python not found"
Fix: Install Python from python.org, check "Add to PATH"

Problem: App opens but shows no data
Fix: Make sure ICT_MASTER.xlsx is in the SAME folder as app.py

Problem: "Permission denied" when saving
Fix: Close ICT_MASTER.xlsx in Excel first, then retry

Problem: App crashes on start
Fix: Open Command Prompt in the folder, run: python app.py
     Read the error message and contact your ICT support


SUPPORT
-------
Managed by: ICT Department, St. Anne Mission Hospital
Master file: ICT_MASTER.xlsx (keep this file safe — it is everything)
