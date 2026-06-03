ST. ANNE MISSION HOSPITAL - ICT COMMAND CENTRE
================================================

SIMPLE START
------------
1. Double-click start.bat.
2. The app opens in your normal web browser.
3. Keep the black command window open while using the app.
4. Close the command window when finished.

This version avoids the old desktop wrapper dependency. It runs a small local
server on this computer only and stores the data in ICT_MASTER.xlsx.


WHAT YOU NEED
-------------
- Python 3.10 or newer
- openpyxl Python package

start.bat checks for openpyxl and installs it only if it is missing.


YOUR DATA
---------
ICT_MASTER.xlsx is the main data file. Keep it safe.

Important files:
- ICT_MASTER.xlsx: inventory, history, replacements, ink, audit log
- auth.json: PIN data
- config.json: custom departments, equipment types, and conditions
- backups: automatic workbook backups


SECURITY
--------
On first launch you will set a PIN. To reset the PIN, close the app, delete
auth.json, and start again.

Do not share auth.json if you want to keep access restricted.


COMMON PROBLEMS
---------------
Python not found:
Install Python from https://www.python.org/downloads/ and tick "Add Python to PATH".

Cannot save or permission denied:
Close ICT_MASTER.xlsx in Excel, then try again.

Browser does not open:
Look in the black command window for the local address, then paste it into
Chrome, Edge, or Firefox.

App crashes:
Open Command Prompt in this folder and run:
python main.py


OPTIONAL DESKTOP MODE
---------------------
The old pywebview desktop wrapper is still available for machines that already
have it installed:
python main.py --desktop
