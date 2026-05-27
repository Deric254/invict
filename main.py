"""
St. Anne Mission Hospital — ICT Command Centre
Desktop app: pywebview + HTML UI + Excel as live database
Run: python main.py
"""
# Suppress pywebview Windows accessibility recursion bug (Python 3.13+/EdgeChromium)
import os as _os
_os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")
_os.environ.setdefault("WEBKIT_DISABLE_COMPOSITING_MODE", "1")
# Prevent Windows UI Automation from recursing into the webview frame
_os.environ.setdefault("PYWEBVIEW_NO_UIAUTOMATION", "1")

import webview, json, os, sys, datetime, shutil, openpyxl, threading, re, hashlib, secrets
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MASTER_XL  = os.path.join(BASE_DIR, "ICT_MASTER.xlsx")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
AUTH_FILE  = os.path.join(BASE_DIR, "auth.json")

SHEET_INV     = "Inventory"
SHEET_HISTORY = "History"
SHEET_REPLACE = "Replacements"
SHEET_INK     = "Ink"
SHEET_LOG     = "Audit Log"
ALL_SHEETS    = [SHEET_INV, SHEET_HISTORY, SHEET_REPLACE, SHEET_INK, SHEET_LOG]

INV_COLS = [
    "Date Collected","Purchase Date","Equipment Type","Brand/Model","Serial No.",
    "Assigned To","Department/Location","Condition/Status","OS/Firmware",
    "IP Address","MAC Address","CPU","RAM","Disk","Remarks"
]
INK_CODES    = ["BK","C","LC","M","LM","Y"]
INK_NAMES    = {"BK":"Black","C":"Cyan","LC":"Light Cyan","M":"Magenta","LM":"Light Magenta","Y":"Yellow"}
INK_DEFAULTS = {"BK":9,"C":3,"LC":5,"M":1,"LM":5,"Y":11}

def today(): return datetime.date.today().strftime("%Y-%m-%d")
def now():   return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ════════════════════════════════════════════
# EXCEL ENGINE
# ════════════════════════════════════════════
class ExcelEngine:
    def __init__(self):
        self._lock = threading.Lock()
        self._ensure_workbook()

    # ── helpers ──────────────────────────────
    def _init_headers(self, ws, name):
        hdrs = {
            SHEET_INV:     INV_COLS,
            SHEET_HISTORY: ["Date","Asset","Serial No.","Department","Event","Description","By"],
            SHEET_REPLACE: ["Date Replaced","Old Asset","Old Serial","Department","Replaced With","New Serial","New Condition","Note","By"],
            SHEET_INK:     ["Timestamp","Colour Code","Colour Name","Action","Quantity","Big ICT","Small ICT","Note"],
            SHEET_LOG:     ["Timestamp","Action","Description","By"],
        }
        ws.append(hdrs.get(name, []))
        for cell in ws[1]:
            cell.font      = Font(bold=True, color="FFFFFF", size=11)
            cell.fill      = PatternFill("solid", fgColor="1F2937")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

    def _set_widths(self, ws):
        widths = {
            "Date Collected":13,"Purchase Date":13,"Equipment Type":22,"Brand/Model":20,
            "Serial No.":22,"Assigned To":20,"Department/Location":24,"Condition/Status":22,
            "OS/Firmware":16,"IP Address":16,"MAC Address":18,"CPU":14,"RAM":12,"Disk":20,"Remarks":32
        }
        if ws.max_row < 1: return
        h = [str(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        for i, col_name in enumerate(h, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(col_name, 16)

    def _write_log(self, wb, action, desc):
        wb[SHEET_LOG].append([now(), action, desc, "ICT Manager"])

    def _backup(self):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(BACKUP_DIR, f"ICT_backup_{ts}.xlsx")
        shutil.copy2(MASTER_XL, dst)

    # ── workbook bootstrap ────────────────────
    def _ensure_workbook(self):
        if not os.path.exists(MASTER_XL):
            self._create_fresh()
            return

        wb = openpyxl.load_workbook(MASTER_XL)

        # Check if this is an old-format workbook (data lives in 'work sheet' etc.)
        OLD_DATA_SHEETS = ["work sheet", "Work Sheet", "worksheet", "WorkSheet"]
        old_data_sheet  = next((s for s in OLD_DATA_SHEETS if s in wb.sheetnames), None)
        needs_migration = SHEET_INV not in wb.sheetnames and old_data_sheet is not None

        if needs_migration:
            print(f"[ICT] Migrating data from '{old_data_sheet}' ...")
            # Back up original before touching it
            os.makedirs(BACKUP_DIR, exist_ok=True)
            shutil.copy2(MASTER_XL, os.path.join(BACKUP_DIR, "ICT_original_before_migration.xlsx"))

            # Read old data BEFORE removing anything
            old_rows = self._read_old_rows(wb, old_data_sheet)

            # Remove ALL old sheets — avoids openpyxl silent rename collisions
            for sname in list(wb.sheetnames):
                del wb[sname]

            # Create all required sheets fresh
            for sheet in ALL_SHEETS:
                ws = wb.create_sheet(sheet)
                self._init_headers(ws, sheet)

            # Write migrated inventory rows
            ws_inv = wb[SHEET_INV]
            for row in old_rows:
                ws_inv.append(row)
            self._set_widths(ws_inv)

            wb.save(MASTER_XL)
            wb.close()
            print(f"[ICT] Migration done — {len(old_rows)} assets imported.")
            return

        # Workbook already has Inventory sheet — check if any other required sheets are missing
        changed = False
        for sheet in ALL_SHEETS:
            if sheet not in wb.sheetnames:
                ws = wb.create_sheet(sheet)
                self._init_headers(ws, sheet)
                changed = True

        # If Inventory exists but still has old column layout, migrate in-place
        if SHEET_INV in wb.sheetnames:
            ws = wb[SHEET_INV]
            if ws.max_row >= 1:
                h = [str(c.value).strip() if c.value else '' for c in ws[1]]
                if "RAM Size" in h or ("Equipment Type" in h and "RAM" not in h):
                    self._migrate_inplace(wb)
                    changed = True

        if changed:
            wb.save(MASTER_XL)
        wb.close()

    def _read_old_rows(self, wb, sheet_name):
        """Read and transform rows from old-format 'work sheet'."""
        ws   = wb[sheet_name]
        rows = list(ws.values)
        if not rows:
            return []

        old_h = [str(c).strip() if c else '' for c in rows[0]]

        def g(row, name, default=''):
            try:
                i = old_h.index(name)
                v = row[i] if i < len(row) else None
                if v is None: return default
                # Handle Python date/datetime objects directly
                if isinstance(v, (datetime.date, datetime.datetime)):
                    return v.strftime('%Y-%m-%d')
                s = str(v).strip()
                # Only strip time if it genuinely looks like a datetime string
                if re.match(r'\d{4}-\d{2}-\d{2}[ T]', s):
                    s = s[:10]
                return default if s.lower() in ['nan', 'none', 'nat', '<na>'] else s
            except ValueError:
                return default

        def fix_ip(ip):
            return (ip or '').replace('192.068.', '192.168.').replace('192.0.168.', '192.168.')

        result = []
        for row in rows[1:]:
            if not any(v for v in row):
                continue
            ram  = ' '.join(p for p in [g(row, 'RAM Size'), g(row, 'RAM Type')] if p)
            disk = ' '.join(p for p in [g(row, 'Disk Size'), g(row, 'Disk Type'), g(row, 'Disk Health')] if p)
            result.append([
                g(row, 'Date Collected'),
                g(row, 'Purchase Date'),
                g(row, 'Equipment Type'),
                g(row, 'Brand/Model'),
                g(row, 'Serial No.'),
                g(row, 'Assigned To'),
                g(row, 'Department/Location').strip(),
                g(row, 'Condition/Status') or 'Good',
                g(row, 'OS/Firmware'),
                fix_ip(g(row, 'IP Address')),
                g(row, 'MAC Address'),
                g(row, 'CPU'),
                ram,
                disk,
                g(row, 'Remarks'),
            ])
        return result

    def _create_fresh(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for s in ALL_SHEETS:
            self._init_headers(wb.create_sheet(s), s)
        wb.save(MASTER_XL)

    def _migrate_inplace(self, wb):
        """Migrate Inventory sheet from old multi-column RAM/Disk format."""
        ws   = wb[SHEET_INV]
        rows = list(ws.values)
        if not rows: return
        old_h = [str(c).strip() if c else '' for c in rows[0]]

        def g(row, name, default=''):
            try:
                i = old_h.index(name)
                v = row[i] if i < len(row) else None
                s = str(v).strip() if v else ''
                return '' if s.lower() in ['nan', 'none', 'nat'] else s
            except ValueError: return default

        ws.delete_rows(1, ws.max_row)
        ws.append(INV_COLS)
        for cell in ws[1]:
            cell.font      = Font(bold=True, color="FFFFFF", size=11)
            cell.fill      = PatternFill("solid", fgColor="1F2937")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in rows[1:]:
            ram  = ' '.join(p for p in [g(row, 'RAM Size'), g(row, 'RAM Type')] if p)
            disk = ' '.join(p for p in [g(row, 'Disk Size'), g(row, 'Disk Type'), g(row, 'Disk Health')] if p)
            ws.append([
                g(row, 'Date Collected').split(' ')[0],
                g(row, 'Purchase Date').split(' ')[0],
                g(row, 'Equipment Type'),
                g(row, 'Brand/Model'),
                g(row, 'Serial No.'),
                g(row, 'Assigned To'),
                g(row, 'Department/Location').strip(),
                g(row, 'Condition/Status') or 'Good',
                g(row, 'OS/Firmware'),
                g(row, 'IP Address'),
                g(row, 'MAC Address'),
                g(row, 'CPU'),
                ram, disk,
                g(row, 'Remarks'),
            ])
        self._set_widths(ws)

    # ── READ ─────────────────────────────────
    def load_all(self):
        with self._lock:
            wb   = openpyxl.load_workbook(MASTER_XL, data_only=True)
            inv  = self._read_inv(wb)
            ink  = self._read_ink(wb)
            log  = self._read_log(wb)
            hist = self._read_hist(wb)
            wb.close()
            return {"inventory": inv, "ink": ink, "log": log, "history": hist}

    def _clean(self, v):
        if v is None: return ''
        if isinstance(v, (datetime.date, datetime.datetime)):
            return v.strftime('%Y-%m-%d')
        s = str(v).strip()
        if re.match(r'\d{4}-\d{2}-\d{2}[ T]', s):
            s = s[:10]
        return '' if s.lower() in ['nan', 'none', 'nat', '<na>'] else s

    def _read_inv(self, wb):
        ws   = wb[SHEET_INV]
        rows = list(ws.values)
        if len(rows) < 2: return []
        h   = [str(c).strip() if c else '' for c in rows[0]]
        out = []
        for i, row in enumerate(rows[1:], start=2):
            d = {'_row': i}
            for j, col in enumerate(h):
                d[col] = self._clean(row[j] if j < len(row) else None)
            out.append(d)
        return out

    def _read_ink(self, wb):
        ws    = wb[SHEET_INK]
        rows  = list(ws.values)
        state = {k: {"store": INK_DEFAULTS[k], "bigIct": 1, "smallIct": 1} for k in INK_CODES}
        for row in rows[1:]:
            if not row or not row[1]: continue
            code   = str(row[1]).strip()
            action = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            try: qty = int(float(str(row[4]))) if len(row) > 4 and row[4] else 0
            except: qty = 0
            try: big = int(float(str(row[5]))) if len(row) > 5 and row[5] else None
            except: big = None
            try: sml = int(float(str(row[6]))) if len(row) > 6 and row[6] else None
            except: sml = None
            if code in state:
                if action == 'restock': state[code]['store'] += qty
                elif action == 'use':   state[code]['store'] = max(0, state[code]['store'] - qty)
                if big is not None: state[code]['bigIct']   = big
                if sml is not None: state[code]['smallIct'] = sml
        return state

    def _read_log(self, wb):
        ws  = wb[SHEET_LOG]
        out = []
        for row in list(ws.values)[1:]:
            out.append({
                'time':   str(row[0]) if row[0] else '',
                'action': str(row[1]) if len(row) > 1 and row[1] else '',
                'desc':   str(row[2]) if len(row) > 2 and row[2] else '',
            })
        return list(reversed(out))

    def _read_hist(self, wb):
        ws  = wb[SHEET_HISTORY]
        out = []
        for row in list(ws.values)[1:]:
            out.append({
                'date':   str(row[0]) if row[0] else '',
                'asset':  str(row[1]) if len(row) > 1 and row[1] else '',
                'serial': str(row[2]) if len(row) > 2 and row[2] else '',
                'dept':   str(row[3]) if len(row) > 3 and row[3] else '',
                'event':  str(row[4]) if len(row) > 4 and row[4] else '',
                'note':   str(row[5]) if len(row) > 5 and row[5] else '',
            })
        return list(reversed(out))

    # ── WRITE ────────────────────────────────
    def add_asset(self, data):
        with self._lock:
            self._backup()
            wb = openpyxl.load_workbook(MASTER_XL)
            ws = wb[SHEET_INV]
            ws.append([data.get(c, '') for c in INV_COLS])
            new_row = ws.max_row
            self._set_widths(ws)
            self._write_log(wb, 'add', f"Added: {data.get('Equipment Type','')} — {data.get('Brand/Model','')} | {data.get('Department/Location','')}")
            wb.save(MASTER_XL); wb.close()
            return new_row

    def update_asset(self, row_num, data):
        with self._lock:
            self._backup()
            wb = openpyxl.load_workbook(MASTER_XL)
            ws = wb[SHEET_INV]
            h  = [ws.cell(1, c).value for c in range(1, len(INV_COLS) + 2)]
            for j, col in enumerate(h, 1):
                if col in data:
                    ws.cell(row=row_num, column=j).value = data[col]
            self._write_log(wb, 'edit', f"Edited row {row_num}: {data.get('Equipment Type','')} — {data.get('Brand/Model','')}")
            wb.save(MASTER_XL); wb.close()

    def delete_asset(self, row_num, desc):
        with self._lock:
            self._backup()
            wb = openpyxl.load_workbook(MASTER_XL)
            wb[SHEET_INV].delete_rows(row_num)
            self._write_log(wb, 'delete', f"Deleted: {desc}")
            wb.save(MASTER_XL); wb.close()

    def record_replacement(self, old_row, old_desc, rep):
        with self._lock:
            self._backup()
            wb     = openpyxl.load_workbook(MASTER_XL)
            ws_inv = wb[SHEET_INV]
            h      = [ws_inv.cell(1, c).value for c in range(1, len(INV_COLS) + 2)]
            def col_idx(name):
                try: return h.index(name) + 1
                except: return None
            ci = col_idx("Condition/Status")
            ri = col_idx("Remarks")
            if ci: ws_inv.cell(row=old_row, column=ci).value = "Replaced"
            if ri:
                old_rem = ws_inv.cell(row=old_row, column=ri).value or ''
                ws_inv.cell(row=old_row, column=ri).value = (str(old_rem) + f" | REPLACED {rep['date']}: {rep['note']}").strip(" |")
            ws_inv.append([{
                "Date Collected": rep['date'], "Purchase Date": rep['date'],
                "Equipment Type": rep.get('newType',''), "Brand/Model": rep.get('newBrand',''),
                "Serial No.": rep.get('newSerial',''), "Assigned To": rep.get('assigned',''),
                "Department/Location": rep.get('dept',''), "Condition/Status": rep.get('newCond','New'),
                "OS/Firmware":"","IP Address": rep.get('ip',''),"MAC Address":"","CPU":"","RAM":"","Disk":"",
                "Remarks": f"Replaced: {old_desc} on {rep['date']}. {rep['note']}"
            }.get(c,'') for c in INV_COLS])
            self._set_widths(ws_inv)
            wb[SHEET_REPLACE].append([rep['date'], old_desc, rep.get('oldSerial',''), rep.get('dept',''),
                f"{rep.get('newType','')} {rep.get('newBrand','')}".strip(), rep.get('newSerial',''),
                rep.get('newCond','New'), rep['note'], "ICT Manager"])
            wb[SHEET_HISTORY].append([rep['date'], old_desc, rep.get('oldSerial',''), rep.get('dept',''), "replaced",
                f"Replaced with: {rep.get('newType','')} {rep.get('newBrand','')} S/N:{rep.get('newSerial','')}. Note: {rep['note']}", "ICT Manager"])
            self._write_log(wb, 'replace', f"Replaced {old_desc} → {rep.get('newBrand','')} on {rep['date']}")
            wb.save(MASTER_XL); wb.close()

    def write_history_event(self, asset_desc, serial, dept, event, note, date):
        with self._lock:
            wb = openpyxl.load_workbook(MASTER_XL)
            wb[SHEET_HISTORY].append([date, asset_desc, serial, dept, event, note, "ICT Manager"])
            if event == 'repaired':
                ws = wb[SHEET_INV]
                h  = [ws.cell(1, c).value for c in range(1, len(INV_COLS) + 2)]
                try: ser_col  = h.index("Serial No.") + 1
                except: ser_col = None
                try: cond_col = h.index("Condition/Status") + 1
                except: cond_col = None
                if ser_col and cond_col:
                    for row in range(2, ws.max_row + 1):
                        if str(ws.cell(row, ser_col).value or '').strip() == serial.strip():
                            if ws.cell(row, cond_col).value == "Fair (Needs Repair)":
                                ws.cell(row, cond_col).value = "Good"
                            break
            self._write_log(wb, 'history', f"{event} on {asset_desc}: {note[:80]}")
            wb.save(MASTER_XL); wb.close()

    def write_ink(self, code, action, qty, big, sml, note=''):
        with self._lock:
            wb = openpyxl.load_workbook(MASTER_XL)
            wb[SHEET_INK].append([now(), code, INK_NAMES.get(code, code), action, qty, big, sml, note])
            self._write_log(wb, 'ink', f"Ink {action}: {qty}x {INK_NAMES.get(code, code)} ({code})")
            wb.save(MASTER_XL); wb.close()


# ════════════════════════════════════════════
# JS API BRIDGE
# ════════════════════════════════════════════
class Api:
    def __init__(self, engine: ExcelEngine):
        self.engine = engine
        self.window = None

    def load_data(self):
        try:
            return json.dumps({"ok": True, "data": self.engine.load_all()})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def add_asset(self, data_json):
        try:
            data = json.loads(data_json)
            if not data.get('Equipment Type'):
                return json.dumps({"ok": False, "error": "Equipment Type is required"})
            if not data.get('Date Collected'):
                data['Date Collected'] = today()
            return json.dumps({"ok": True, "row": self.engine.add_asset(data)})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def update_asset(self, row_num, data_json):
        try:
            self.engine.update_asset(int(row_num), json.loads(data_json))
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def delete_asset(self, row_num, desc):
        try:
            self.engine.delete_asset(int(row_num), desc)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def record_replacement(self, rep_json):
        try:
            rep = json.loads(rep_json)
            if not rep.get('note'):
                return json.dumps({"ok": False, "error": "Replacement note is required"})
            self.engine.record_replacement(int(rep['oldRow']), rep['oldDesc'], rep)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def write_history(self, asset_desc, serial, dept, event, note, date):
        try:
            self.engine.write_history_event(asset_desc, serial, dept, event, note, date)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def ink_action(self, code, action, qty, big, sml, note):
        try:
            if action == 'use':
                cur = self.engine.load_all()['ink'].get(code, {}).get('store', 0)
                if cur < int(qty):
                    return json.dumps({"ok": False, "error": f"Only {cur} in store"})
            self.engine.write_ink(code, action, int(qty), int(big), int(sml), note)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def get_master_path(self):
        return MASTER_XL

    def open_excel(self):
        import subprocess
        try:
            if sys.platform == 'win32':    os.startfile(MASTER_XL)
            elif sys.platform == 'darwin': subprocess.Popen(['open', MASTER_XL])
            else:                          subprocess.Popen(['xdg-open', MASTER_XL])
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    # ── AUTH ──────────────────────────────────
    def is_pin_set(self):
        return json.dumps({"ok": True, "set": os.path.exists(AUTH_FILE)})

    def set_pin(self, pin):
        try:
            if len(pin.strip()) < 4:
                return json.dumps({"ok": False, "error": "PIN must be at least 4 characters"})
            salt = secrets.token_hex(16)
            h    = hashlib.sha256((salt + pin).encode()).hexdigest()
            with open(AUTH_FILE, 'w') as f:
                json.dump({"hash": h, "salt": salt}, f)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def verify_pin(self, pin):
        try:
            if not os.path.exists(AUTH_FILE):
                return json.dumps({"ok": False, "error": "No PIN set"})
            with open(AUTH_FILE) as f:
                data = json.load(f)
            h = hashlib.sha256((data['salt'] + pin).encode()).hexdigest()
            return json.dumps({"ok": True, "valid": h == data['hash']})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})


# ════════════════════════════════════════════
# LAUNCH
# ════════════════════════════════════════════
def main():
    engine = ExcelEngine()
    api    = Api(engine)

    ui_path = os.path.join(BASE_DIR, 'ui.html')
    if not os.path.exists(ui_path):
        print("ERROR: ui.html not found in", BASE_DIR)
        sys.exit(1)

    window = webview.create_window(
        title            = "St. Anne Mission Hospital — ICT Command Centre",
        url              = ui_path,
        js_api           = api,
        width            = 1400,
        height           = 860,
        min_size         = (1000, 650),
        background_color = "#0d1117",
        confirm_close    = True,
        text_select      = True,
    )
    api.window = window
    webview.start(debug=False, private_mode=False)

if __name__ == "__main__":
    main()
