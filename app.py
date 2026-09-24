"""
Username Shuffler v3.0.0
A professional desktop application for shuffling, formatting, managing,
and exporting usernames from Excel (.xlsx, .xlsm), CSV, and Text files.

Author: MD Showrav Zaman
License: MIT
"""

from __future__ import annotations

import csv
import json
import os
import queue
import random
import re
import sys
import threading
import time
import zipfile
from datetime import datetime, timezone
from posixpath import normpath
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

# Optional darkdetect for automatic OS theme detection
try:
    import darkdetect  # type: ignore
except ImportError:
    darkdetect = None

# Optional openpyxl for robust Excel appending
try:
    from openpyxl import Workbook, load_workbook  # type: ignore
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

APP_NAME = "Username Shuffler"
APP_VERSION = "v3.0.0"
APP_SUBTITLE = "Professional Username Randomizer & Dataset Manager"
DATA_EXTENSIONS = (".xlsx", ".xlsm", ".csv", ".txt", ".xls")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

HEADER_KEYWORDS = {
    "username", "usernames", "user", "users", "handle", "handles",
    "name", "names", "account", "accounts", "alias", "tag", "member",
    "members", "player", "players", "participant", "participants", "id", "email"
}


# ============================================================================
# Theme Definitions
# ============================================================================

THEMES: dict[str, dict[str, str]] = {
    "light": {
        "bg": "#F1F5F9",                # Slate 100
        "card_bg": "#FFFFFF",
        "card_border": "#CBD5E1",       # Slate 300
        "card_border_focus": "#2563EB",
        "toolbar_bg": "#FFFFFF",
        "toolbar_border": "#E2E8F0",
        "text": "#0F172A",              # Slate 900
        "text_secondary": "#334155",    # Slate 700
        "text_muted": "#64748B",        # Slate 500
        "input_bg": "#FFFFFF",
        "input_fg": "#0F172A",
        "input_border": "#CBD5E1",
        "input_focus": "#2563EB",
        "combo_bg": "#FFFFFF",
        "combo_hover_bg": "#F8FAFC",
        "combo_arrow_bg": "#E2E8F0",
        "combo_arrow_hover": "#CBD5E1",
        "primary": "#2563EB",           # Blue 600
        "primary_hover": "#1D4ED8",
        "primary_fg": "#FFFFFF",
        "success": "#059669",           # Emerald 600
        "success_hover": "#047857",
        "success_fg": "#FFFFFF",
        "warning": "#D97706",           # Amber 600
        "warning_hover": "#B45309",
        "warning_fg": "#FFFFFF",
        "secondary_btn": "#F8FAFC",
        "secondary_btn_border": "#CBD5E1",
        "secondary_btn_hover": "#E2E8F0",
        "secondary_btn_fg": "#1E293B",
        "chip_bg": "#E2E8F0",
        "chip_hover": "#CBD5E1",
        "chip_fg": "#1E293B",
        "result_bg": "#FFFFFF",
        "result_fg": "#0F172A",
        "result_border": "#CBD5E1",
        "badge_bg": "#EFF6FF",
        "badge_fg": "#1D4ED8",
        "badge_border": "#BFDBFE",
        "accent_chip": "#DCFCE7",
        "accent_chip_fg": "#15803D",
        "status_bg": "#FFFFFF",
        "status_fg": "#334155",
        "status_border": "#E2E8F0",
        "tab_active_bg": "#FFFFFF",
        "tab_inactive_bg": "#E2E8F0",
    },
    "dark": {
        "bg": "#0B0F19",                # Deep slate background
        "card_bg": "#141C2B",           # Elevated dark card
        "card_border": "#28364B",       # Subtle card border
        "card_border_focus": "#3B82F6",
        "toolbar_bg": "#111827",        # Dark toolbar
        "toolbar_border": "#1E293B",
        "text": "#F8FAFC",              # Crisp off-white text
        "text_secondary": "#CBD5E1",    # Slate 300
        "text_muted": "#94A3B8",        # Slate 400 - high contrast
        "input_bg": "#1E293B",          # Elevated dark field
        "input_fg": "#F8FAFC",
        "input_border": "#3B4D66",
        "input_focus": "#3B82F6",
        "combo_bg": "#1E293B",          # Elevated dark combobox background
        "combo_hover_bg": "#253347",
        "combo_arrow_bg": "#2B3A50",    # Distinct, visible arrow button
        "combo_arrow_hover": "#3B4D66",
        "primary": "#3B82F6",           # Blue 500
        "primary_hover": "#2563EB",
        "primary_fg": "#FFFFFF",
        "success": "#10B981",           # Emerald 500
        "success_hover": "#059669",
        "success_fg": "#FFFFFF",
        "warning": "#F59E0B",           # Amber 500
        "warning_hover": "#D97706",
        "warning_fg": "#0F172A",
        "secondary_btn": "#1E293B",
        "secondary_btn_border": "#334155",
        "secondary_btn_hover": "#2B3A50",
        "secondary_btn_fg": "#F8FAFC",
        "chip_bg": "#1E293B",
        "chip_hover": "#2B3C53",
        "chip_fg": "#E2E8F0",
        "result_bg": "#0D131F",
        "result_fg": "#F8FAFC",
        "result_border": "#28364B",
        "badge_bg": "#1E293B",
        "badge_fg": "#60A5FA",
        "badge_border": "#3B82F6",
        "accent_chip": "#064E3B",
        "accent_chip_fg": "#6EE7B7",
        "status_bg": "#111827",
        "status_fg": "#94A3B8",
        "status_border": "#1E293B",
        "tab_active_bg": "#141C2B",
        "tab_inactive_bg": "#1E293B",
    },
}


# ============================================================================
# Helpers & Utilities
# ============================================================================

def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def get_app_folder() -> str:
    """Return the folder where the executable or main script resides."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_config_dir() -> str:
    """Return an AppData directory for storing user settings and history."""
    app_data = os.getenv("APPDATA")
    if app_data and os.path.isdir(app_data):
        target = os.path.join(app_data, "UsernameShuffler")
    else:
        target = os.path.join(get_app_folder(), ".settings")
    os.makedirs(target, exist_ok=True)
    return target


def clean_username(value) -> str:
    """Normalize and sanitize a single username string."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "undefined", "n/a"}:
        return ""
    return text


def cell_col_to_index(col_letters: str) -> int:
    """Convert Excel column letters (A, B, AA) to zero-based column index."""
    idx = 0
    for ch in col_letters.upper():
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        else:
            break
    return max(0, idx - 1)


# ============================================================================
# Configuration Manager
# ============================================================================

class ConfigManager:
    """Persists user preferences, window geometry, and recent files."""

    def __init__(self) -> None:
        self.file_path = os.path.join(get_config_dir(), "settings.json")
        self.data: dict = {
            "theme": "dark" if (darkdetect and darkdetect.isDark()) else "light",
            "last_file": "",
            "recent_files": [],
            "delimiter": "newline",
            "custom_delimiter": ", ",
            "prefix": "",
            "suffix": "",
            "numbered": False,
            "case": "original",
            "mode": "unique",
            "auto_shuffle": True,
            "window_geometry": "1040x680",
        }
        self.load()

    def load(self) -> None:
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if isinstance(saved, dict):
                        self.data.update(saved)
            except Exception:
                pass

    def save(self) -> None:
        try:
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2)
        except Exception:
            pass

    def get(self, key: str, default=None):
        return self.data.get(key, default)

    def set(self, key: str, value) -> None:
        self.data[key] = value
        self.save()

    def add_recent_file(self, path: str) -> None:
        if not path or not os.path.exists(path):
            return
        recents: list[str] = [p for p in self.data.get("recent_files", []) if p != path and os.path.exists(p)]
        recents.insert(0, path)
        self.data["recent_files"] = recents[:8]
        self.data["last_file"] = path
        self.save()


# ============================================================================
# High-Performance Dataset Parsers & Exporters
# ============================================================================

def parse_xlsx_stream(path: str) -> list[list[str]]:
    """Stream XML from .xlsx or .xlsm file with minimal memory usage."""
    with zipfile.ZipFile(path, "r") as zf:
        namelist = set(zf.namelist())

        # Read shared strings if present
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in namelist:
            with zf.open("xl/sharedStrings.xml") as f:
                for _, elem in ET.iterparse(f):
                    if elem.tag.endswith("si"):
                        shared_strings.append("".join(elem.itertext()))
                        elem.clear()

        # Locate first worksheet
        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in namelist and "xl/_rels/workbook.xml.rels" in namelist:
            try:
                wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
                sheets = wb_xml.find(qn(MAIN_NS, "sheets"))
                if sheets is not None and list(sheets):
                    first_sheet = list(sheets)[0]
                    rel_id = first_sheet.attrib.get(qn(REL_NS, "id"))
                    if rel_id:
                        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                        for rel in rels_xml:
                            if rel.attrib.get("Id") == rel_id:
                                target = rel.attrib.get("Target", "worksheets/sheet1.xml")
                                sheet_path = normpath(f"xl/{target}") if not target.startswith("/") else normpath(target.lstrip("/"))
                                break
            except Exception:
                sheet_path = "xl/worksheets/sheet1.xml"

        if sheet_path not in namelist:
            sheets_available = [s for s in namelist if s.startswith("xl/worksheets/sheet") and s.endswith(".xml")]
            if sheets_available:
                sheet_path = sorted(sheets_available)[0]
            else:
                raise RuntimeError("No worksheet found inside Excel file.")

        rows: list[list[str]] = []
        with zf.open(sheet_path) as f:
            for _, elem in ET.iterparse(f):
                if elem.tag.endswith("row"):
                    row_cells: dict[int, str] = {}
                    for c in elem.findall(qn(MAIN_NS, "c")):
                        cell_ref = c.attrib.get("r", "")
                        col_letters = "".join(filter(str.isalpha, cell_ref))
                        col_idx = cell_col_to_index(col_letters) if col_letters else len(row_cells)
                        cell_type = c.attrib.get("t", "")

                        val = ""
                        if cell_type == "s":
                            v = c.find(qn(MAIN_NS, "v"))
                            if v is not None and v.text and v.text.isdigit():
                                idx = int(v.text)
                                if 0 <= idx < len(shared_strings):
                                    val = shared_strings[idx]
                        elif cell_type == "inlineStr":
                            val = "".join(c.itertext())
                        else:
                            v = c.find(qn(MAIN_NS, "v"))
                            val = v.text if (v is not None and v.text is not None) else ""

                        cleaned = clean_username(val)
                        if cleaned:
                            row_cells[col_idx] = cleaned

                    if row_cells:
                        max_col = max(row_cells.keys())
                        row_list = [row_cells.get(i, "") for i in range(max_col + 1)]
                        rows.append(row_list)
                    elem.clear()

        return rows


def parse_csv_stream(path: str) -> list[list[str]]:
    """Read CSV file with automatic encoding and delimiter detection."""
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
                    delimiter = dialect.delimiter
                except Exception:
                    delimiter = ","

                reader = csv.reader(handle, delimiter=delimiter)
                rows: list[list[str]] = []
                for row in reader:
                    cleaned_row = [clean_username(c) for c in row]
                    if any(cleaned_row):
                        rows.append(cleaned_row)
                return rows
        except (UnicodeDecodeError, Exception):
            continue
    raise RuntimeError(f"Could not decode CSV file '{os.path.basename(path)}'.")


def parse_txt_stream(path: str) -> list[list[str]]:
    """Read plain text files line by line."""
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as handle:
                rows: list[list[str]] = []
                for line in handle:
                    val = clean_username(line)
                    if val:
                        rows.append([val])
                return rows
        except (UnicodeDecodeError, Exception):
            continue
    raise RuntimeError(f"Could not decode text file '{os.path.basename(path)}'.")


def load_dataset_table(path: str) -> tuple[list[str], int, list[str], bool]:
    """
    Loads rows, detects headers, selects the most relevant column,
    and returns (column_names, active_col_index, usernames_in_active_col, has_header).
    """
    lower = path.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        rows = parse_xlsx_stream(path)
    elif lower.endswith(".csv"):
        rows = parse_csv_stream(path)
    elif lower.endswith(".txt"):
        rows = parse_txt_stream(path)
    elif lower.endswith(".xls"):
        raise RuntimeError("Old .xls format is not supported. Please save as .xlsx, .csv, or .txt.")
    else:
        raise RuntimeError(f"Unsupported file format: {os.path.splitext(path)[1]}")

    if not rows:
        return ["Column 1"], 0, [], False

    num_cols = max(len(r) for r in rows)
    first_row = [rows[0][i] if i < len(rows[0]) else "" for i in range(num_cols)]

    has_header = False
    best_col_idx = 0

    # Test if first row looks like a header row
    for idx, cell in enumerate(first_row):
        clean_tag = cell.lower().replace(" ", "").replace("_", "").replace("-", "")
        if clean_tag in HEADER_KEYWORDS:
            has_header = True
            best_col_idx = idx
            break

    if has_header:
        column_names = [first_row[i] if (i < len(first_row) and first_row[i]) else f"Column {i+1}" for i in range(num_cols)]
        data_rows = rows[1:]
    else:
        column_names = [f"Column {i+1}" for i in range(num_cols)]
        data_rows = rows

    usernames = []
    for r in data_rows:
        if best_col_idx < len(r):
            val = r[best_col_idx]
            if val:
                usernames.append(val)

    return column_names, best_col_idx, usernames, has_header


def find_default_dataset_file(folder: str) -> str | None:
    """Find the best dataset file in the directory."""
    if not os.path.isdir(folder):
        return None

    files: list[str] = []
    for fname in os.listdir(folder):
        lower = fname.lower()
        if fname.startswith("~$") or fname.startswith("."):
            continue
        if lower.endswith(DATA_EXTENSIONS):
            files.append(fname)

    if not files:
        return None

    # Priority filenames
    priority = [
        "username dataset.xlsx", "usernames.xlsx", "username.xlsx",
        "usernames.csv", "username.csv", "users.xlsx", "users.csv",
        "usernames.txt", "users.txt"
    ]
    for pref in priority:
        for fname in files:
            if fname.lower() == pref:
                return os.path.join(folder, fname)

    # Return first alphabetically
    files.sort(key=str.lower)
    return os.path.join(folder, files[0])


def write_simple_xlsx(path: str, names: list[str]) -> None:
    """Create a clean .xlsx spreadsheet without external libraries."""
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    rows: list[str] = []
    for row_index, name in enumerate(names, start=1):
        safe = escape(name, {'"': "&quot;"})
        rows.append(
            f'<row r="{row_index}"><c r="A{row_index}" t="inlineStr"><is><t xml:space="preserve">{safe}</t></is></c></row>'
        )
    sheet_xml = "".join(rows)
    dimension = f"A1:A{max(1, len(names))}"

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>'
        ),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f'<Application>{APP_NAME}</Application>'
            '</Properties>'
        ),
        "docProps/core.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{APP_NAME}</dc:creator>'
            f'<cp:lastModifiedBy>{APP_NAME}</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
            '</cp:coreProperties>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Usernames" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        ),
        "xl/worksheets/sheet1.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="{dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="36" customWidth="1"/></cols>'
            f'<sheetData>{sheet_xml}</sheetData></worksheet>'
        ),
    }

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_name, content in files.items():
            zf.writestr(file_name, content)


def append_names_to_file(path: str, new_names: list[str], existing_all: list[str]) -> str:
    """Append new usernames to the given file format, preserving file type."""
    lower = path.lower()

    if lower.endswith(".csv"):
        with open(path, "a", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            for name in new_names:
                writer.writerow([name])
        return path

    if lower.endswith(".txt"):
        with open(path, "a", encoding="utf-8") as handle:
            for name in new_names:
                handle.write(f"{name}\n")
        return path

    if lower.endswith((".xlsx", ".xlsm")):
        if HAS_OPENPYXL and os.path.exists(path):
            try:
                wb = load_workbook(path)
                ws = wb.active
                for name in new_names:
                    ws.append([name])
                wb.save(path)
                wb.close()
                return path
            except Exception:
                pass
        # Fallback to write_simple_xlsx with all names
        combined = list(existing_all) + new_names
        write_simple_xlsx(path, combined)
        return path

    # Fallback to new xlsx
    new_path = os.path.join(get_app_folder(), "usernames.xlsx")
    write_simple_xlsx(new_path, list(existing_all) + new_names)
    return new_path


# ============================================================================
# Main Application GUI
# ============================================================================

class UsernameShufflerApp:
    def __init__(self, root: tk.Tk, target_file: str | None = None) -> None:
        self.root = root
        self.config = ConfigManager()
        self.folder = get_app_folder()
        self.task_queue: queue.Queue = queue.Queue()

        # Theme & Styling
        self.theme_name = self.config.get("theme", "light")
        if self.theme_name not in THEMES:
            self.theme_name = "light"
        self.colors = THEMES[self.theme_name]

        # Dataset State
        self.loaded_file: str | None = None
        self.usernames: list[str] = []
        self.columns: list[str] = ["Column 1"]
        self.active_col_index = 0
        self.has_header = False
        self.elimination_pool: list[str] = []

        # Tracking widgets for dynamic theme recoloring
        self.cards: list[tk.Widget] = []
        self.card_labels: list[tk.Label] = []
        self.muted_labels: list[tk.Label] = []
        self.secondary_buttons: list[tk.Button] = []
        self.chip_buttons: list[tk.Button] = []
        self.radio_buttons: list[tk.Widget] = []

        # Results & History
        self.last_shuffled_list: list[str] = []
        self.formatted_result: str = ""
        self.draw_history: list[dict] = []
        self._trace_active = False

        # Tkinter StringVars
        self.count_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Initializing application...")
        self.file_name_var = tk.StringVar(value="No file loaded")
        self.file_badge_var = tk.StringVar(value="Looking for dataset...")
        self.total_count_var = tk.StringVar(value="0")
        self.picked_count_var = tk.StringVar(value="0")
        self.remaining_pool_var = tk.StringVar(value="0")
        self.col_var = tk.StringVar(value="Column 1")

        # Options Vars
        self.mode_var = tk.StringVar(value=self.config.get("mode", "unique"))
        self.delimiter_var = tk.StringVar(value=self.config.get("delimiter", "newline"))
        self.custom_delimiter_var = tk.StringVar(value=self.config.get("custom_delimiter", ", "))
        self.prefix_var = tk.StringVar(value=self.config.get("prefix", ""))
        self.suffix_var = tk.StringVar(value=self.config.get("suffix", ""))
        self.numbered_var = tk.BooleanVar(value=self.config.get("numbered", False))
        self.case_var = tk.StringVar(value=self.config.get("case", "original"))
        self.auto_shuffle_var = tk.BooleanVar(value=self.config.get("auto_shuffle", True))

        # Setup GUI
        self.configure_window()
        self.build_menu_bar()
        self.build_top_toolbar()
        self.build_main_layout()
        self.build_status_bar()
        self.bind_shortcuts()
        self.setup_ttk_styles()

        # Connect Traces
        self.count_var.trace_add("write", self._on_count_change)
        self.delimiter_var.trace_add("write", lambda *_: self.reformat_current_result())
        self.numbered_var.trace_add("write", lambda *_: self.reformat_current_result())
        self.case_var.trace_add("write", lambda *_: self.reformat_current_result())
        self.prefix_var.trace_add("write", lambda *_: self.reformat_current_result())
        self.suffix_var.trace_add("write", lambda *_: self.reformat_current_result())
        self.mode_var.trace_add("write", self._on_mode_change)

        self._trace_active = True

        # Start queue poller
        self._poll_task_queue()

        # Initial Load
        initial_target = target_file or self.config.get("last_file")
        if not initial_target or not os.path.exists(initial_target):
            initial_target = find_default_dataset_file(self.folder)

        self.load_file_async(initial_target, initial=True)

    def _poll_task_queue(self) -> None:
        """Process callbacks dispatched from worker threads with 100% thread safety."""
        while not self.task_queue.empty():
            try:
                fn = self.task_queue.get_nowait()
                fn()
            except queue.Empty:
                break
        self.root.after(30, self._poll_task_queue)

    # ------------------------------------------------------------------------
    # Window & Styling Setup
    # ------------------------------------------------------------------------

    def configure_window(self) -> None:
        self.root.title(f"{APP_NAME} {APP_VERSION} - {APP_SUBTITLE}")
        saved_geom = self.config.get("window_geometry", "1040x680")
        self.root.geometry(saved_geom)
        self.root.minsize(920, 580)
        self.root.configure(bg=self.colors["bg"])

        self.apply_window_icons()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def apply_window_icons(self) -> None:
        """Apply custom icons cleanly without leaks."""
        titlebar_ico = resource_path("titlebar.ico")
        icon_ico = resource_path("icon.ico")
        icon_png = resource_path("icon.png")

        preferred_ico = titlebar_ico if os.path.exists(titlebar_ico) else icon_ico
        try:
            if os.path.exists(preferred_ico):
                self.root.iconbitmap(default=preferred_ico)
        except Exception:
            pass

        try:
            if os.path.exists(icon_png):
                self.photo_icon = tk.PhotoImage(file=icon_png)
                self.root.iconphoto(True, self.photo_icon)
        except Exception:
            pass

        if os.name == "nt" and os.path.exists(preferred_ico):
            try:
                import ctypes
                hwnd = self.root.winfo_id()
                WM_SETICON = 0x0080
                LR_LOADFROMFILE = 0x00000010
                IMAGE_ICON = 1
                user32 = ctypes.windll.user32
                h_icon_small = user32.LoadImageW(None, preferred_ico, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
                h_icon_big = user32.LoadImageW(None, preferred_ico, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
                if h_icon_small:
                    user32.SendMessageW(hwnd, WM_SETICON, 0, h_icon_small)
                if h_icon_big:
                    user32.SendMessageW(hwnd, WM_SETICON, 1, h_icon_big)
            except Exception:
                pass

    def setup_ttk_styles(self) -> None:
        try:
            style = ttk.Style()
            if "clam" in style.theme_names():
                style.theme_use("clam")
            c = self.colors
            style.configure(
                "TNotebook",
                background=c["card_bg"],
                borderwidth=0,
                tabmargins=[2, 5, 2, 0]
            )
            style.configure(
                "TNotebook.Tab",
                background=c["tab_inactive_bg"],
                foreground=c["text_secondary"],
                font=("Segoe UI", 9, "bold"),
                padding=[12, 5]
            )
            style.map(
                "TNotebook.Tab",
                background=[("selected", c["tab_active_bg"])],
                foreground=[("selected", c["primary"])]
            )
            style.configure(
                "TCombobox",
                fieldbackground=c["combo_bg"],
                background=c["combo_arrow_bg"],
                foreground=c["text"],
                arrowcolor=c["text"],
                bordercolor=c["card_border"],
                darkcolor=c["card_border"],
                lightcolor=c["card_border"],
                padding=[8, 4],
            )
            style.map(
                "TCombobox",
                fieldbackground=[
                    ("readonly", "focus", c["combo_bg"]),
                    ("readonly", "hover", c["combo_hover_bg"]),
                    ("readonly", c["combo_bg"]),
                    ("disabled", c["input_bg"]),
                ],
                foreground=[
                    ("readonly", "focus", c["text"]),
                    ("readonly", "hover", c["text"]),
                    ("readonly", c["text"]),
                    ("disabled", c["text_muted"]),
                ],
                selectbackground=[
                    ("readonly", "focus", c["combo_bg"]),
                    ("readonly", c["combo_bg"]),
                ],
                selectforeground=[
                    ("readonly", "focus", c["text"]),
                    ("readonly", c["text"]),
                ],
                background=[
                    ("readonly", "hover", c["combo_arrow_hover"]),
                    ("readonly", "pressed", c["primary"]),
                    ("readonly", c["combo_arrow_bg"]),
                    ("disabled", c["input_bg"]),
                ],
                arrowcolor=[
                    ("readonly", "disabled", c["text_muted"]),
                    ("readonly", c["text"]),
                ],
                bordercolor=[
                    ("focus", c["primary"]),
                    ("hover", c["card_border_focus"]),
                    ("!disabled", c["card_border"]),
                ]
            )

            # Style the dropdown pop-down listbox
            self.root.option_add("*TCombobox*Listbox.background", c["combo_bg"])
            self.root.option_add("*TCombobox*Listbox.foreground", c["text"])
            self.root.option_add("*TCombobox*Listbox.selectBackground", c["primary"])
            self.root.option_add("*TCombobox*Listbox.selectForeground", c["primary_fg"])
            self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 9, "bold"))
            self.root.option_add("*TCombobox*Listbox.relief", "flat")
            self.root.option_add("*ComboboxPopdown*Frame.relief", "solid")
            self.root.option_add("*ComboboxPopdown*Frame.borderWidth", 1)
        except Exception:
            pass

    # ------------------------------------------------------------------------
    # Top Menu Bar & Top Toolbar
    # ------------------------------------------------------------------------

    def build_menu_bar(self) -> None:
        menubar = tk.Menu(self.root)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open File... (Ctrl+O)", command=self.on_open_file_dialog)
        file_menu.add_command(label="Reload Current File (Ctrl+R / F5)", command=self.reload_active_file)
        file_menu.add_separator()
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Recent Files", menu=self.recent_menu)
        self.update_recent_menu()
        file_menu.add_command(label="Open File Location in Explorer", command=self.open_file_location)
        file_menu.add_separator()
        file_menu.add_command(label="Export Results... (Ctrl+E)", command=self.export_results)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        # Edit Menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Copy Results (Ctrl+C)", command=self.copy_result_to_clipboard)
        edit_menu.add_command(label="Clear Results (Ctrl+K)", command=self.clear_results)
        edit_menu.add_separator()
        edit_menu.add_command(label="Reset Elimination Pool", command=self.reset_elimination_pool)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        # Shuffle Menu
        shuffle_menu = tk.Menu(menubar, tearoff=0)
        shuffle_menu.add_command(label="Shuffle Now (Ctrl+Enter)", command=self.shuffle_manual)
        shuffle_menu.add_separator()
        shuffle_menu.add_radiobutton(label="Mode: Unique Picks (No Repeats)", variable=self.mode_var, value="unique")
        shuffle_menu.add_radiobutton(label="Mode: With Replacement (Allow Repeats)", variable=self.mode_var, value="replacement")
        shuffle_menu.add_radiobutton(label="Mode: Elimination Draw (Giveaway/Raffle)", variable=self.mode_var, value="elimination")
        shuffle_menu.add_separator()
        shuffle_menu.add_checkbutton(label="Auto-Shuffle on Type", variable=self.auto_shuffle_var)
        menubar.add_cascade(label="Shuffle", menu=shuffle_menu)

        # View Menu
        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Dark / Light Theme (Ctrl+D)", command=self.toggle_theme)
        menubar.add_cascade(label="View", menu=view_menu)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts_dialog)
        help_menu.add_command(label="About Username Shuffler", command=self.show_about_dialog)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def update_recent_menu(self) -> None:
        self.recent_menu.delete(0, tk.END)
        recents = self.config.get("recent_files", [])
        if not recents:
            self.recent_menu.add_command(label="(No recent files)", state="disabled")
            return
        for path in recents:
            self.recent_menu.add_command(
                label=os.path.basename(path),
                command=lambda p=path: self.load_file_async(p, initial=False)
            )

    def build_top_toolbar(self) -> None:
        """Top action toolbar ribbon matching modern desktop standards."""
        self.toolbar_frame = tk.Frame(
            self.root,
            bg=self.colors["toolbar_bg"],
            highlightbackground=self.colors["toolbar_border"],
            highlightthickness=1,
            height=48,
        )
        self.toolbar_frame.pack(side="top", fill="x")
        self.toolbar_frame.pack_propagate(False)

        # Left action group
        left_group = tk.Frame(self.toolbar_frame, bg=self.colors["toolbar_bg"])
        left_group.pack(side="left", fill="y", padx=12)

        self.btn_open = self.create_toolbar_button(left_group, "📂 Open File", self.on_open_file_dialog)
        self.btn_reload = self.create_toolbar_button(left_group, "🔄 Reload", self.reload_active_file)

        # Separator
        self.create_toolbar_separator(left_group)

        # Primary actions
        self.btn_shuffle = self.create_toolbar_button(
            left_group, "🔀 Shuffle", self.shuffle_manual,
            bg=self.colors["primary"], fg=self.colors["primary_fg"],
            hover=self.colors["primary_hover"]
        )
        self.btn_copy = self.create_toolbar_button(
            left_group, "📋 Copy Results", self.copy_result_to_clipboard,
            bg=self.colors["success"], fg=self.colors["success_fg"],
            hover=self.colors["success_hover"]
        )
        self.btn_export = self.create_toolbar_button(left_group, "💾 Export...", self.export_results)

        # Right group (Theme toggle & File status badge)
        right_group = tk.Frame(self.toolbar_frame, bg=self.colors["toolbar_bg"])
        right_group.pack(side="right", fill="y", padx=12)

        # Theme toggle button
        theme_icon = "🌙 Dark" if self.theme_name == "light" else "☀️ Light"
        self.btn_theme = self.create_toolbar_button(right_group, theme_icon, self.toggle_theme)

        # Separator
        self.create_toolbar_separator(right_group)

        # Dataset Status Pill / Badge
        self.dataset_pill = tk.Label(
            right_group,
            textvariable=self.file_badge_var,
            bg=self.colors["badge_bg"],
            fg=self.colors["badge_fg"],
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=4,
            relief="solid",
            bd=1,
            highlightthickness=0,
            cursor="hand2"
        )
        self.dataset_pill.pack(side="left", pady=8)
        self.dataset_pill.bind("<Button-1>", lambda _: self.open_file_location())

    def create_toolbar_button(
        self, parent: tk.Widget, text: str, command,
        bg: str | None = None, fg: str | None = None, hover: str | None = None
    ) -> tk.Button:
        btn_bg = bg or self.colors["secondary_btn"]
        btn_fg = fg or self.colors["secondary_btn_fg"]
        btn_hover = hover or self.colors["secondary_btn_hover"]

        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=btn_bg,
            fg=btn_fg,
            activebackground=btn_hover,
            activeforeground=btn_fg,
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
            highlightthickness=1,
            highlightbackground=self.colors["card_border"],
        )
        btn.pack(side="left", padx=3, pady=8)
        btn.bind("<Enter>", lambda _, b=btn, h=btn_hover: b.configure(bg=h))
        btn.bind("<Leave>", lambda _, b=btn, orig=btn_bg: b.configure(bg=orig))
        if not bg:
            self.secondary_buttons.append(btn)
        return btn

    def create_toolbar_separator(self, parent: tk.Widget) -> None:
        sep = tk.Frame(parent, bg=self.colors["toolbar_border"], width=1, height=22)
        sep.pack(side="left", padx=8, pady=12)

    # ------------------------------------------------------------------------
    # Main Body Layout (Left Controls, Right Results & Tools)
    # ------------------------------------------------------------------------

    def build_main_layout(self) -> None:
        self.main_container = tk.Frame(self.root, bg=self.colors["bg"])
        self.main_container.pack(side="top", fill="both", expand=True, padx=14, pady=10)

        self.main_container.grid_columnconfigure(0, weight=0, minsize=340)  # Left panel
        self.main_container.grid_columnconfigure(1, weight=1)               # Right panel
        self.main_container.grid_rowconfigure(0, weight=1)

        # Build Left Panel
        self.left_panel = tk.Frame(self.main_container, bg=self.colors["bg"], width=340)
        self.left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.left_panel.grid_propagate(False)

        self.build_amount_card(self.left_panel)
        self.build_format_card(self.left_panel)
        self.build_file_info_card(self.left_panel)

        # Build Right Panel
        self.right_panel = tk.Frame(self.main_container, bg=self.colors["bg"])
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self.right_panel.grid_rowconfigure(0, weight=3)  # Result card
        self.right_panel.grid_rowconfigure(1, weight=2)  # Tools notebook
        self.right_panel.grid_columnconfigure(0, weight=1)

        self.build_result_card(self.right_panel)
        self.build_tools_card(self.right_panel)

    def create_card(self, parent: tk.Widget) -> tk.Frame:
        """Create a rounded-feel modern card container."""
        card = tk.Frame(
            parent,
            bg=self.colors["card_bg"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        self.cards.append(card)
        return card

    # ------------------------------------------------------------------------
    # Left Panel Cards: Pick Amount, Format, File Info
    # ------------------------------------------------------------------------

    def build_amount_card(self, parent: tk.Widget) -> None:
        card = self.create_card(parent)
        card.pack(side="top", fill="x", pady=(0, 10))

        # Title & Subtitle
        title_row = tk.Frame(card, bg=self.colors["card_bg"])
        title_row.pack(fill="x")
        self.cards.append(title_row)

        lbl_amt = tk.Label(
            title_row, text="Pick Amount",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 12, "bold")
        )
        lbl_amt.pack(side="left")
        self.card_labels.append(lbl_amt)

        lbl_sub = tk.Label(
            card, text="Type number for instant auto-shuffle:",
            bg=self.colors["card_bg"], fg=self.colors["text_muted"],
            font=("Segoe UI", 9)
        )
        lbl_sub.pack(anchor="w", pady=(2, 6))
        self.muted_labels.append(lbl_sub)

        # Big Number Entry Box
        self.count_entry = tk.Entry(
            card,
            textvariable=self.count_var,
            justify="center",
            bg=self.colors["input_bg"],
            fg=self.colors["input_fg"],
            insertbackground=self.colors["text"],
            font=("Segoe UI", 24, "bold"),
            relief="solid",
            bd=1,
            highlightthickness=2,
            highlightbackground=self.colors["input_border"],
            highlightcolor=self.colors["input_focus"],
        )
        self.count_entry.pack(fill="x", ipady=4, pady=(0, 8))

        # Quick Count Chips Row
        chips_frame = tk.Frame(card, bg=self.colors["card_bg"])
        chips_frame.pack(fill="x", pady=(0, 10))
        self.cards.append(chips_frame)

        for chip_val in ["1", "5", "10", "25", "50", "All"]:
            btn = tk.Button(
                chips_frame,
                text=chip_val,
                command=lambda v=chip_val: self.set_quick_count(v),
                bg=self.colors["chip_bg"],
                fg=self.colors["chip_fg"],
                activebackground=self.colors["chip_hover"],
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI", 8, "bold"),
                padx=6,
                pady=2,
            )
            btn.pack(side="left", expand=True, fill="x", padx=1)
            btn.bind("<Enter>", lambda _, b=btn: b.configure(bg=self.colors["chip_hover"]))
            btn.bind("<Leave>", lambda _, b=btn: b.configure(bg=self.colors["chip_bg"]))
            self.chip_buttons.append(btn)

        # Shuffle Mode Selector
        mode_frame = tk.Frame(card, bg=self.colors["card_bg"])
        mode_frame.pack(fill="x", pady=(0, 8))
        self.cards.append(mode_frame)

        lbl_mode = tk.Label(
            mode_frame, text="Shuffle Mode:",
            bg=self.colors["card_bg"], fg=self.colors["text_secondary"],
            font=("Segoe UI", 9, "bold")
        )
        lbl_mode.pack(anchor="w")
        self.card_labels.append(lbl_mode)

        r1 = tk.Radiobutton(
            mode_frame, text="Unique Picks (No Repeats)",
            variable=self.mode_var, value="unique",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            selectcolor=self.colors["input_bg"],
            activebackground=self.colors["card_bg"],
            font=("Segoe UI", 9)
        )
        r1.pack(anchor="w", pady=1)
        self.radio_buttons.append(r1)

        r2 = tk.Radiobutton(
            mode_frame, text="With Replacement (Allow Repeats)",
            variable=self.mode_var, value="replacement",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            selectcolor=self.colors["input_bg"],
            activebackground=self.colors["card_bg"],
            font=("Segoe UI", 9)
        )
        r2.pack(anchor="w", pady=1)
        self.radio_buttons.append(r2)

        elim_row = tk.Frame(mode_frame, bg=self.colors["card_bg"])
        elim_row.pack(fill="x")
        self.cards.append(elim_row)

        r3 = tk.Radiobutton(
            elim_row, text="Elimination (Giveaway / Raffle)",
            variable=self.mode_var, value="elimination",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            selectcolor=self.colors["input_bg"],
            activebackground=self.colors["card_bg"],
            font=("Segoe UI", 9)
        )
        r3.pack(side="left")
        self.radio_buttons.append(r3)

        self.btn_reset_pool = tk.Button(
            elim_row, text="↺ Reset",
            command=self.reset_elimination_pool,
            bg=self.colors["chip_bg"], fg=self.colors["text_secondary"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 8, "bold"),
            padx=4, pady=1
        )
        self.btn_reset_pool.pack(side="right")
        self.chip_buttons.append(self.btn_reset_pool)

        # Big Re-Shuffle Button
        self.card_shuffle_btn = tk.Button(
            card,
            text="🔀  SHUFFLE NOW",
            command=self.shuffle_manual,
            bg=self.colors["primary"],
            fg=self.colors["primary_fg"],
            activebackground=self.colors["primary_hover"],
            activeforeground=self.colors["primary_fg"],
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            pady=7,
        )
        self.card_shuffle_btn.pack(fill="x", pady=(4, 0))

    def build_format_card(self, parent: tk.Widget) -> None:
        card = self.create_card(parent)
        card.pack(side="top", fill="x", pady=(0, 10))

        lbl_fmt = tk.Label(
            card, text="Output Formatting",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 11, "bold")
        )
        lbl_fmt.pack(anchor="w", pady=(0, 6))
        self.card_labels.append(lbl_fmt)

        # Delimiter Selector
        delims_frame = tk.Frame(card, bg=self.colors["card_bg"])
        delims_frame.pack(fill="x", pady=(0, 4))
        self.cards.append(delims_frame)

        lbl_sep = tk.Label(
            delims_frame, text="Separator:",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 9, "bold")
        )
        lbl_sep.pack(side="left")
        self.card_labels.append(lbl_sep)

        delim_combo = ttk.Combobox(
            delims_frame,
            textvariable=self.delimiter_var,
            values=["newline", "space", "comma", "semicolon", "custom"],
            state="readonly",
            width=13,
            font=("Segoe UI", 9, "bold"),
        )
        delim_combo.pack(side="right")

        # Prefix & Suffix in two columns
        affix_row = tk.Frame(card, bg=self.colors["card_bg"])
        affix_row.pack(fill="x", pady=4)
        affix_row.grid_columnconfigure(0, weight=1)
        affix_row.grid_columnconfigure(1, weight=1)
        self.cards.append(affix_row)

        p_frame = tk.Frame(affix_row, bg=self.colors["card_bg"])
        p_frame.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.cards.append(p_frame)

        lbl_p = tk.Label(p_frame, text="Prefix (e.g. @):", bg=self.colors["card_bg"], fg=self.colors["text_secondary"], font=("Segoe UI", 8, "bold"))
        lbl_p.pack(anchor="w")
        self.card_labels.append(lbl_p)

        self.entry_prefix = tk.Entry(
            p_frame, textvariable=self.prefix_var,
            bg=self.colors["input_bg"], fg=self.colors["input_fg"],
            insertbackground=self.colors["text"],
            relief="solid", bd=1,
            highlightthickness=1,
            highlightbackground=self.colors["input_border"],
            highlightcolor=self.colors["input_focus"],
            font=("Segoe UI", 9, "bold")
        )
        self.entry_prefix.pack(fill="x", ipady=3, pady=(2, 0))

        s_frame = tk.Frame(affix_row, bg=self.colors["card_bg"])
        s_frame.grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.cards.append(s_frame)

        lbl_s = tk.Label(s_frame, text="Suffix:", bg=self.colors["card_bg"], fg=self.colors["text_secondary"], font=("Segoe UI", 8, "bold"))
        lbl_s.pack(anchor="w")
        self.card_labels.append(lbl_s)

        self.entry_suffix = tk.Entry(
            s_frame, textvariable=self.suffix_var,
            bg=self.colors["input_bg"], fg=self.colors["input_fg"],
            insertbackground=self.colors["text"],
            relief="solid", bd=1,
            highlightthickness=1,
            highlightbackground=self.colors["input_border"],
            highlightcolor=self.colors["input_focus"],
            font=("Segoe UI", 9, "bold")
        )
        self.entry_suffix.pack(fill="x", ipady=3, pady=(2, 0))

        # Numbered & Case
        opt_row = tk.Frame(card, bg=self.colors["card_bg"])
        opt_row.pack(fill="x", pady=(4, 0))
        self.cards.append(opt_row)

        chk_num = tk.Checkbutton(
            opt_row, text="Numbered (1., 2.)",
            variable=self.numbered_var,
            bg=self.colors["card_bg"], fg=self.colors["text"],
            selectcolor=self.colors["input_bg"],
            activebackground=self.colors["card_bg"],
            font=("Segoe UI", 9, "bold")
        )
        chk_num.pack(side="left")
        self.radio_buttons.append(chk_num)

        case_combo = ttk.Combobox(
            opt_row,
            textvariable=self.case_var,
            values=["original", "lowercase", "UPPERCASE", "Title Case"],
            state="readonly",
            width=12,
            font=("Segoe UI", 9, "bold")
        )
        case_combo.pack(side="right")

    def build_file_info_card(self, parent: tk.Widget) -> None:
        card = self.create_card(parent)
        card.pack(side="top", fill="both", expand=True)

        header_row = tk.Frame(card, bg=self.colors["card_bg"])
        header_row.pack(fill="x", pady=(0, 4))
        self.cards.append(header_row)

        lbl_info = tk.Label(
            header_row, text="Dataset Info",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 11, "bold")
        )
        lbl_info.pack(side="left")
        self.card_labels.append(lbl_info)

        btn_browse = tk.Button(
            header_row, text="Browse...", command=self.on_open_file_dialog,
            bg=self.colors["chip_bg"], fg=self.colors["chip_fg"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 8, "bold"),
            padx=6, pady=2
        )
        btn_browse.pack(side="right")
        self.chip_buttons.append(btn_browse)

        # File metrics
        self.metric_label(card, "Active File:", self.file_name_var)
        self.metric_label(card, "Total Pool:", self.total_count_var)
        self.metric_label(card, "Remaining in Pool:", self.remaining_pool_var)

        # Multi-column selector
        col_frame = tk.Frame(card, bg=self.colors["card_bg"])
        col_frame.pack(fill="x", pady=(6, 0))
        self.cards.append(col_frame)

        lbl_col = tk.Label(
            col_frame, text="Active Column:",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 9, "bold")
        )
        lbl_col.pack(side="left")
        self.card_labels.append(lbl_col)

        self.col_combo = ttk.Combobox(
            col_frame,
            textvariable=self.col_var,
            values=self.columns,
            state="readonly",
            width=15,
            font=("Segoe UI", 9, "bold")
        )
        self.col_combo.pack(side="right")
        self.col_combo.bind("<<ComboboxSelected>>", self.on_column_changed)

    def metric_label(self, parent: tk.Widget, title: str, var: tk.StringVar) -> None:
        row = tk.Frame(parent, bg=self.colors["card_bg"])
        row.pack(fill="x", pady=2)
        self.cards.append(row)

        lbl_t = tk.Label(
            row, text=title,
            bg=self.colors["card_bg"], fg=self.colors["text_muted"],
            font=("Segoe UI", 9)
        )
        lbl_t.pack(side="left")
        self.muted_labels.append(lbl_t)

        lbl_v = tk.Label(
            row, textvariable=var,
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 9, "bold")
        )
        lbl_v.pack(side="right")
        self.card_labels.append(lbl_v)

    # ------------------------------------------------------------------------
    # Right Panel: Results Card & Tools Notebook
    # ------------------------------------------------------------------------

    def build_result_card(self, parent: tk.Widget) -> None:
        card = self.create_card(parent)
        card.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        # Result Header Row
        header = tk.Frame(card, bg=self.colors["card_bg"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.grid_columnconfigure(1, weight=1)
        self.cards.append(header)

        lbl_res = tk.Label(
            header, text="Shuffled Results",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 12, "bold")
        )
        lbl_res.pack(side="left")
        self.card_labels.append(lbl_res)

        # Picked count pill
        self.picked_pill = tk.Label(
            header, textvariable=self.picked_count_var,
            bg=self.colors["badge_bg"], fg=self.colors["badge_fg"],
            font=("Segoe UI", 9, "bold"), padx=8, pady=2
        )
        self.picked_pill.pack(side="left", padx=8)

        # Quick action buttons on right of header
        actions = tk.Frame(header, bg=self.colors["card_bg"])
        actions.pack(side="right")
        self.cards.append(actions)

        self.btn_copy_card = tk.Button(
            actions, text="📋 Copy",
            command=self.copy_result_to_clipboard,
            bg=self.colors["success"], fg=self.colors["success_fg"],
            activebackground=self.colors["success_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9, "bold"),
            padx=12, pady=4
        )
        self.btn_copy_card.pack(side="left", padx=3)

        btn_exp = tk.Button(
            actions, text="💾 Export",
            command=self.export_results,
            bg=self.colors["secondary_btn"], fg=self.colors["secondary_btn_fg"],
            activebackground=self.colors["secondary_btn_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9),
            padx=8, pady=4
        )
        btn_exp.pack(side="left", padx=3)
        self.secondary_buttons.append(btn_exp)

        btn_clr = tk.Button(
            actions, text="🗑 Clear",
            command=self.clear_results,
            bg=self.colors["secondary_btn"], fg=self.colors["secondary_btn_fg"],
            activebackground=self.colors["secondary_btn_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9),
            padx=8, pady=4
        )
        btn_clr.pack(side="left", padx=3)
        self.secondary_buttons.append(btn_clr)

        # Output Text Box Area with Scrollbar
        self.box_frame = tk.Frame(card, bg=self.colors["result_border"])
        self.box_frame.grid(row=1, column=0, sticky="nsew")
        self.box_frame.grid_rowconfigure(0, weight=1)
        self.box_frame.grid_columnconfigure(0, weight=1)

        self.output_box = tk.Text(
            self.box_frame,
            wrap="word",
            bg=self.colors["result_bg"],
            fg=self.colors["result_fg"],
            relief="flat",
            font=("Consolas", 12),
            padx=12,
            pady=10,
            selectbackground=self.colors["primary"],
            selectforeground=self.colors["primary_fg"],
        )
        self.output_box.grid(row=0, column=0, sticky="nsew")

        scrollbar = tk.Scrollbar(self.box_frame, orient="vertical", command=self.output_box.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.output_box.configure(yscrollcommand=scrollbar.set, state="disabled")

    def build_tools_card(self, parent: tk.Widget) -> None:
        """Tabs for: Add Usernames, Search Dataset, and History Log."""
        card = self.create_card(parent)
        card.grid(row=1, column=0, sticky="nsew")
        card.grid_rowconfigure(0, weight=1)
        card.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(card)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Tab 1: Add Usernames
        self.tab_add = tk.Frame(self.notebook, bg=self.colors["card_bg"], padx=10, pady=8)
        self.notebook.add(self.tab_add, text=" ➕ Add Usernames ")
        self.cards.append(self.tab_add)
        self.build_add_tab(self.tab_add)

        # Tab 2: Search Dataset
        self.tab_search = tk.Frame(self.notebook, bg=self.colors["card_bg"], padx=10, pady=8)
        self.notebook.add(self.tab_search, text=" 🔍 Search Dataset ")
        self.cards.append(self.tab_search)
        self.build_search_tab(self.tab_search)

        # Tab 3: History Log
        self.tab_history = tk.Frame(self.notebook, bg=self.colors["card_bg"], padx=10, pady=8)
        self.notebook.add(self.tab_history, text=" 📜 Draw History ")
        self.cards.append(self.tab_history)
        self.build_history_tab(self.tab_history)

    def build_add_tab(self, parent: tk.Widget) -> None:
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        main_box = tk.Frame(parent, bg=self.colors["card_bg"])
        main_box.grid(row=0, column=0, sticky="nsew")
        main_box.grid_rowconfigure(1, weight=1)
        main_box.grid_columnconfigure(0, weight=1)
        self.cards.append(main_box)

        lbl_inst = tk.Label(
            main_box,
            text="Paste new usernames below (separated by lines, commas, or semicolons):",
            bg=self.colors["card_bg"], fg=self.colors["text_muted"],
            font=("Segoe UI", 9)
        )
        lbl_inst.grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.muted_labels.append(lbl_inst)

        self.add_text = tk.Text(
            main_box,
            height=3,
            wrap="word",
            bg=self.colors["input_bg"],
            fg=self.colors["input_fg"],
            relief="solid",
            bd=1,
            font=("Segoe UI", 9),
            padx=8,
            pady=6,
        )
        self.add_text.grid(row=1, column=0, sticky="nsew")

        btn_row = tk.Frame(main_box, bg=self.colors["card_bg"])
        btn_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        self.cards.append(btn_row)

        self.btn_save_add = tk.Button(
            btn_row, text="💾 Save & Append to File",
            command=self.save_new_usernames,
            bg=self.colors["warning"], fg=self.colors["warning_fg"],
            activebackground=self.colors["warning_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9, "bold"),
            padx=12, pady=4
        )
        self.btn_save_add.pack(side="left")

        btn_clear_add = tk.Button(
            btn_row, text="Clear Box",
            command=lambda: self.add_text.delete("1.0", tk.END),
            bg=self.colors["secondary_btn"], fg=self.colors["secondary_btn_fg"],
            activebackground=self.colors["secondary_btn_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9),
            padx=8, pady=4
        )
        btn_clear_add.pack(side="left", padx=8)
        self.secondary_buttons.append(btn_clear_add)

    def build_search_tab(self, parent: tk.Widget) -> None:
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        search_bar = tk.Frame(parent, bg=self.colors["card_bg"])
        search_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.cards.append(search_bar)

        lbl_s = tk.Label(
            search_bar, text="Search Username:",
            bg=self.colors["card_bg"], fg=self.colors["text"],
            font=("Segoe UI", 9, "bold")
        )
        lbl_s.pack(side="left")
        self.card_labels.append(lbl_s)

        self.search_entry = tk.Entry(
            search_bar,
            bg=self.colors["input_bg"], fg=self.colors["input_fg"],
            relief="solid", bd=1, font=("Segoe UI", 9)
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=8)
        self.search_entry.bind("<KeyRelease>", self.on_search_key)

        self.search_count_lbl = tk.Label(
            search_bar, text="0 matches",
            bg=self.colors["card_bg"], fg=self.colors["text_muted"],
            font=("Segoe UI", 8)
        )
        self.search_count_lbl.pack(side="right")
        self.muted_labels.append(self.search_count_lbl)

        self.search_results_list = tk.Listbox(
            parent,
            bg=self.colors["input_bg"], fg=self.colors["input_fg"],
            relief="solid", bd=1, font=("Consolas", 10),
            selectbackground=self.colors["primary"],
            selectforeground=self.colors["primary_fg"],
        )
        self.search_results_list.grid(row=1, column=0, sticky="nsew")

    def build_history_tab(self, parent: tk.Widget) -> None:
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        history_box_frame = tk.Frame(parent, bg=self.colors["card_bg"])
        history_box_frame.grid(row=0, column=0, sticky="nsew")
        history_box_frame.grid_rowconfigure(0, weight=1)
        history_box_frame.grid_columnconfigure(0, weight=1)
        self.cards.append(history_box_frame)

        self.history_listbox = tk.Listbox(
            history_box_frame,
            bg=self.colors["input_bg"], fg=self.colors["input_fg"],
            relief="solid", bd=1, font=("Segoe UI", 9),
            selectbackground=self.colors["primary"],
            selectforeground=self.colors["primary_fg"],
        )
        self.history_listbox.grid(row=0, column=0, sticky="nsew")

        btn_row = tk.Frame(parent, bg=self.colors["card_bg"])
        btn_row.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.cards.append(btn_row)

        btn_load_hist = tk.Button(
            btn_row, text="📋 Load Selected Draw to Result Area",
            command=self.load_selected_history_draw,
            bg=self.colors["secondary_btn"], fg=self.colors["secondary_btn_fg"],
            activebackground=self.colors["secondary_btn_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 9, "bold"),
            padx=10, pady=3
        )
        btn_load_hist.pack(side="left")
        self.secondary_buttons.append(btn_load_hist)

        btn_clear_hist = tk.Button(
            btn_row, text="Clear History",
            command=self.clear_history,
            bg=self.colors["secondary_btn"], fg=self.colors["secondary_btn_fg"],
            activebackground=self.colors["secondary_btn_hover"],
            relief="flat", bd=0, cursor="hand2", font=("Segoe UI", 8),
            padx=6, pady=3
        )
        btn_clear_hist.pack(side="right")
        self.secondary_buttons.append(btn_clear_hist)

    # ------------------------------------------------------------------------
    # Status Bar
    # ------------------------------------------------------------------------

    def build_status_bar(self) -> None:
        self.status_bar = tk.Frame(
            self.root,
            bg=self.colors["status_bg"],
            highlightbackground=self.colors["status_border"],
            highlightthickness=1,
            height=30,
        )
        self.status_bar.pack(side="bottom", fill="x")
        self.status_bar.pack_propagate(False)

        self.status_label = tk.Label(
            self.status_bar,
            textvariable=self.status_var,
            bg=self.colors["status_bg"],
            fg=self.colors["status_fg"],
            font=("Segoe UI", 9),
            anchor="w",
            padx=12,
        )
        self.status_label.pack(side="left", fill="both", expand=True)

        # Version Pill
        ver_pill = tk.Label(
            self.status_bar,
            text=f"{APP_NAME} {APP_VERSION}",
            bg=self.colors["status_bg"],
            fg=self.colors["text_muted"],
            font=("Segoe UI", 8),
            padx=12,
        )
        ver_pill.pack(side="right")

    # ------------------------------------------------------------------------
    # Theme & Dynamic Styling Engine
    # ------------------------------------------------------------------------

    def toggle_theme(self) -> None:
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self.colors = THEMES[self.theme_name]
        self.config.set("theme", self.theme_name)
        self.apply_theme_colors()
        self.set_status(f"Switched to {self.theme_name.capitalize()} Mode.")

    def apply_theme_colors(self) -> None:
        c = self.colors
        self.root.configure(bg=c["bg"])
        self.main_container.configure(bg=c["bg"])
        self.left_panel.configure(bg=c["bg"])
        self.right_panel.configure(bg=c["bg"])

        # Toolbar
        self.toolbar_frame.configure(bg=c["toolbar_bg"], highlightbackground=c["toolbar_border"])
        self.btn_theme.configure(text="🌙 Dark" if self.theme_name == "light" else "☀️ Light")
        self.dataset_pill.configure(bg=c["badge_bg"], fg=c["badge_fg"])

        # Cards
        for card in self.cards:
            try:
                card.configure(bg=c["card_bg"])
                if "highlightbackground" in card.keys():
                    card.configure(highlightbackground=c["card_border"])
            except Exception:
                pass

        for lbl in self.card_labels:
            try:
                lbl.configure(bg=c["card_bg"], fg=c["text"])
            except Exception:
                pass

        for lbl in self.muted_labels:
            try:
                lbl.configure(bg=c["card_bg"], fg=c["text_muted"])
            except Exception:
                pass

        for btn in self.secondary_buttons:
            try:
                btn.configure(
                    bg=c["secondary_btn"], fg=c["secondary_btn_fg"],
                    activebackground=c["secondary_btn_hover"]
                )
            except Exception:
                pass

        for btn in self.chip_buttons:
            try:
                btn.configure(
                    bg=c["chip_bg"], fg=c["chip_fg"],
                    activebackground=c["chip_hover"]
                )
            except Exception:
                pass

        for rb in self.radio_buttons:
            try:
                rb.configure(
                    bg=c["card_bg"], fg=c["text"],
                    selectcolor=c["input_bg"], activebackground=c["card_bg"]
                )
            except Exception:
                pass

        # Text Entries
        for entry in [self.count_entry, self.entry_prefix, self.entry_suffix, self.search_entry]:
            try:
                entry.configure(
                    bg=c["input_bg"], fg=c["input_fg"],
                    insertbackground=c["text"]
                )
            except Exception:
                pass

        # Text Areas & Lists
        for widget in [self.add_text, self.search_results_list, self.history_listbox]:
            try:
                widget.configure(
                    bg=c["input_bg"], fg=c["input_fg"],
                    selectbackground=c["primary"], selectforeground=c["primary_fg"]
                )
            except Exception:
                pass

        # Output box
        self.box_frame.configure(bg=c["result_border"])
        self.output_box.configure(
            bg=c["result_bg"], fg=c["result_fg"],
            selectbackground=c["primary"], selectforeground=c["primary_fg"]
        )

        # Status Bar
        self.status_bar.configure(bg=c["status_bg"], highlightbackground=c["status_border"])
        self.status_label.configure(bg=c["status_bg"], fg=c["status_fg"])

        # TTK styles
        self.setup_ttk_styles()

    # ------------------------------------------------------------------------
    # Core Logic: Data Loading & Threading
    # ------------------------------------------------------------------------

    def load_file_async(self, path: str | None, *, initial: bool = False) -> None:
        """Asynchronously load file without freezing GUI."""
        if not path or not os.path.exists(path):
            self.usernames = []
            self.elimination_pool = []
            self.loaded_file = None
            self.update_metrics()
            self.set_status("No dataset loaded. Click 'Open File' to select an Excel or CSV file.")
            self.file_badge_var.set("No file loaded")
            return

        self.set_status(f"Loading {os.path.basename(path)}...")
        self.file_badge_var.set(f"Loading {os.path.basename(path)}...")

        def worker() -> None:
            t0 = time.perf_counter()
            try:
                cols, active_idx, names, has_hdr = load_dataset_table(path)
                elapsed_ms = (time.perf_counter() - t0) * 1000
                error = None
            except Exception as exc:
                cols, active_idx, names, has_hdr, elapsed_ms = ["Column 1"], 0, [], False, 0.0
                error = str(exc)

            self.task_queue.put(lambda: self._on_load_finished(
                path, cols, active_idx, names, has_hdr, elapsed_ms, error, initial
            ))

        threading.Thread(target=worker, daemon=True).start()

    def _on_load_finished(
        self, path: str, cols: list[str], active_idx: int,
        names: list[str], has_hdr: bool, elapsed_ms: float,
        error: str | None, initial: bool
    ) -> None:
        if error:
            self.set_status(f"Error loading file: {error}", error=True)
            self.file_badge_var.set("Error loading file")
            messagebox.showerror("File Load Error", f"Could not read {os.path.basename(path)}:\n{error}")
            return

        self.loaded_file = path
        self.columns = cols
        self.active_col_index = active_idx
        self.usernames = names
        self.has_header = has_hdr
        self.elimination_pool = list(names)

        # Update combobox
        self.col_combo.configure(values=self.columns)
        if 0 <= active_idx < len(self.columns):
            self.col_var.set(self.columns[active_idx])

        self.config.add_recent_file(path)
        self.update_recent_menu()
        self.update_metrics()

        header_info = " (header row skipped)" if has_hdr else ""
        self.set_status(
            f"✓ Loaded {len(names)} usernames from {os.path.basename(path)} in {elapsed_ms:.1f}ms{header_info}."
        )

        if initial:
            self.count_entry.focus_set()

        # Trigger shuffle if count is already entered
        if self.count_var.get().strip():
            self.shuffle_names(automatic=True)

    def reload_active_file(self) -> None:
        if self.loaded_file and os.path.exists(self.loaded_file):
            self.load_file_async(self.loaded_file, initial=False)
        else:
            self.on_open_file_dialog()

    def on_open_file_dialog(self) -> None:
        initial_dir = os.path.dirname(self.loaded_file) if self.loaded_file else self.folder
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title="Open Username Dataset",
            initialdir=initial_dir,
            filetypes=[
                ("All Supported Datasets", "*.xlsx *.xlsm *.csv *.txt"),
                ("Excel Files", "*.xlsx *.xlsm"),
                ("CSV Files", "*.csv"),
                ("Text Files", "*.txt"),
                ("All Files", "*.*")
            ]
        )
        if chosen:
            self.load_file_async(chosen, initial=False)

    def on_column_changed(self, _event=None) -> None:
        chosen_col = self.col_var.get()
        if chosen_col in self.columns and self.loaded_file:
            idx = self.columns.index(chosen_col)
            self.active_col_index = idx
            try:
                # Read all rows and extract column idx
                lower = self.loaded_file.lower()
                if lower.endswith((".xlsx", ".xlsm")):
                    rows = parse_xlsx_stream(self.loaded_file)
                elif lower.endswith(".csv"):
                    rows = parse_csv_stream(self.loaded_file)
                else:
                    rows = parse_txt_stream(self.loaded_file)

                data_rows = rows[1:] if self.has_header else rows
                names = [r[idx] for r in data_rows if idx < len(r) and r[idx]]

                self.usernames = names
                self.elimination_pool = list(names)
                self.update_metrics()
                self.set_status(f"Switched to column '{chosen_col}' ({len(names)} usernames).")
                if self.count_var.get().strip():
                    self.shuffle_names(automatic=True)
            except Exception as e:
                self.set_status(f"Error switching column: {e}", error=True)

    def open_file_location(self) -> None:
        if self.loaded_file and os.path.exists(self.loaded_file):
            target_dir = os.path.dirname(os.path.abspath(self.loaded_file))
        else:
            target_dir = self.folder
        try:
            if os.name == "nt":
                os.startfile(target_dir)
            else:
                import subprocess
                subprocess.Popen(["xdg-open", target_dir])
        except Exception as e:
            self.set_status(f"Could not open directory: {e}", error=True)

    # ------------------------------------------------------------------------
    # Core Logic: Shuffle Engine & Modes
    # ------------------------------------------------------------------------

    def _on_count_change(self, *_) -> None:
        if not self._trace_active:
            return
        if self.auto_shuffle_var.get():
            self.shuffle_names(automatic=True)

    def _on_mode_change(self, *_) -> None:
        self.config.set("mode", self.mode_var.get())
        if self.count_var.get().strip():
            self.shuffle_names(automatic=True)

    def set_quick_count(self, value: str) -> None:
        if value == "All":
            self.count_var.set(str(len(self.usernames)))
        else:
            self.count_var.set(value)
        self.shuffle_names(automatic=False)

    def shuffle_manual(self) -> None:
        self.shuffle_names(automatic=False)

    def shuffle_names(self, *, automatic: bool) -> None:
        raw_val = self.count_var.get().strip()

        if not raw_val:
            self.last_shuffled_list = []
            self.set_output("")
            self.picked_count_var.set("0")
            if self.usernames:
                self.set_status(f"Ready. {len(self.usernames)} usernames available.")
            return

        if not raw_val.isdigit():
            self.set_output("")
            self.picked_count_var.set("0")
            self.set_status("Please enter a valid positive number.", error=not automatic)
            return

        amount = int(raw_val)
        if amount <= 0:
            self.set_output("")
            self.picked_count_var.set("0")
            self.set_status("Enter a number greater than 0.", error=not automatic)
            return

        if not self.usernames:
            self.set_output("")
            self.picked_count_var.set("0")
            self.set_status("No usernames loaded. Open a file or add names below.", error=not automatic)
            return

        mode = self.mode_var.get()
        t0 = time.perf_counter()

        if mode == "unique":
            if amount > len(self.usernames):
                self.set_output("")
                self.picked_count_var.set("0")
                self.set_status(
                    f"Pick count ({amount}) exceeds pool ({len(self.usernames)}). Switch to 'With Replacement' or pick fewer.",
                    error=not automatic
                )
                return
            selected = random.sample(self.usernames, amount)

        elif mode == "replacement":
            selected = [random.choice(self.usernames) for _ in range(amount)]

        elif mode == "elimination":
            if not self.elimination_pool:
                self.set_output("")
                self.picked_count_var.set("0")
                self.set_status("Elimination pool exhausted! Click 'Reset Pool' to restart.", error=True)
                return

            pick_n = min(amount, len(self.elimination_pool))
            selected = random.sample(self.elimination_pool, pick_n)
            # Remove picked from pool
            for u in selected:
                self.elimination_pool.remove(u)
            self.remaining_pool_var.set(str(len(self.elimination_pool)))

            if len(selected) < amount:
                self.set_status(f"Pool only had {len(selected)} items remaining. All remaining picked!")
        else:
            selected = random.sample(self.usernames, min(amount, len(self.usernames)))

        elapsed_ms = (time.perf_counter() - t0) * 1000
        self.last_shuffled_list = selected
        self.reformat_current_result()

        # Record history
        self.record_history_entry(selected, mode)
        self.set_status(f"🔀 Shuffled {len(selected)} username{'s' if len(selected) != 1 else ''} in {elapsed_ms:.2f}ms.")

    def reset_elimination_pool(self) -> None:
        self.elimination_pool = list(self.usernames)
        self.remaining_pool_var.set(str(len(self.elimination_pool)))
        self.set_status(f"Elimination pool reset to {len(self.elimination_pool)} usernames.")

    # ------------------------------------------------------------------------
    # Formatting & Output Rendering
    # ------------------------------------------------------------------------

    def reformat_current_result(self) -> None:
        if not self.last_shuffled_list:
            self.set_output("")
            self.picked_count_var.set("0")
            return

        names = list(self.last_shuffled_list)
        prefix = self.prefix_var.get()
        suffix = self.suffix_var.get()
        case_mode = self.case_var.get()
        numbered = self.numbered_var.get()
        delim_mode = self.delimiter_var.get()

        # 1. Case transformation
        if case_mode == "lowercase":
            names = [n.lower() for n in names]
        elif case_mode == "UPPERCASE":
            names = [n.upper() for n in names]
        elif case_mode == "Title Case":
            names = [n.title() for n in names]

        # 2. Prefix & Suffix (smart prefix avoid double @@)
        formatted_names = []
        for n in names:
            item = n
            if prefix:
                if prefix == "@" and item.startswith("@"):
                    pass
                else:
                    item = f"{prefix}{item}"
            if suffix:
                item = f"{item}{suffix}"
            formatted_names.append(item)

        # 3. Numbered list
        if numbered:
            formatted_names = [f"{i}. {n}" for i, n in enumerate(formatted_names, start=1)]

        # 4. Delimiter
        if delim_mode == "newline":
            sep = "\n"
        elif delim_mode == "space":
            sep = " "
        elif delim_mode == "comma":
            sep = ", "
        elif delim_mode == "semicolon":
            sep = "; "
        elif delim_mode == "custom":
            sep = self.custom_delimiter_var.get()
        else:
            sep = "\n"

        self.formatted_result = sep.join(formatted_names)
        self.picked_count_var.set(f"{len(formatted_names)} items")
        self.set_output(self.formatted_result)

    def set_output(self, text: str) -> None:
        self.output_box.configure(state="normal")
        self.output_box.delete("1.0", tk.END)
        self.output_box.insert("1.0", text)
        self.output_box.configure(state="disabled")

    def copy_result_to_clipboard(self) -> None:
        if not self.formatted_result:
            self.set_status("Nothing to copy yet. Enter a number to shuffle first.", error=True)
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(self.formatted_result)
        self.root.update_idletasks()

        # Temporary button animation
        orig_text = self.btn_copy_card.cget("text")
        self.btn_copy_card.configure(text="✓ Copied!")
        self.root.after(1500, lambda: self.btn_copy_card.configure(text=orig_text))

        self.set_status(f"✓ Copied {len(self.last_shuffled_list)} usernames to clipboard!")

    def clear_results(self) -> None:
        self.last_shuffled_list = []
        self.formatted_result = ""
        self.set_output("")
        self.count_var.set("")
        self.picked_count_var.set("0")
        self.set_status("Results cleared.")

    def export_results(self) -> None:
        if not self.formatted_result:
            messagebox.showinfo("Export Results", "There are no shuffled results to export. Shuffle first!")
            return

        target_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export Shuffled Usernames",
            defaultextension=".txt",
            filetypes=[
                ("Text File (*.txt)", "*.txt"),
                ("CSV File (*.csv)", "*.csv"),
                ("Excel Spreadsheet (*.xlsx)", "*.xlsx"),
            ]
        )
        if not target_path:
            return

        try:
            lower = target_path.lower()
            if lower.endswith(".xlsx"):
                write_simple_xlsx(target_path, self.last_shuffled_list)
            elif lower.endswith(".csv"):
                with open(target_path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Username"])
                    for name in self.last_shuffled_list:
                        writer.writerow([name])
            else:
                with open(target_path, "w", encoding="utf-8") as f:
                    f.write(self.formatted_result)

            self.set_status(f"✓ Successfully exported results to {os.path.basename(target_path)}!")
            messagebox.showinfo("Export Successful", f"Saved {len(self.last_shuffled_list)} usernames to:\n{target_path}")
        except Exception as exc:
            self.set_status(f"Export error: {exc}", error=True)
            messagebox.showerror("Export Error", f"Failed to export file:\n{exc}")

    # ------------------------------------------------------------------------
    # Dataset Tools: Add Usernames, Search, Draw History
    # ------------------------------------------------------------------------

    def save_new_usernames(self) -> None:
        raw_text = self.add_text.get("1.0", tk.END).strip()
        if not raw_text:
            self.set_status("Please type or paste usernames to add.", error=True)
            return

        # Split on newlines, commas, semicolons
        tokens = re.split(r"[\r\n,;]+", raw_text)
        candidates = [clean_username(t) for t in tokens if clean_username(t)]

        if not candidates:
            self.set_status("No valid usernames found in text.", error=True)
            return

        # Deduplication against current dataset and within batch
        existing_lookup = {u.casefold() for u in self.usernames}
        added: list[str] = []
        skipped = 0

        for name in candidates:
            k = name.casefold()
            if k in existing_lookup:
                skipped += 1
            else:
                existing_lookup.add(k)
                added.append(name)

        if not added:
            self.set_status(f"All {skipped} usernames already exist in dataset.", error=True)
            messagebox.showinfo("No New Usernames", f"All {skipped} usernames are already present in the dataset.")
            return

        target_path = self.loaded_file
        if not target_path or not os.path.exists(target_path):
            target_path = os.path.join(self.folder, "usernames.xlsx")

        self.set_status("Saving new usernames...")

        def worker() -> None:
            try:
                saved_path = append_names_to_file(target_path, added, self.usernames)
                error = None
            except Exception as e:
                saved_path = target_path
                error = str(e)

            self.task_queue.put(lambda: self._on_save_finished(saved_path, added, skipped, error))

        threading.Thread(target=worker, daemon=True).start()

    def _on_save_finished(self, saved_path: str, added: list[str], skipped: int, error: str | None) -> None:
        if error:
            self.set_status(f"Error saving: {error}", error=True)
            messagebox.showerror("Save Error", f"Could not save usernames:\n{error}")
            return

        self.loaded_file = saved_path
        self.usernames.extend(added)
        self.elimination_pool.extend(added)
        self.add_text.delete("1.0", tk.END)
        self.update_metrics()

        msg = f"✓ Added {len(added)} new username{'s' if len(added) != 1 else ''} to {os.path.basename(saved_path)}."
        if skipped > 0:
            msg += f" ({skipped} duplicate{'s' if skipped != 1 else ''} skipped)"
        self.set_status(msg)

        if self.count_var.get().strip():
            self.shuffle_names(automatic=True)

    def on_search_key(self, _event=None) -> None:
        query = self.search_entry.get().strip().casefold()
        self.search_results_list.delete(0, tk.END)
        if not query:
            self.search_count_lbl.configure(text="0 matches")
            return

        matches = [u for u in self.usernames if query in u.casefold()]
        self.search_count_lbl.configure(text=f"{len(matches)} match{'es' if len(matches) != 1 else ''}")
        for m in matches[:100]:  # limit view to 100
            self.search_results_list.insert(tk.END, m)

    def record_history_entry(self, results: list[str], mode: str) -> None:
        now_str = datetime.now().strftime("%H:%M:%S")
        summary = f"[{now_str}] {len(results)} items ({mode}) → {', '.join(results[:3])}"
        if len(results) > 3:
            summary += f" ... (+{len(results) - 3} more)"
        self.draw_history.insert(0, {"time": now_str, "results": results, "mode": mode, "summary": summary})
        self.history_listbox.insert(0, summary)

    def load_selected_history_draw(self) -> None:
        sel = self.history_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self.draw_history):
            entry = self.draw_history[idx]
            self.last_shuffled_list = list(entry["results"])
            self.reformat_current_result()
            self.set_status(f"Loaded draw from {entry['time']} ({len(self.last_shuffled_list)} usernames).")

    def clear_history(self) -> None:
        self.draw_history.clear()
        self.history_listbox.delete(0, tk.END)
        self.set_status("Draw history cleared.")

    # ------------------------------------------------------------------------
    # Shortcuts, Dialogs & Cleanup
    # ------------------------------------------------------------------------

    def bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _: self.on_open_file_dialog())
        self.root.bind("<Control-r>", lambda _: self.reload_active_file())
        self.root.bind("<F5>", lambda _: self.reload_active_file())
        self.root.bind("<Control-Return>", lambda _: self.shuffle_manual())
        self.root.bind("<Control-c>", self._handle_ctrl_c)
        self.root.bind("<Control-Shift-C>", lambda _: self.copy_result_to_clipboard())
        self.root.bind("<Control-e>", lambda _: self.export_results())
        self.root.bind("<Control-s>", lambda _: self.export_results())
        self.root.bind("<Control-d>", lambda _: self.toggle_theme())
        self.root.bind("<Control-k>", lambda _: self.clear_results())
        self.count_entry.bind("<Return>", lambda _: self.shuffle_manual())

    def _handle_ctrl_c(self, event) -> None:
        focused = self.root.focus_get()
        if focused in {self.add_text, self.search_entry}:
            return
        if focused == self.count_entry:
            if self.count_entry.selection_present():
                return
        self.copy_result_to_clipboard()

    def set_status(self, msg: str, *, error: bool = False) -> None:
        self.status_var.set(msg)
        if error:
            try:
                self.root.bell()
            except Exception:
                pass

    def update_metrics(self) -> None:
        self.total_count_var.set(str(len(self.usernames)))
        self.remaining_pool_var.set(str(len(self.elimination_pool)))
        if self.loaded_file:
            fname = os.path.basename(self.loaded_file)
            self.file_name_var.set(fname)
            self.file_badge_var.set(f"📄 {fname}  •  {len(self.usernames)} usernames")
        else:
            self.file_name_var.set("No file loaded")
            self.file_badge_var.set("No file loaded")

    def show_shortcuts_dialog(self) -> None:
        shortcuts = (
            "Keyboard Shortcuts:\n\n"
            "• Ctrl + O : Open Dataset File\n"
            "• Ctrl + R / F5 : Reload Current File\n"
            "• Ctrl + Enter / Enter : Shuffle Usernames\n"
            "• Ctrl + C : Copy Results to Clipboard\n"
            "• Ctrl + E / Ctrl + S : Export Results\n"
            "• Ctrl + D : Toggle Dark / Light Theme\n"
            "• Ctrl + K : Clear Results"
        )
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)

    def show_about_dialog(self) -> None:
        about = (
            f"{APP_NAME} {APP_VERSION}\n"
            f"{APP_SUBTITLE}\n\n"
            "Designed and built with modern Python & Tkinter.\n"
            "Features fast streaming Excel (.xlsx, .xlsm), CSV, and Text reading,\n"
            "custom delimiters, social prefixes, elimination draws, and dark mode.\n\n"
            "Author: MD Showrav Zaman\n"
            "GitHub: https://github.com/Smokianlord/Username-Shuffler"
        )
        messagebox.showinfo(f"About {APP_NAME}", about)

    def on_close(self) -> None:
        # Save geometry and settings
        self.config.set("window_geometry", self.root.geometry())
        self.config.set("delimiter", self.delimiter_var.get())
        self.config.set("prefix", self.prefix_var.get())
        self.config.set("suffix", self.suffix_var.get())
        self.config.set("numbered", self.numbered_var.get())
        self.config.set("case", self.case_var.get())
        self.config.set("auto_shuffle", self.auto_shuffle_var.get())
        self.root.destroy()


# ============================================================================
# High DPI Awareness & Process ID
# ============================================================================

def setup_windows_environment() -> None:
    if os.name != "nt":
        return
    # Set explicit AppUserModelID for taskbar grouping and icons
    try:
        import ctypes
        app_id = f"ShowravZaman.UsernameShuffler.{APP_VERSION}"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass

    # Enable High-DPI awareness on Windows 10/11
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    setup_windows_environment()
    target_arg = sys.argv[1] if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else None
    root = tk.Tk()
    UsernameShufflerApp(root, target_file=target_arg)
    root.mainloop()


if __name__ == "__main__":
    main()
