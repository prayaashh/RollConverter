#!/usr/bin/env python3
"""Roll Converter - turn lists of roll numbers into student names.

A single-file desktop app that runs on Windows and macOS using only the
Python standard library. The student list lives in roster.csv next to this
file and can be replaced from inside the app (CSV, TSV or XLSX).
"""

from __future__ import annotations

import csv
import io
import os
import re
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

APP_NAME = "Roll Converter"
ROSTER_FILE = "roster.csv"
BACKUP_FILE = "roster.backup.csv"
LEGACY_ROSTER = "rollnumberswithnames.csv"

# ---------------------------------------------------------------- palette ---
BG = "#0a0a0b"
PANEL = "#121214"
PANEL_HOVER = "#1b1b20"
BORDER = "#26262b"
BORDER_FOCUS = "#4b4b57"
TEXT = "#ececef"
MUTED = "#85858f"
ACCENT = "#8b7cff"
ACCENT_HOVER = "#a294ff"
ACCENT_TEXT = "#0a0a0b"
AMBER = "#f5cf5b"
RED = "#f2777a"
SELECTION = "#2c2c38"

PLACEHOLDER = "Paste roll numbers here - commas, spaces or new lines all work"


# ------------------------------------------------------------------ paths ---
def app_dir() -> Path:
    """Folder the app lives in (works for scripts and PyInstaller bundles)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def user_data_dir() -> Path:
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / "RollConverter"
    elif os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home()) / "RollConverter"
    else:
        base = Path.home() / ".local" / "share" / "RollConverter"
    base.mkdir(parents=True, exist_ok=True)
    return base


def choose_roster_dir() -> Path:
    """Prefer the app folder; fall back to a user data folder if read-only."""
    folder = app_dir()
    probe = folder / ".write-test"
    try:
        probe.write_text("", encoding="utf-8")
        probe.unlink()
        return folder
    except OSError:
        return user_data_dir()


ROSTER_DIR = choose_roster_dir()
ROSTER_PATH = ROSTER_DIR / ROSTER_FILE


# ----------------------------------------------------------- roster input ---
ROLL_RE = re.compile(r"^\d{2,}$")
CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
TOKEN_SPLIT_RE = re.compile(r"([,;\s]+)")
TRIM = ".,:;!?()[]{}\"'"
NAMESPACE = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RELATIONSHIP_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def tidy(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def looks_like_roll(value: str) -> bool:
    return bool(ROLL_RE.match(re.sub(r"\s+", "", value)))


def read_csv_rows(path: Path) -> list[list[str]]:
    raw = path.read_bytes()
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("The file is not valid CSV text.")

    sample = text[:4096]
    lines = [line for line in sample.splitlines() if line.strip()][:20]
    score = {delimiter: sum(1 for line in lines if delimiter in line) for delimiter in ("\t", ";", ",")}
    delimiter = ","
    best = -1
    for candidate in ("\t", ";", ","):
        if score[candidate] > best:
            delimiter, best = candidate, score[candidate]

    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def _column_index(reference: str) -> int:
    letters = CELL_REF_RE.match(reference).group(1).upper()
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _natural_key(name: str):
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", name)]


def _first_sheet_member(archive) -> str:
    import xml.etree.ElementTree as ET

    names = archive.namelist()
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

        targets = {}
        for rel in relationships.iter(f"{RELATIONSHIP_NS}Relationship"):
            target = (rel.get("Target") or "").lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            targets[rel.get("Id")] = target

        for sheet in workbook.iter(f"{NAMESPACE}sheet"):
            rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if rid in targets and targets[rid] in names:
                return targets[rid]
    except (KeyError, ET.ParseError):
        pass

    sheets = sorted(
        (n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")),
        key=_natural_key,
    )
    if not sheets:
        raise ValueError("This workbook has no worksheets.")
    return sheets[0]


def read_xlsx_rows(path: Path) -> list[list[str]]:
    import xml.etree.ElementTree as ET
    import zipfile

    with zipfile.ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.iter(f"{NAMESPACE}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{NAMESPACE}t")))
        sheet = ET.fromstring(archive.read(_first_sheet_member(archive)))

    rows: list[list[str]] = []
    for row in sheet.iter(f"{NAMESPACE}row"):
        cells: dict[int, str] = {}
        for cell in row.findall(f"{NAMESPACE}c"):
            reference = cell.get("r") or ""
            index = _column_index(reference) if CELL_REF_RE.match(reference) else len(cells)
            kind = cell.get("t")
            value_node = cell.find(f"{NAMESPACE}v")
            inline_node = cell.find(f"{NAMESPACE}is")

            if kind == "s":
                try:
                    value = shared[int(value_node.text)] if value_node is not None else ""
                except (TypeError, ValueError, IndexError):
                    value = ""
            elif kind == "inlineStr":
                value = "".join(n.text or "" for n in inline_node.iter(f"{NAMESPACE}t")) if inline_node is not None else ""
            elif kind == "b":
                value = "TRUE" if value_node is not None and value_node.text == "1" else "FALSE"
            else:
                value = (value_node.text or "") if value_node is not None else ""
                if value:
                    try:
                        number = float(value)
                        if number.is_integer():
                            value = str(int(number))
                    except ValueError:
                        pass

            cells[index] = tidy(value)

        if cells:
            rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
    return rows


def pick_columns(rows: list[list[str]]) -> tuple[int, int]:
    """Find the roll-number column and the name column."""
    width = max(len(row) for row in rows)
    if width < 2:
        raise ValueError("The file needs two columns: roll number, then name.")

    sample = rows[:250]
    roll_col, roll_score = 0, -1.0
    for column in range(width):
        values = [row[column] for row in sample if column < len(row) and row[column]]
        if not values:
            continue
        score = sum(looks_like_roll(v) for v in values) / len(values)
        if score > roll_score:
            roll_col, roll_score = column, score

    name_col, name_score = None, -1.0
    for column in range(width):
        if column == roll_col:
            continue
        values = [row[column] for row in sample if column < len(row) and row[column]]
        if not values:
            continue
        score = sum(any(ch.isalpha() for ch in v) for v in values) / len(values)
        if score > name_score:
            name_col, name_score = column, score

    if name_col is None:
        name_col = 1 if roll_col == 0 else 0
    return roll_col, name_col


def has_header(rows: list[list[str]], roll_col: int) -> bool:
    if not rows:
        return False
    first = rows[0][roll_col] if roll_col < len(rows[0]) else ""
    if not first:
        return True
    if looks_like_roll(first):
        return False
    rest = [row[roll_col] for row in rows[1:11] if roll_col < len(row) and row[roll_col]]
    if not rest:
        return False
    return sum(looks_like_roll(v) for v in rest) >= max(1, len(rest) // 2)


def load_pairs(path: Path) -> tuple[list[tuple[str, str]], int]:
    """Read a spreadsheet and return (roll, name) pairs plus a skipped count."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        rows = read_xlsx_rows(path)
    elif suffix == ".xls":
        raise ValueError("Old .xls files are not supported. Save it as .xlsx or .csv first.")
    else:
        rows = read_csv_rows(path)

    rows = [[tidy(cell) for cell in row] for row in rows]
    rows = [row for row in rows if any(row)]
    if not rows:
        return [], 0

    roll_col, name_col = pick_columns(rows)
    if has_header(rows, roll_col):
        rows = rows[1:]

    pairs: list[tuple[str, str]] = []
    skipped = 0
    for row in rows:
        roll = row[roll_col] if roll_col < len(row) else ""
        name = row[name_col] if name_col < len(row) else ""
        roll = re.sub(r"\s+", "", roll)
        if not roll or not name or roll == name:
            skipped += 1
            continue
        pairs.append((roll, name))
    return pairs, skipped


class Roster:
    """Roll number -> name lookup, kept in one dict to stay small."""

    def __init__(self, pairs: list[tuple[str, str]], skipped: int = 0):
        self.names: dict[str, str] = {}
        self.duplicates = 0
        self.skipped = skipped
        self.lengths: set[int] = set()

        for roll, name in pairs:
            if roll in self.names and self.names[roll] != name:
                self.duplicates += 1
            self.names[roll] = name
            if roll.isdigit():
                self.lengths.add(len(roll))
                canonical = str(int(roll))
                if canonical != roll:
                    self.names.setdefault(canonical, name)

        if not self.lengths:
            self.lengths = {4}

    def __len__(self) -> int:
        return len(self.names)

    def lookup(self, token: str) -> str | None:
        name = self.names.get(token)
        if name:
            return name
        if token.isdigit() and len(token) > 1:
            return self.names.get(str(int(token)))
        return None


def classify(token: str, roster: Roster) -> tuple[str, str]:
    """Sort one token into a name, an unmatched roll number, or plain text."""
    cleaned = token.strip(TRIM)
    if cleaned:
        name = roster.lookup(cleaned)
        if name:
            return "name", name
        if cleaned.isdigit() and len(cleaned) in roster.lengths:
            return "roll", cleaned
    return "text", token


def build_view(text: str, roster: Roster, auto_format: bool):
    """Turn the input into (segments, found, unmatched).

    Segments are (text, tag) pairs so the UI can colour unmatched rolls.
    Anything that is not a roll number is passed through as text. With
    auto_format on, roll numbers that sit next to each other on the same
    line are comma separated; text and line breaks are left untouched.
    """
    parts = TOKEN_SPLIT_RE.split(text)
    items: list[tuple[str, str, str]] = []

    for index in range(0, len(parts), 2):
        token = parts[index]
        separator = parts[index + 1] if index + 1 < len(parts) else ""
        if not token:
            if separator:
                items.append(("text", "", separator))
            continue
        kind, value = classify(token, roster)
        items.append((kind, value, separator))

    found = sum(1 for kind, _value, _sep in items if kind == "name")
    unmatched = sum(1 for kind, _value, _sep in items if kind == "roll")

    segments: list[tuple[str, str | None]] = []
    for position, (kind, value, separator) in enumerate(items):
        if position == len(items) - 1:
            separator = ""
        elif auto_format and kind in ("name", "roll"):
            following = items[position + 1][0]
            if following in ("name", "roll") and "\n" not in separator:
                separator = ", "
        segments.append((value, "unknown" if kind == "roll" else None))
        if separator:
            segments.append((separator, None))
    return segments, found, unmatched


def save_roster(pairs: list[tuple[str, str]], source: Path | None = None) -> None:
    import shutil

    if ROSTER_PATH.exists() and source is not None and source.resolve() != ROSTER_PATH.resolve():
        shutil.copy2(ROSTER_PATH, ROSTER_DIR / BACKUP_FILE)
    temporary = ROSTER_PATH.with_name(ROSTER_PATH.name + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerows(pairs)
    os.replace(temporary, ROSTER_PATH)


# ------------------------------------------------------------------- theme ---
def set_windows_dpi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def set_windows_dark_titlebar(root: tk.Misc) -> None:
    """Ask DWM to paint the title bar dark (Windows 10/11)."""
    if os.name != "nt":
        return
    try:
        import ctypes

        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        enabled = ctypes.c_int(1)
        for attribute in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (new, old)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(enabled), ctypes.sizeof(enabled)
            )
    except Exception:
        pass


def build_icon(root: tk.Misc) -> tk.PhotoImage | None:
    """Draw a tiny window icon in code so the app ships without image files."""
    size = 32
    bars = ((9, 16), (14, 11), (19, 13))
    rows = []
    for y in range(size):
        row = []
        for x in range(size):
            inside = _inside_rounded(x, y, 2, 7, size)
            color = "#151518" if inside else BG
            if inside:
                for top, width in bars:
                    if top <= y < top + 3 and 8 <= x < 8 + width:
                        color = ACCENT
                        break
            row.append(color)
        rows.append("{" + " ".join(row) + "}")
    try:
        image = tk.PhotoImage(master=root, width=size, height=size)
        image.put(" ".join(rows))
        return image
    except tk.TclError:
        return None


def _inside_rounded(x: int, y: int, inset: int, radius: int, size: int) -> bool:
    low, high = inset, size - inset - 1
    if not (low <= x <= high and low <= y <= high):
        return False
    if low + radius <= x <= high - radius or low + radius <= y <= high - radius:
        return True
    cx = low + radius if x < low + radius else high - radius
    cy = low + radius if y < low + radius else high - radius
    return (x - cx) ** 2 + (y - cy) ** 2 <= radius * radius


class FlatButton(tk.Label):
    """A label styled as a button so the dark look survives macOS' native theme."""

    VARIANTS = {
        "primary": (ACCENT, ACCENT_HOVER, ACCENT_TEXT, ACCENT),
        "secondary": (PANEL, PANEL_HOVER, TEXT, BORDER),
        "ghost": (BG, PANEL, MUTED, BG),
    }

    def __init__(self, master, text, command, *, variant="secondary", **kwargs):
        self.base, self.hover, self.foreground, self.outline = self.VARIANTS[variant]
        self.command = command
        self.enabled = True
        kwargs.setdefault("padx", 14)
        kwargs.setdefault("pady", 7)
        super().__init__(
            master,
            text=text,
            bg=self.base,
            fg=self.foreground,
            cursor="hand2",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.outline,
            highlightcolor=self.outline,
            **kwargs,
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _on_enter(self, _event=None):
        if self.enabled:
            self.configure(bg=self.hover)

    def _on_leave(self, _event=None):
        if self.enabled:
            self.configure(bg=self.base)

    def _on_press(self, _event=None):
        if self.enabled:
            self.configure(bg=self.hover)

    def _on_release(self, event):
        if not self.enabled:
            return
        self.configure(bg=self.hover)
        if 0 <= event.x < self.winfo_width() and 0 <= event.y < self.winfo_height():
            self.command()

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        if enabled:
            self.configure(bg=self.base, fg=self.foreground, cursor="hand2")
        else:
            self.configure(bg=self.base, fg=BORDER_FOCUS, cursor="arrow")


class Toggle(tk.Canvas):
    """Small iOS-style switch drawn on a canvas, so it looks the same everywhere."""

    def __init__(self, master, value: bool = True, command=None, width: int = 40, height: int = 22):
        super().__init__(
            master,
            width=width,
            height=height,
            bg=BG,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.value = value
        self.command = command
        self._draw()
        self.bind("<Button-1>", self.flip)

    def _draw(self) -> None:
        self.delete("all")
        width = int(self.cget("width"))
        height = int(self.cget("height"))
        radius = height // 2
        track = ACCENT if self.value else BORDER
        self.create_oval(0, 0, height, height, fill=track, outline="")
        self.create_oval(width - height, 0, width, height, fill=track, outline="")
        self.create_rectangle(radius, 0, width - radius, height, fill=track, outline="")
        knob = radius - 4
        centre = width - radius if self.value else radius
        self.create_oval(
            centre - knob,
            radius - knob,
            centre + knob,
            radius + knob,
            fill=ACCENT_TEXT if self.value else MUTED,
            outline="",
        )

    def flip(self, _event=None) -> None:
        self.value = not self.value
        self._draw()
        if self.command is not None:
            self.command(self.value)


# --------------------------------------------------------------------- app ---
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.roster = Roster([], 0)
        self.roster_path = ROSTER_PATH
        self.roster_mtime: datetime | None = None
        self.roster_missing = True
        self._after_id: str | None = None
        self._status_after_id: str | None = None
        self._placeholder_active = True
        self._copied_text = ""
        self._result_count = 0
        self.auto_format = True
        self.modifier = "Cmd" if sys.platform == "darwin" else "Ctrl"

        self._build_window()
        self._build_fonts()
        self._build_layout()
        self._bind_keys()
        self._load_initial_roster()

    # -- window ------------------------------------------------------------
    def _build_window(self) -> None:
        self.root.title(APP_NAME)
        self.root.configure(bg=BG)
        width, height = 780, 640
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 3)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(660, 560)
        self.icon = build_icon(self.root)
        if self.icon is not None:
            try:
                self.root.iconphoto(True, self.icon)
            except tk.TclError:
                pass
        if os.name == "nt":
            # Must run after the window exists, so it costs no startup time.
            self.root.after_idle(lambda: set_windows_dark_titlebar(self.root))

    def _build_fonts(self) -> None:
        available = set(tkfont.families(self.root))

        def pick(candidates: list[str], fallback: str) -> str:
            for family in candidates:
                if family in available:
                    return family
            return fallback

        ui = pick(
            ["Segoe UI Variable Text", "Segoe UI", "SF Pro Text", "Helvetica Neue", "Inter", "Arial"],
            "TkDefaultFont",
        )
        self.f_title = (ui, 21, "bold")
        self.f_section = (ui, 8, "bold")
        self.f_body = (ui, 12)
        self.f_small = (ui, 9)
        self.f_chip = (ui, 9)
        self.f_button = (ui, 9, "bold")

    # -- layout ------------------------------------------------------------
    def _build_layout(self) -> None:
        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=26, pady=20)

        # header ----------------------------------------------------------
        header = tk.Frame(container, bg=BG)
        header.pack(fill="x")

        titles = tk.Frame(header, bg=BG)
        titles.pack(side="left", anchor="w")
        tk.Label(titles, text="Roll Converter", bg=BG, fg=TEXT, font=self.f_title).pack(anchor="w")

        actions = tk.Frame(header, bg=BG)
        actions.pack(side="right", anchor="e")
        self.chip = tk.Label(
            actions,
            text="no list",
            bg=PANEL,
            fg=MUTED,
            font=self.f_chip,
            padx=12,
            pady=7,
        )
        self.chip.pack(side="left", padx=(0, 8))
        FlatButton(actions, "Change roster", self.change_roster, variant="primary", font=self.f_button).pack(
            side="left"
        )

        # footer (packed before the expanding output so it is never squeezed out)
        footer = tk.Frame(container, bg=BG)
        footer.pack(side="bottom", fill="x", pady=(14, 0))
        self.status = tk.Label(footer, text="", bg=BG, fg=MUTED, font=self.f_small, anchor="w")
        self.status.pack(side="left")
        tk.Label(
            footer,
            text=f"Esc clears  ·  {self.modifier}+O changes roster",
            bg=BG,
            fg=BORDER_FOCUS,
            font=self.f_small,
        ).pack(side="right")

        # spacer ---------------------------------------------------------
        tk.Frame(container, bg=BG, height=18).pack(fill="x")

        # input header ---------------------------------------------------
        input_header = tk.Frame(container, bg=BG)
        input_header.pack(fill="x", pady=(0, 6))
        self.input_label = tk.Label(
            input_header,
            text="INPUT",
            bg=BG,
            fg=MUTED,
            font=self.f_section,
        )
        self.input_label.pack(side="left")
        self.clear_button = FlatButton(input_header, "Clear", self.clear_input, variant="ghost", font=self.f_small)
        self.clear_button.pack(side="right", padx=(8, 0))

        # input ----------------------------------------------------------
        self.input = tk.Text(
            container,
            height=5,
            bg=PANEL,
            fg=TEXT,
            font=self.f_body,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER_FOCUS,
            insertbackground=ACCENT,
            insertwidth=2,
            selectbackground=SELECTION,
            selectforeground=TEXT,
            padx=14,
            pady=12,
            spacing1=3,
            spacing3=3,
            wrap="word",
        )
        self.input.pack(fill="x")
        self.input.tag_configure("placeholder", foreground=MUTED)
        self.input.bind("<FocusIn>", self._on_input_focus_in)
        self.input.bind("<FocusOut>", self._on_input_focus_out)
        self.input.bind("<<Modified>>", self._on_input_modified)
        self.input.edit_modified(False)
        self._show_placeholder()

        # output header --------------------------------------------------
        output_header = tk.Frame(container, bg=BG)
        output_header.pack(fill="x", pady=(18, 6))
        self.output_label = tk.Label(output_header, text="OUTPUT", bg=BG, fg=MUTED, font=self.f_section)
        self.output_label.pack(side="left")
        self.count_label = tk.Label(output_header, text="", bg=BG, fg=MUTED, font=self.f_small)
        self.count_label.pack(side="left", padx=(10, 0))
        self.copy_button = FlatButton(
            output_header, "Copy names", self.copy_output, variant="secondary", font=self.f_button
        )
        self.copy_button.pack(side="right")
        self.copy_button.set_enabled(False)

        self.format_toggle = Toggle(output_header, value=self.auto_format, command=self.set_auto_format)
        self.format_toggle.pack(side="right", padx=(0, 16))
        self.format_label = tk.Label(
            output_header,
            text="Auto format",
            bg=BG,
            fg=MUTED,
            font=self.f_small,
            cursor="hand2",
        )
        self.format_label.pack(side="right", padx=(0, 8))
        self.format_label.bind("<Button-1>", lambda _e: self.format_toggle.flip())

        # output ---------------------------------------------------------
        output_wrap = tk.Frame(container, bg=BG)
        output_wrap.pack(fill="both", expand=True)

        self.scrollbar = tk.Scrollbar(
            output_wrap,
            orient="vertical",
            width=10,
            bd=0,
            relief="flat",
            troughcolor=BG,
            bg=BORDER,
            activebackground=BORDER_FOCUS,
            highlightthickness=0,
            elementborderwidth=0,
        )
        self.scrollbar.pack(side="right", fill="y", padx=(6, 0))

        self.output = tk.Text(
            output_wrap,
            height=8,
            bg=PANEL,
            fg=TEXT,
            font=self.f_body,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
            selectbackground=SELECTION,
            selectforeground=TEXT,
            padx=14,
            pady=12,
            spacing1=4,
            spacing3=4,
            wrap="word",
            cursor="arrow",
            yscrollcommand=self.scrollbar.set,
        )
        self.output.pack(side="left", fill="both", expand=True)
        self.scrollbar.configure(command=self.output.yview)
        self.output.tag_configure("empty", foreground=MUTED)
        self.output.tag_configure("unknown", foreground=AMBER)
        self.output.configure(state="disabled")

        self._render()

    def _bind_keys(self) -> None:
        self.root.bind("<Escape>", lambda _e: self.clear_input())
        self.root.bind("<Control-o>", lambda _e: self.change_roster())
        self.root.bind("<Command-o>", lambda _e: self.change_roster())
        self.root.bind("<Control-Shift-C>", lambda _e: self.copy_output())
        self.root.bind("<Command-Shift-C>", lambda _e: self.copy_output())

    # -- input placeholder -------------------------------------------------
    def _show_placeholder(self) -> None:
        if self.input.get("1.0", "end-1c"):
            return
        self._placeholder_active = True
        self.input.insert("1.0", PLACEHOLDER, "placeholder")
        self.input.edit_modified(False)

    def _hide_placeholder(self) -> None:
        if not self._placeholder_active:
            return
        self._placeholder_active = False
        self.input.delete("1.0", "end")
        self.input.edit_modified(False)

    def _on_input_focus_in(self, _event=None):
        if self._placeholder_active:
            self._hide_placeholder()

    def _on_input_focus_out(self, _event=None):
        if not self.input.get("1.0", "end-1c"):
            self._show_placeholder()

    def _input_text(self) -> str:
        if self._placeholder_active:
            return ""
        return self.input.get("1.0", "end-1c")

    # -- conversion --------------------------------------------------------
    def _on_input_modified(self, _event=None):
        if not self.input.edit_modified():
            return
        self.input.edit_modified(False)
        if self._placeholder_active:
            return
        if self._after_id is not None:
            self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(40, self._render)

    def _render(self) -> None:
        self._after_id = None
        text = self._input_text()
        segments, found, unmatched = build_view(text, self.roster, self.auto_format)

        self.output.configure(state="normal")
        self.output.delete("1.0", "end")

        if not segments:
            message = "Add a student list to get started." if not len(self.roster) else "Type or paste above."
            self.output.insert("1.0", message, "empty")
            self.count_label.configure(text="")
            self.copy_button.set_enabled(False)
            self._copied_text = ""
            self._result_count = 0
        else:
            for value, tag in segments:
                if tag:
                    self.output.insert("end", value, tag)
                else:
                    self.output.insert("end", value)

            parts = []
            if found:
                parts.append(f"{found} name" + ("" if found == 1 else "s"))
            if unmatched:
                parts.append(f"{unmatched} not found")
            self.count_label.configure(text="  ·  ".join(parts))

            self._copied_text = "".join(value for value, _tag in segments)
            self._result_count = found + unmatched
            self.copy_button.set_enabled(True)

        self.output.configure(state="disabled")

    def set_auto_format(self, value: bool) -> None:
        self.auto_format = value
        self.format_label.configure(fg=TEXT if value else MUTED)
        self._render()

    def clear_input(self) -> None:
        self.input.delete("1.0", "end")
        self._placeholder_active = False
        self.input.edit_modified(False)
        if self.root.focus_get() is not self.input:
            self._show_placeholder()
        self._render()

    def copy_output(self) -> None:
        if not self._copied_text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._copied_text)
        self.root.update_idletasks()
        count = self._result_count
        if count:
            self.flash(f"Copied {count} name" + ("" if count == 1 else "s"))
        else:
            self.flash("Copied")

    # -- roster ------------------------------------------------------------
    def _load_initial_roster(self) -> None:
        source: Path | None = None
        if ROSTER_PATH.exists():
            source = ROSTER_PATH
        else:
            legacy = app_dir() / LEGACY_ROSTER
            if legacy.exists():
                source = legacy

        if source is None:
            self.set_roster(Roster([], 0), ROSTER_PATH, missing=True)
            return

        try:
            pairs, skipped = load_pairs(source)
        except Exception as error:  # noqa: BLE001 - surfaced to the user
            self.set_roster(Roster([], 0), source, missing=True)
            self.flash(f"Could not read {source.name}: {error}", error=True)
            return

        if source != ROSTER_PATH:
            try:
                save_roster(pairs, source)
            except OSError:
                pass
        self.set_roster(Roster(pairs, skipped), ROSTER_PATH, missing=not pairs)

    def set_roster(self, roster: Roster, path: Path, *, missing: bool = False) -> None:
        self.roster = roster
        self.roster_path = path
        self.roster_missing = missing
        try:
            self.roster_mtime = datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            self.roster_mtime = None

        if missing or not len(roster):
            self.chip.configure(text="no list", fg=MUTED)
            self._set_status("No student list yet — click Change roster to load one.", error=True)
        else:
            self.chip.configure(text=f"{len(roster):,} students · {path.name}", fg=TEXT)
            stamp = f" · updated {self.roster_mtime:%d %b %Y}" if self.roster_mtime else ""
            self._set_status(f"{path.name} · {len(roster):,} students{stamp}")
        self._render()

    def _set_status(self, message: str, *, error: bool = False) -> None:
        if self._status_after_id is not None:
            self.root.after_cancel(self._status_after_id)
            self._status_after_id = None
        self.status.configure(text=message, fg=RED if error else MUTED)
        self._default_status = message

    def flash(self, message: str, *, error: bool = False) -> None:
        if self._status_after_id is not None:
            self.root.after_cancel(self._status_after_id)
        self.status.configure(text=message, fg=RED if error else ACCENT)
        self._status_after_id = self.root.after(3200, self._restore_status)

    def _restore_status(self) -> None:
        self._status_after_id = None
        self.status.configure(text=getattr(self, "_default_status", ""), fg=MUTED)

    def change_roster(self) -> None:
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title="Choose the student spreadsheet",
            initialdir=str(self.roster_path.parent),
            filetypes=[
                ("Spreadsheets", "*.csv *.tsv *.txt *.xlsx *.xlsm"),
                ("CSV files", "*.csv"),
                ("Excel workbooks", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not chosen:
            return

        source = Path(chosen)
        try:
            pairs, skipped = load_pairs(source)
        except Exception as error:  # noqa: BLE001 - surfaced to the user
            messagebox.showerror(
                APP_NAME,
                f"Could not read that file.\n\n{error}",
                parent=self.root,
            )
            return

        if not pairs:
            messagebox.showerror(
                APP_NAME,
                "No roll number and name pairs were found in that file.\n\n"
                "Expected two columns: the roll number, then the student's name.",
                parent=self.root,
            )
            return

        roster = Roster(pairs, skipped)
        summary = f"Found {len(roster):,} students in {source.name}."
        if skipped:
            summary += f"\n{skipped} row(s) had no name or no roll number and were skipped."
        if roster.duplicates:
            summary += f"\n{roster.duplicates} duplicate roll number(s) were overwritten."
        summary += "\n\nUse this list from now on?"

        if not messagebox.askokcancel(APP_NAME, summary, parent=self.root):
            return

        try:
            save_roster(pairs, source)
        except OSError as error:
            messagebox.showerror(
                APP_NAME,
                f"Could not save the new list.\n\n{error}",
                parent=self.root,
            )
            return

        self.set_roster(roster, ROSTER_PATH)
        self.flash(f"Roster updated · {len(roster):,} students")


# ---------------------------------------------------------------- selftest ---
def selftest() -> int:
    print(f"roster file : {ROSTER_PATH}")
    if not ROSTER_PATH.exists():
        print("FAIL: roster file not found")
        return 1

    pairs, skipped = load_pairs(ROSTER_PATH)
    roster = Roster(pairs, skipped)
    print(f"students    : {len(roster)}")
    print(f"skipped     : {skipped}")
    print(f"duplicates  : {roster.duplicates}")
    if len(roster) == 0:
        print("FAIL: no students loaded")
        return 1

    first_roll, first_name = pairs[0]
    checks = [
        (first_roll, first_name),
        ("0" + first_roll if first_roll.isdigit() else first_roll, first_name),
        ("999999", None),
    ]
    for token, expected in checks:
        actual = roster.lookup(token)
        status = "ok" if actual == expected else "FAIL"
        print(f"  {status:4} {token!r:<12} -> {actual!r}")
        if actual != expected:
            return 1

    auto_segments, found, unmatched = build_view(f"{first_roll} 0000\n{first_roll} hello", roster, True)
    auto_text = "".join(value for value, _tag in auto_segments)
    print(f"auto view   : {auto_text!r}")
    if (found, unmatched) != (2, 1):
        print("FAIL: wrong name / unmatched counts")
        return 1
    if auto_text != f"{first_name}, 0000\n{first_name} hello":
        print("FAIL: auto format should add commas within a line only")
        return 1

    same_line, _f, _u = build_view(f"{first_roll} {first_roll}\n{first_roll}", roster, True)
    same_text = "".join(value for value, _tag in same_line)
    print(f"auto newline: {same_text!r}")
    if same_text != f"{first_name}, {first_name}\n{first_name}":
        print("FAIL: auto format should keep the line breaks")
        return 1

    exact_segments, _f, _u = build_view(f"{first_roll} {first_roll}\n{first_roll}", roster, False)
    exact_text = "".join(value for value, _tag in exact_segments)
    print(f"exact view  : {exact_text!r}")
    if exact_text != f"{first_name} {first_name}\n{first_name}":
        print("FAIL: exact format did not keep the input layout")
        return 1

    print("selftest ok")
    return 0


def main(argv: list[str]) -> int:
    if "--selftest" in argv:
        return selftest()

    set_windows_dpi()
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
