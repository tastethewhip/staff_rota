#!/usr/bin/env python3
"""Staff Rota - a tkinter spreadsheet for pharmacy staff and weekly rota generation.

Tab 1 "Staff":       import / edit / save staff (CSV or Excel .xlsx).
Tab 2 "Pharmacies":  each pharmacy has a number and its own settings
                     (staff per day, shift hours, minimum narcotic-aware / PC-competent
                     staff per shift, and opening days).
Tab 3 "Willingness": a 0-5 score per person per pharmacy (0 = never roster there),
                     with an adjustable importance weighting.
Tab 4 "Rota":        generate a weekly rota for a chosen week, respecting each person's
                     weekly hours, days available, holidays and preferred town.

Double-click any cell to edit it.  Run with:  python3 staff_rota.py
"""
import csv
import re
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import ttk, filedialog, messagebox

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ALL_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# ---- Staff columns --------------------------------------------------------- #
COLUMNS = ["Name", "Minutes from pharmacy", "Hours per week", "Preferred town",
           "Narcotic awareness", "PC competence", "Days available", "Holidays"]
REQUIRED_COLS = COLUMNS[:6]          # the ones expected in an uploaded file
YESNO_COLS = {"Narcotic awareness", "PC competence"}
NUMERIC_COLS = {"Minutes from pharmacy", "Hours per week"}

# Accept the headings however they're spelt in the uploaded file.
ALIASES = {
    "name": "Name", "staffname": "Name", "fullname": "Name", "employee": "Name",
    "minutesfrompharmacy": "Minutes from pharmacy", "minutes": "Minutes from pharmacy",
    "mins": "Minutes from pharmacy", "minsfrompharmacy": "Minutes from pharmacy",
    "traveltime": "Minutes from pharmacy", "minutestopharmacy": "Minutes from pharmacy",
    "traveltimemins": "Minutes from pharmacy", "traveltimeminutes": "Minutes from pharmacy",
    "hoursperweek": "Hours per week", "hours": "Hours per week",
    "weeklyhours": "Hours per week", "contractedhours": "Hours per week",
    "preferredtown": "Preferred town", "prefferredtown": "Preferred town",
    "preferedtown": "Preferred town", "prefferedtown": "Preferred town",
    "town": "Preferred town", "branch": "Preferred town",
    "narcoticawareness": "Narcotic awareness", "narcotics": "Narcotic awareness",
    "narcoticsawareness": "Narcotic awareness", "cdawareness": "Narcotic awareness",
    "pccompetence": "PC competence", "pc": "PC competence",
    "computercompetence": "PC competence", "pccompetent": "PC competence",
    "daysavailable": "Days available", "availability": "Days available",
    "workingdays": "Days available", "days": "Days available",
    "holidays": "Holidays", "holiday": "Holidays", "leave": "Holidays",
    "annualleave": "Holidays",
}

# ---- Pharmacy columns ------------------------------------------------------ #
PH_COLS = ["No.", "Pharmacy", "Town", "Staff per day", "Shift hours",
           "Min narcotic aware", "Min PC competent"] + ALL_DAYS
PH_NUMERIC = {"No.", "Staff per day", "Shift hours", "Min narcotic aware", "Min PC competent"}
PH_DEFAULTS = {"Staff per day": "2", "Shift hours": "8",
               "Min narcotic aware": "1", "Min PC competent": "1",
               **{d: ("No" if d == "Sun" else "Yes") for d in ALL_DAYS}}


WILL_CHOICES = ["auto", "0", "1", "2", "3", "4", "5"]
WILL_HOME, WILL_OTHER = 5, 3          # what "auto" means: preferred town / anywhere else


def norm(header):
    return "".join(ch for ch in str(header).lower() if ch.isalnum())


def is_yes(value):
    return str(value).strip().lower() in {"yes", "y", "true", "1", "x", "✓", "✔"}


def yes_no(value):
    return "Yes" if is_yes(value) else "No"


def to_num(value):
    try:
        return float(str(value).strip() or 0)
    except ValueError:
        return 0.0


def fmt_num(h):
    return str(int(h)) if float(h).is_integer() else f"{h:g}"


# --------------------------------------------------------------------------- #
# Parsing days available and holidays
# --------------------------------------------------------------------------- #
def parse_days(text):
    """'Mon-Fri', 'Mon, Wed, Fri', 'Tuesday Thursday' -> set of day names. Blank = all."""
    text = str(text).strip()
    if not text or text.lower() in {"any", "all"}:
        return set(ALL_DAYS)
    days = set()
    for part in re.split(r"[,;/ ]+(?![-–])(?<![-–] )", text):
        part = part.strip()
        if not part:
            continue
        ends = [p.strip()[:3].title() for p in re.split(r"\s*[-–]\s*|\s+to\s+", part)]
        ends = [e for e in ends if e in ALL_DAYS]
        if len(ends) == 2:
            i, j = ALL_DAYS.index(ends[0]), ALL_DAYS.index(ends[1])
            rng = range(i, j + 1) if i <= j else list(range(i, 7)) + list(range(0, j + 1))
            days.update(ALL_DAYS[k] for k in rng)
        elif ends:
            days.add(ends[0])
    return days


DATE_RE = re.compile(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{1,2}/\d{2,4}")


def parse_date(s):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Bad date '{s}' (use dd/mm/yyyy)")


def parse_holidays(text):
    """'24/12/2026 - 02/01/2027; 15/11/2026' -> set of dates."""
    out = set()
    for item in re.split(r"[;,\n]", str(text)):
        found = DATE_RE.findall(item)
        if not found:
            if item.strip():
                raise ValueError(f"Can't read holiday '{item.strip()}' (use dd/mm/yyyy)")
            continue
        start = parse_date(found[0])
        end = parse_date(found[1]) if len(found) > 1 else start
        if end < start:
            start, end = end, start
        while start <= end:
            out.add(start)
            start += timedelta(days=1)
    return out


# --------------------------------------------------------------------------- #
# File reading
# --------------------------------------------------------------------------- #
def read_table(path):
    """Return (headers, rows) from a CSV or XLSX file."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = []
        for r in wb.active.iter_rows(values_only=True):
            row = []
            for c in r:
                if c is None:
                    c = ""
                elif isinstance(c, datetime):
                    c = c.strftime("%d/%m/%Y")
                row.append(c)
            rows.append(row)
        wb.close()
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        return [], []
    return [str(h) for h in rows[0]], rows[1:]


def map_rows(headers, rows, columns, aliases):
    """Map uploaded columns onto `columns`. Returns (records, missing_columns)."""
    index = {}
    for i, h in enumerate(headers):
        target = aliases.get(norm(h))
        if target and target not in index:
            index[target] = i
    missing = [c for c in columns if c not in index]
    records = []
    for r in rows:
        rec = {}
        for col in columns:
            i = index.get(col)
            val = r[i] if i is not None and i < len(r) else ""
            rec[col] = "" if val is None else str(val).strip()
        records.append(rec)
    return records, missing


def clean_staff(rec):
    for col in YESNO_COLS:
        rec[col] = yes_no(rec[col])
    for col in NUMERIC_COLS:
        if rec[col]:
            rec[col] = fmt_num(to_num(rec[col]))
    return rec


# --------------------------------------------------------------------------- #
# Excel (.xlsx) export
# --------------------------------------------------------------------------- #
NUM_RE = re.compile(r"^-?\d+(\.\d+)?$")
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
WARN_FILL = PatternFill("solid", fgColor="FFD7D7")


def write_xlsx(path, sheets):
    """Write a styled workbook.

    sheets: list of dicts with keys
      name   - sheet name
      rows   - list of rows; the first row is the bold header
      widths - optional list of column widths (characters)
      warn   - optional set of row indexes to shade red
      wrap   - optional bool: wrap text (multi-line cells)
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sh in sheets:
        ws = wb.create_sheet(re.sub(r"[\[\]:*?/\\]", "-", sh["name"])[:31])
        warn, wrap = sh.get("warn", set()), sh.get("wrap", False)
        for r, row in enumerate(sh["rows"]):
            values = []
            for v in row:
                v = "" if v is None else str(v)
                values.append(float(v) if r and NUM_RE.match(v) else v)
            ws.append(values)
            lines = max([str(v).count("\n") + 1 for v in row] or [1])
            if lines > 1:
                ws.row_dimensions[r + 1].height = 15 * lines
            for cell in ws[r + 1]:
                if r == 0:
                    cell.font = Font(bold=True)
                    cell.fill = HEADER_FILL
                    continue
                if r in warn:
                    cell.fill = WARN_FILL
                if wrap:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        for i, w in enumerate(sh.get("widths") or []):
            ws.column_dimensions[get_column_letter(i + 1)].width = w
        ws.freeze_panes = "A2"
    wb.save(path)


# --------------------------------------------------------------------------- #
# Spreadsheet-like Treeview with inline editing
# --------------------------------------------------------------------------- #
class EditableTree(ttk.Treeview):
    def __init__(self, master, columns, choices=None, readonly=(), numeric=(),
                 on_change=None, **kw):
        super().__init__(master, columns=columns, show="headings", **kw)
        self.choices = choices or {}
        self.readonly = set(readonly)
        self.numeric = set(numeric)
        self.on_change = on_change
        self._editor = None
        for c in columns:
            self.heading(c, text=c, command=lambda c=c: self.sort_by(c, False))
            self.column(c, width=120, anchor="w")
        self.bind("<Double-1>", self._begin_edit)
        self.tag_configure("odd", background="#f3f6fa")
        self.tag_configure("warn", background="#ffe1e1")

    def records(self):
        cols = self["columns"]
        return [dict(zip(cols, self.item(i, "values"))) for i in self.get_children()]

    def restripe(self):
        for i, iid in enumerate(self.get_children()):
            tags = [t for t in self.item(iid, "tags") if t != "odd"]
            if i % 2 and "warn" not in tags:
                tags.append("odd")
            self.item(iid, tags=tags)

    def sort_by(self, col, reverse):
        def key(iid):
            v = self.set(iid, col)
            try:
                return (0, float(v), "")
            except ValueError:
                return (1, 0, v.lower())
        for i, iid in enumerate(sorted(self.get_children(), key=key, reverse=reverse)):
            self.move(iid, "", i)
        self.heading(col, command=lambda: self.sort_by(col, not reverse))
        self.restripe()

    def _begin_edit(self, event):
        if self.identify_region(event.x, event.y) != "cell":
            return
        iid = self.identify_row(event.y)
        col = self.column(self.identify_column(event.x), "id")
        if not iid or col in self.readonly:
            return
        self.see(iid)
        x, y, w, h = self.bbox(iid, col)
        value = self.set(iid, col)
        if col in self.choices:
            ed = ttk.Combobox(self, values=self.choices[col], state="readonly")
            ed.set(value)
            ed.bind("<<ComboboxSelected>>", lambda e: self._commit())
        else:
            ed = ttk.Entry(self)
            ed.insert(0, value)
            ed.select_range(0, "end")
        ed.place(x=x, y=y, width=max(w, 60), height=h)
        ed.focus_set()
        ed.bind("<Return>", lambda e: self._commit())
        ed.bind("<KP_Enter>", lambda e: self._commit())
        ed.bind("<Escape>", lambda e: self._cancel())
        # Opening a dropdown list moves focus into the list, so don't treat
        # that as "finished editing" - wait, then check where focus went.
        ed.bind("<FocusOut>", lambda e: self.after(150, self._focus_check, ed))
        self._editor = (ed, iid, col)
        if isinstance(ed, ttk.Combobox):
            try:
                self.tk.call("ttk::combobox::Post", ed)   # open the list straight away
            except tk.TclError:
                pass

    def _focus_check(self, ed):
        if not self._editor or self._editor[0] is not ed:
            return                          # already saved or cancelled
        focus = self.tk.call("focus")
        if str(focus).startswith(str(ed)):
            return                          # focus is in this dropdown's own list
        self._commit()

    def _commit(self):
        if not self._editor:
            return
        ed, iid, col = self._editor
        self._editor = None
        new = ed.get().strip()
        ed.destroy()
        if col in self.numeric and new:
            try:
                new = fmt_num(float(new))
            except ValueError:
                messagebox.showwarning("Invalid number", f"'{new}' is not a number.")
                return
        if col == "Holidays" and new:
            try:
                parse_holidays(new)
            except ValueError as e:
                messagebox.showwarning("Invalid holiday", str(e))
                return
        if self.exists(iid):
            self.set(iid, col, new)
            if self.on_change:
                self.on_change()

    def _cancel(self):
        if self._editor:
            self._editor[0].destroy()
            self._editor = None


def scrolled(frame, tree):
    vs = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hs = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vs.grid(row=0, column=1, sticky="ns")
    hs.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)


# --------------------------------------------------------------------------- #
# Rota generation
# --------------------------------------------------------------------------- #
def ph_key(p):
    """Stable key for a pharmacy: its number, else its name."""
    return p.get("No.", "").strip() or p.get("Pharmacy", "").strip()


def pharmacy_label(p):
    no = p.get("No.", "").strip()
    name = p.get("Pharmacy", "").strip() or p.get("Town", "").strip()
    return f"{no}. {name}" if no else name


def generate_rota(staff, pharmacies, week_start, willingness=None, weight=20):
    """Greedy weekly rota.

    staff:      list of staff dicts (COLUMNS)
    pharmacies: list of pharmacy dicts (PH_COLS)
    week_start: date of the Monday (or first day) of the week
    willingness: {(name, ph_key): 0-5}; missing = auto (5 preferred town, 3 elsewhere)
    weight:     score points per willingness point (how much willingness matters)

    Returns (rota, info, warnings, bad) where
      rota[label][day] -> list of names ("Closed" days are absent),
      info[name] -> dict(assigned, shifts, holiday_days),
      warnings -> list of strings, bad -> set of pharmacy labels with problems.
    """
    dates = {d: week_start + timedelta(days=i) for i, d in enumerate(ALL_DAYS)}
    by_name = {s["Name"]: s for s in staff}
    contracted = {n: to_num(s["Hours per week"]) for n, s in by_name.items()}
    remaining = dict(contracted)
    avail = {n: parse_days(s.get("Days available", "")) for n, s in by_name.items()}
    hols = {n: parse_holidays(s.get("Holidays", "")) for n, s in by_name.items()}
    worked = {n: set() for n in by_name}
    shifts = {n: 0 for n in by_name}

    phs = []
    for p in pharmacies:
        phs.append({
            "label": pharmacy_label(p),
            "key": ph_key(p),
            "match": {p["Town"].strip().lower(), p["Pharmacy"].strip().lower()} - {""},
            "per": int(to_num(p["Staff per day"])),
            "len": to_num(p["Shift hours"]),
            "narc": int(to_num(p["Min narcotic aware"])),
            "pc": int(to_num(p["Min PC competent"])),
            "days": {d for d in ALL_DAYS if is_yes(p[d])},
        })
    rota = {p["label"]: {} for p in phs}
    warnings, bad, notes = [], set(), []
    willingness = willingness or {}

    def is_home(n, p):
        return by_name[n]["Preferred town"].strip().lower() in p["match"]

    def willing_for(n, p):
        w = willingness.get((n, p["key"]))
        return (WILL_HOME if is_home(n, p) else WILL_OTHER) if w is None else w

    def skills(names):
        return (sum(is_yes(by_name[n]["Narcotic awareness"]) for n in names),
                sum(is_yes(by_name[n]["PC competence"]) for n in names))

    def pick(day, p, home_only):
        shift = rota[p["label"]][day]
        slots_left = p["per"] - len(shift)
        narc, pc = skills(shift)
        need_narc, need_pc = max(0, p["narc"] - narc), max(0, p["pc"] - pc)
        urgent = slots_left <= need_narc + need_pc
        best, best_score = None, None
        for n, s in by_name.items():
            if (day in worked[n] or day not in avail[n] or dates[day] in hols[n]
                    or remaining[n] < p["len"] - 1e-9):
                continue
            home = is_home(n, p)
            if home_only and not home:
                continue
            will = willing_for(n, p)
            if will <= 0:
                continue                                  # 0 = never roster here
            score = (50 if home else 0) + weight * will
            if home:
                score -= 0.3 * to_num(s["Minutes from pharmacy"])  # shorter journey first
            if need_narc and is_yes(s["Narcotic awareness"]):
                score += 60 if urgent else 20
            if need_pc and is_yes(s["PC competence"]):
                score += 60 if urgent else 20
            score += 10 * remaining[n] / max(contracted[n], 1)    # spread hours
            if best_score is None or score > best_score:
                best, best_score = n, score
        return best

    for day in ALL_DAYS:
        open_ph = [p for p in phs if day in p["days"]]
        for p in open_ph:
            rota[p["label"]][day] = []
        for home_only in (True, False):
            for p in open_ph:
                while len(rota[p["label"]][day]) < p["per"]:
                    n = pick(day, p, home_only)
                    if n is None:
                        break
                    rota[p["label"]][day].append(n)
                    worked[n].add(day)
                    remaining[n] -= p["len"]
                    shifts[n] += 1
                    if willing_for(n, p) <= 1:
                        notes.append(f"Note: {day} {dates[day]:%d/%m} - {n} at {p['label']} "
                                     f"(willingness {willing_for(n, p)})")
        for p in open_ph:
            shift = rota[p["label"]][day]
            narc, pc = skills(shift)
            problems = []
            if len(shift) < p["per"]:
                problems.append(f"short staffed ({len(shift)}/{p['per']})")
            if narc < p["narc"]:
                problems.append(f"narcotic-aware {narc}/{p['narc']}")
            if pc < p["pc"]:
                problems.append(f"PC-competent {pc}/{p['pc']}")
            if problems:
                bad.add(p["label"])
                warnings.append(f"{day} {dates[day]:%d/%m} - {p['label']}: " + ", ".join(problems))

    warnings += notes
    info = {n: {"assigned": contracted[n] - remaining[n], "shifts": shifts[n],
                "holiday_days": sum(dates[d] in hols[n] for d in ALL_DAYS)}
            for n in by_name}
    return rota, info, warnings, bad


# --------------------------------------------------------------------------- #
# Application
# --------------------------------------------------------------------------- #
class RotaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Staff Rota")
        self.geometry("1200x720")
        self.minsize(850, 520)
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Treeview", rowheight=24)
        style.configure("Treeview.Heading", font=("TkDefaultFont", 10, "bold"))

        self.dirty = False
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.staff_tab, self.ph_tab, self.will_tab, self.rota_tab = (ttk.Frame(nb) for _ in range(4))
        nb.add(self.staff_tab, text="  Staff  ")
        nb.add(self.ph_tab, text="  Pharmacies  ")
        nb.add(self.will_tab, text="  Willingness  ")
        nb.add(self.rota_tab, text="  Rota  ")
        self.willing = {}                     # (name, ph_key) -> 0..5
        self._build_staff_tab()
        self._build_pharmacy_tab()
        self._build_willing_tab()
        self._build_rota_tab()
        nb.bind("<<NotebookTabChanged>>",
                lambda e: self.refresh_willing() if nb.select() == str(self.will_tab) else None)

        self.status = tk.StringVar(value="Import a CSV/XLSX file or add staff manually.")
        ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken").pack(fill="x", side="bottom")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- Staff tab ----------------
    def _build_staff_tab(self):
        bar = ttk.Frame(self.staff_tab)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="Import…", command=self.import_staff).pack(side="left", padx=2)
        ttk.Button(bar, text="Add row", command=self.add_staff).pack(side="left", padx=2)
        ttk.Button(bar, text="Delete selected", command=lambda: self.delete_rows(self.staff_tree)).pack(side="left", padx=2)
        ttk.Button(bar, text="Save as CSV…", command=lambda: self.save_tree(self.staff_tree, "staff.csv")).pack(side="left", padx=2)
        ttk.Label(bar, text="  Double-click a cell to edit · click a heading to sort").pack(side="left")
        ttk.Label(self.staff_tab, foreground="#555",
                  text="Days available: e.g. 'Mon-Fri' or 'Mon, Wed, Sat' (blank = any day).   "
                       "Holidays: e.g. '24/12/2026 - 02/01/2027; 15/11/2026'").pack(fill="x", padx=4)

        frame = ttk.Frame(self.staff_tab)
        frame.pack(fill="both", expand=True)
        self.staff_tree = EditableTree(
            frame, COLUMNS, choices={c: ["Yes", "No"] for c in YESNO_COLS},
            numeric=NUMERIC_COLS, on_change=self._staff_changed, selectmode="extended")
        widths = {"Name": 150, "Minutes from pharmacy": 150, "Hours per week": 105,
                  "Preferred town": 125, "Narcotic awareness": 130, "PC competence": 110,
                  "Days available": 120, "Holidays": 240}
        for c, w in widths.items():
            self.staff_tree.column(c, width=w, anchor="w" if c in ("Name", "Holidays") else "center")
        scrolled(frame, self.staff_tree)
        self.staff_tree.bind("<Delete>", lambda e: self.delete_rows(self.staff_tree))

    def _staff_changed(self):
        self.dirty = True
        self.staff_tree.restripe()
        recs = self.staff_tree.records()
        total = sum(to_num(r["Hours per week"]) for r in recs)
        self.status.set(f"{len(recs)} staff · {fmt_num(total)} contracted hours per week")

    def import_staff(self):
        path = self._ask_open("Import staff list")
        if not path:
            return
        try:
            headers, rows = read_table(path)
            records, missing = map_rows(headers, rows, COLUMNS, ALIASES)
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return
        if "Name" in missing:
            messagebox.showerror("Import failed",
                                 "Couldn't find a 'Name' column.\nHeadings found: " + ", ".join(headers))
            return
        records = [clean_staff(r) for r in records if r["Name"]]
        if not self._replace_or_append(self.staff_tree):
            return
        for rec in records:
            self.staff_tree.insert("", "end", values=[rec[c] for c in COLUMNS])
        self._staff_changed()
        msg = f"Imported {len(records)} staff from {path}"
        missing = [c for c in missing if c in REQUIRED_COLS]
        if missing:
            msg += f"  (missing columns left blank: {', '.join(missing)})"
        self.status.set(msg)
        self.add_pharmacies_from_staff(quiet=True)

    def add_staff(self):
        iid = self.staff_tree.insert("", "end", values=["", "", "", "", "No", "No", "", ""])
        self.staff_tree.see(iid)
        self.staff_tree.selection_set(iid)
        self._staff_changed()

    # ---------------- Pharmacies tab ----------------
    def _build_pharmacy_tab(self):
        bar = ttk.Frame(self.ph_tab)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="Add pharmacy", command=self.add_pharmacy).pack(side="left", padx=2)
        ttk.Button(bar, text="Add from staff towns", command=self.add_pharmacies_from_staff).pack(side="left", padx=2)
        ttk.Button(bar, text="Delete selected", command=lambda: self.delete_rows(self.ph_tree)).pack(side="left", padx=2)
        ttk.Button(bar, text="Load…", command=self.load_pharmacies).pack(side="left", padx=2)
        ttk.Button(bar, text="Save as CSV…", command=lambda: self.save_tree(self.ph_tree, "pharmacies.csv")).pack(side="left", padx=2)
        ttk.Label(self.ph_tab, foreground="#555",
                  text="Every setting is per pharmacy - double-click to change. Staff are matched to a "
                       "pharmacy when their Preferred town equals its Town (or Pharmacy name).").pack(fill="x", padx=4)

        frame = ttk.Frame(self.ph_tab)
        frame.pack(fill="both", expand=True)
        self.ph_tree = EditableTree(
            frame, PH_COLS, choices={d: ["Yes", "No"] for d in ALL_DAYS},
            numeric=PH_NUMERIC, on_change=self.ph_tree_changed, selectmode="extended")
        widths = {"No.": 45, "Pharmacy": 170, "Town": 130, "Staff per day": 100, "Shift hours": 90,
                  "Min narcotic aware": 135, "Min PC competent": 125}
        for c in PH_COLS:
            self.ph_tree.column(c, width=widths.get(c, 50),
                                anchor="w" if c in ("Pharmacy", "Town") else "center")
        scrolled(frame, self.ph_tree)
        self.ph_tree.bind("<Delete>", lambda e: self.delete_rows(self.ph_tree))

    def ph_tree_changed(self):
        self.dirty = True
        self.ph_tree.restripe()

    def _next_ph_no(self):
        nums = [to_num(r["No."]) for r in self.ph_tree.records()]
        return fmt_num(int(max(nums, default=0)) + 1)

    def _insert_pharmacy(self, name="", town=""):
        rec = dict(PH_DEFAULTS, **{"No.": self._next_ph_no(), "Pharmacy": name, "Town": town})
        iid = self.ph_tree.insert("", "end", values=[rec[c] for c in PH_COLS])
        self.ph_tree_changed()
        return iid

    def add_pharmacy(self):
        iid = self._insert_pharmacy()
        self.ph_tree.see(iid)
        self.ph_tree.selection_set(iid)

    def add_pharmacies_from_staff(self, quiet=False):
        known = set()
        for r in self.ph_tree.records():
            known |= {r["Town"].strip().lower(), r["Pharmacy"].strip().lower()}
        added = []
        for r in self.staff_tree.records():
            t = r["Preferred town"].strip()
            if t and t.lower() not in known:
                known.add(t.lower())
                self._insert_pharmacy(f"{t} Pharmacy", t)
                added.append(t)
        if added:
            self.status.set(f"Added pharmacies for: {', '.join(added)} - check their settings on the Pharmacies tab.")
        elif not quiet:
            messagebox.showinfo("Pharmacies", "Every staff town already has a pharmacy.")

    def load_pharmacies(self):
        path = self._ask_open("Load pharmacies")
        if not path:
            return
        try:
            headers, rows = read_table(path)
            aliases = {norm(c): c for c in PH_COLS}
            aliases.update({"no": "No.", "number": "No.", "name": "Pharmacy"})
            records, missing = map_rows(headers, rows, PH_COLS, aliases)
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            return
        if not self._replace_or_append(self.ph_tree):
            return
        for rec in records:
            for c in PH_COLS:
                if c in missing and c in PH_DEFAULTS:
                    rec[c] = PH_DEFAULTS[c]
                elif c in ALL_DAYS:
                    rec[c] = yes_no(rec[c])
            if not rec["No."]:
                rec["No."] = self._next_ph_no()
            self.ph_tree.insert("", "end", values=[rec[c] for c in PH_COLS])
            self.ph_tree_changed()
        self.status.set(f"Loaded {len(records)} pharmacies from {path}")

    # ---------------- Willingness tab ----------------
    def _build_willing_tab(self):
        bar = ttk.Frame(self.will_tab)
        bar.pack(fill="x", pady=4)
        ttk.Label(bar, text="Importance of willingness:").pack(side="left")
        self.will_weight = tk.IntVar(value=5)
        ttk.Spinbox(bar, from_=0, to=10, textvariable=self.will_weight, width=4).pack(side="left", padx=4)
        ttk.Label(bar, text="(0 = ignore, 10 = strongest)").pack(side="left")
        ttk.Button(bar, text="Reset all to auto", command=self.reset_willing).pack(side="left", padx=(16, 2))
        ttk.Button(bar, text="Load…", command=self.load_willing).pack(side="left", padx=2)
        ttk.Button(bar, text="Save as CSV…", command=self.save_willing).pack(side="left", padx=2)
        ttk.Label(self.will_tab, foreground="#555",
                  text="Score each person for each pharmacy: 5 = very willing, 3 = neutral, 1 = reluctant, "
                       "0 = never roster there.  'auto' = 5 for their preferred town, 3 elsewhere."
                  ).pack(fill="x", padx=4)
        self.will_frame = ttk.Frame(self.will_tab)
        self.will_frame.pack(fill="both", expand=True)
        self.will_tree, self._will_cols = None, {}

    def refresh_willing(self):
        for w in self.will_frame.winfo_children():
            w.destroy()
        phs = self.ph_tree.records()
        self._will_cols = {f"ph{i}": ph_key(p) for i, p in enumerate(phs)}
        cols = ["Name"] + list(self._will_cols)
        self.will_tree = EditableTree(
            self.will_frame, cols, readonly=["Name"],
            choices={c: WILL_CHOICES for c in self._will_cols}, on_change=self._willing_changed)
        self.will_tree.column("Name", width=160)
        for (col, key), p in zip(self._will_cols.items(), phs):
            self.will_tree.heading(col, text=pharmacy_label(p))
            self.will_tree.column(col, width=140, anchor="center")
        for s in self.staff_tree.records():
            n = s["Name"].strip()
            if n:
                self.will_tree.insert("", "end", values=[n] + [
                    str(self.willing.get((n, k), "auto")) for k in self._will_cols.values()])
        self.will_tree.restripe()
        scrolled(self.will_frame, self.will_tree)

    def _willing_changed(self):
        for rec in self.will_tree.records():
            for col, key in self._will_cols.items():
                v = rec[col]
                if v in ("", "auto"):
                    self.willing.pop((rec["Name"], key), None)
                else:
                    self.willing[(rec["Name"], key)] = int(v)
        self.dirty = True
        self.will_tree.restripe()

    def reset_willing(self):
        if messagebox.askyesno("Willingness", "Set every willingness score back to 'auto'?"):
            self.willing.clear()
            self.refresh_willing()

    def save_willing(self):
        self.refresh_willing()
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                            initialfile="willingness.csv")
        if not path:
            return
        cols = self.will_tree["columns"]
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow([self.will_tree.heading(c, "text") for c in cols])
            for iid in self.will_tree.get_children():
                w.writerow(self.will_tree.item(iid, "values"))
        self.status.set(f"Saved willingness to {path}")

    def load_willing(self):
        path = self._ask_open("Load willingness")
        if not path:
            return
        try:
            headers, rows = read_table(path)
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            return
        lookup = {}
        for p in self.ph_tree.records():
            for name in (pharmacy_label(p), p["No."], p["Pharmacy"], p["Town"]):
                if str(name).strip():
                    lookup.setdefault(norm(name), ph_key(p))
        matched = {i: lookup[norm(h)] for i, h in enumerate(headers) if i and norm(h) in lookup}
        unmatched = [h for i, h in enumerate(headers) if i and i not in matched]
        count = 0
        for r in rows:
            n = str(r[0]).strip() if r else ""
            for i, key in matched.items():
                v = str(r[i]).strip() if i < len(r) else ""
                if v.isdigit() and 0 <= int(v) <= 5:
                    self.willing[(n, key)] = int(v)
                    count += 1
                else:
                    self.willing.pop((n, key), None)
        self.refresh_willing()
        msg = f"Loaded {count} willingness scores from {path}"
        if unmatched:
            msg += f"  (no pharmacy found for: {', '.join(unmatched)})"
        self.status.set(msg)

    # ---------------- Rota tab ----------------
    def _build_rota_tab(self):
        bar = ttk.Frame(self.rota_tab, padding=(0, 4))
        bar.pack(fill="x")
        ttk.Label(bar, text="Week commencing (dd/mm/yyyy):").pack(side="left")
        today = date.today()
        monday = today + timedelta(days=(7 - today.weekday()) % 7 or 7)
        self.week_var = tk.StringVar(value=monday.strftime("%d/%m/%Y"))
        ttk.Entry(bar, textvariable=self.week_var, width=12).pack(side="left", padx=4)
        ttk.Button(bar, text="Generate rota", command=self.generate).pack(side="left", padx=(16, 4))
        ttk.Button(bar, text="Export rota (Excel / CSV)…", command=self.export_rota).pack(side="left", padx=4)
        self.rota_week = None

        paned = ttk.PanedWindow(self.rota_tab, orient="vertical")
        paned.pack(fill="both", expand=True)
        self.rota_frame = ttk.Frame(paned)
        self.rota_tree = None
        paned.add(self.rota_frame, weight=3)

        bottom = ttk.Frame(paned)
        paned.add(bottom, weight=2)
        hrs = ttk.LabelFrame(bottom, text="Hours this week", padding=4)
        hrs.pack(side="left", fill="both", expand=True)
        hcols = ["Name", "Contracted", "Rostered", "Shifts", "Holiday days", "Unfilled"]
        self.hours_tree = EditableTree(hrs, hcols, readonly=hcols, height=6)
        for c in hcols:
            self.hours_tree.column(c, width=150 if c == "Name" else 85,
                                   anchor="w" if c == "Name" else "center")
        scrolled(hrs, self.hours_tree)

        warn = ttk.LabelFrame(bottom, text="Warnings", padding=4)
        warn.pack(side="left", fill="both", expand=True, padx=(6, 0))
        self.warn_text = tk.Text(warn, height=6, wrap="word", foreground="#a00000")
        self.warn_text.pack(fill="both", expand=True)

    def generate(self):
        staff = [s for s in self.staff_tree.records() if s["Name"].strip()]
        if not staff:
            messagebox.showinfo("Rota", "Add or import some staff first.")
            return
        names = [s["Name"] for s in staff]
        dupes = sorted({n for n in names if names.count(n) > 1})
        if dupes:
            messagebox.showwarning("Rota", "Duplicate names - please make them unique:\n" + ", ".join(dupes))
            return
        if not self.ph_tree.get_children():
            self.add_pharmacies_from_staff(quiet=True)
        pharmacies = self.ph_tree.records()
        if not pharmacies:
            messagebox.showinfo("Rota", "Add at least one pharmacy on the Pharmacies tab.")
            return
        try:
            week = parse_date(self.week_var.get().strip())
        except ValueError as e:
            messagebox.showerror("Rota", str(e))
            return
        try:
            weight = 4 * max(0, min(10, int(self.will_weight.get())))
        except (tk.TclError, ValueError):
            weight = 20
        try:
            rota, info, warnings, bad = generate_rota(staff, pharmacies, week, self.willing, weight)
        except ValueError as e:      # a bad holiday entry
            messagebox.showerror("Rota", str(e))
            return

        open_days = [d for d in ALL_DAYS if any(d in rota[l] for l in rota)]
        self.rota_week = week
        self._make_rota_tree(open_days, week)
        for p in pharmacies:
            label = pharmacy_label(p)
            vals = [label] + [", ".join(rota[label][d]) if d in rota[label] else "Closed"
                              for d in open_days]
            self.rota_tree.insert("", "end", values=vals, tags=("warn",) if label in bad else ())
        self.rota_tree.restripe()

        self.hours_tree.delete(*self.hours_tree.get_children())
        for s in staff:
            n, c = s["Name"], to_num(s["Hours per week"])
            i = info[n]
            self.hours_tree.insert("", "end", values=[
                n, fmt_num(c), fmt_num(i["assigned"]), i["shifts"], i["holiday_days"],
                fmt_num(c - i["assigned"])])
        self.hours_tree.restripe()

        self.warn_text.delete("1.0", "end")
        self.warn_text.insert("end", "\n".join(warnings) if warnings else "No problems - all shifts covered.")
        self.status.set(f"Rota for w/c {week:%d/%m/%Y}: {len(pharmacies)} pharmacies, "
                        f"{len(warnings)} warning(s). Double-click a rota cell to adjust it.")

    def _make_rota_tree(self, days, week):
        for w in self.rota_frame.winfo_children():
            w.destroy()
        self.rota_tree = EditableTree(self.rota_frame, ["Pharmacy"] + days, readonly=["Pharmacy"])
        self.rota_tree.column("Pharmacy", width=170)
        for d in days:
            dt = week + timedelta(days=ALL_DAYS.index(d))
            self.rota_tree.heading(d, text=f"{d} {dt:%d/%m}")
            self.rota_tree.column(d, width=170)
        scrolled(self.rota_frame, self.rota_tree)

    def export_rota(self):
        if not self.rota_tree or not self.rota_tree.get_children():
            messagebox.showinfo("Export", "Generate a rota first.")
            return
        wc = self.rota_week.strftime("%Y-%m-%d") if self.rota_week else "week"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=f"rota_wc_{wc}.xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        cols = self.rota_tree["columns"]
        header = [self.rota_tree.heading(c, "text") for c in cols]
        body = [list(self.rota_tree.item(i, "values")) for i in self.rota_tree.get_children()]
        try:
            if path.lower().endswith(".csv"):
                with open(path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(header)
                    w.writerows(body)
            else:
                warn = {r + 1 for r, iid in enumerate(self.rota_tree.get_children())
                        if "warn" in self.rota_tree.item(iid, "tags")}
                # one name per line inside each rota cell
                rota_rows = [header] + [[row[0]] + [v.replace(", ", "\n") for v in row[1:]]
                                        for row in body]
                hcols = self.hours_tree["columns"]
                hours_rows = [list(hcols)] + [list(self.hours_tree.item(i, "values"))
                                              for i in self.hours_tree.get_children()]
                msgs = [m for m in self.warn_text.get("1.0", "end").strip().splitlines() if m]
                write_xlsx(path, [
                    {"name": "Rota", "rows": rota_rows, "warn": warn, "wrap": True,
                     "widths": [26] + [20] * (len(cols) - 1)},
                    {"name": "Hours", "rows": hours_rows, "widths": [24, 12, 12, 9, 13, 11]},
                    {"name": "Warnings", "rows": [["Warning"]] + [[m] for m in msgs], "widths": [90]},
                ])
        except OSError as e:
            messagebox.showerror("Export failed", str(e))
            return
        self.status.set(f"Exported rota to {path}")

    # ---------------- shared helpers ----------------
    def _ask_open(self, title):
        return filedialog.askopenfilename(
            title=title, filetypes=[("Spreadsheets", "*.csv *.xlsx"), ("CSV", "*.csv"),
                                    ("Excel", "*.xlsx"), ("All files", "*.*")])

    def _replace_or_append(self, tree):
        if not tree.get_children():
            return True
        ans = messagebox.askyesnocancel("Import", "Replace the current list?\n\nYes = replace, No = append")
        if ans is None:
            return False
        if ans:
            tree.delete(*tree.get_children())
        return True

    def delete_rows(self, tree):
        sel = tree.selection()
        if sel and messagebox.askyesno("Delete", f"Delete {len(sel)} row(s)?"):
            tree.delete(*sel)
            tree.on_change()

    def save_tree(self, tree, default_name):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                            initialfile=default_name)
        if not path:
            return False
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=tree["columns"])
            w.writeheader()
            w.writerows(tree.records())
        self.dirty = False
        self.status.set(f"Saved to {path}")
        return True

    def _on_close(self):
        if self.dirty and (self.staff_tree.get_children() or self.ph_tree.get_children()):
            ans = messagebox.askyesnocancel(
                "Quit", "You have unsaved changes. Quit anyway?\n\n"
                        "Yes = quit without saving, No = go back and save")
            if not ans:
                return
        self.destroy()


if __name__ == "__main__":
    RotaApp().mainloop()
