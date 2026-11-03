#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_excels.py — Build SOP Portable v12 starter workbooks.

Parses the source SOP HTML at:
    /home/z/my-project/download/SOP-DRV-001-DriverProductionSOP.html
and produces:
    edit/steps.xlsx     — DocumentInfo / Revisions / Phases / Steps (v12 schema)
    edit/content.xlsx   — supplemental content (all sections hidden, show=0)

Steps sheet columns:
    step_number, phase, station, title,
    action_1..action_8,
    tools, materials, acceptance_criteria,
    process_family, doc_number, section_responsible,
    step_version, step_date

Usage:
    python scripts/build_excels.py
    python scripts/build_excels.py --html <path> --out <dir>
"""
from __future__ import annotations

import argparse
import html as html_mod
import re
import sys
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except Exception as e:
    sys.stderr.write(f"[ERROR] openpyxl not available: {e}\n")
    OPENPYXL_OK = False

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFAULT_HTML = "/home/z/my-project/download/SOP-DRV-001-DriverProductionSOP.html"
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
DEFAULT_OUT_DIR = str(REPO_ROOT / "edit")

PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
ENGLISH_DIGITS = "0123456789"

# DocumentInfo defaults
DOC_INFO = [
    ("title",              "دستورالعمل استاندارد تولید درایور"),
    ("title_html",         "دستورالعمل استاندارد<br>تولید درایور"),
    ("doc_code",           "SOP-DRV-001"),
    ("revision",           "02"),
    ("issue_date",         "۱۴۰۳/۰۷/۱۲"),
    ("effective_date",     "۱۴۰۳/۰۷/۱۲"),
    ("department",         "تولید / مونتاژ"),
    ("process_owner",      "مهندسی فرآیند"),
    ("company",            "پیشگامان پاکران الکترونیک"),
    ("subtitle",           "سند کنترل‌شده فرآیند تولید."),
    ("scope",              "خط تولید درایور"),
    ("doc_type_prefix",    "WI"),
    ("last_revision_date", "۱۴۰۳/۰۷/۱۲"),
]

# Revisions sheet rows
REVISIONS = [
    ("v00", "۱۴۰۳/۰۴/۰۵", "کارشناس تولید",        "صدور اولیه دستورالعمل"),
    ("v01", "۱۴۰۳/۰۵/۲۰", "کارشناس تولید",        "افزودن فاز آزمون و کنترل کیفیت با جزئیات کامل"),
    ("v02", "۱۴۰۳/۰۷/۱۲", "کارشناس تولید + QC",   "بازنویسی کامل به همراه تصاویر راهنما و معیارهای پذیرش"),
]

# Phases sheet rows: letter, short_name, title, desc, owner, hold_point
PHASES_META = [
    {"letter": "A", "short_name": "پیش‌مونتاژ",
     "title": "فاز پیش‌مونتاژ — آماده‌سازی سیم، قاب و کارتن",
     "desc": "ایستگاه ۱ — مسئول: اپراتور مونتاژ. خروجی فاز: قاب پلاستیکی با گلندها و سیم‌ها آماده، کارتن با لیبل کد محصول و شرینک‌های بریده‌شده آماده.",
     "owner": "اپراتور مونتاژ", "hold_point": "بازرسی اولیه قاب و سیم‌ها"},
    {"letter": "B", "short_name": "مونتاژ اصلی",
     "title": "فاز مونتاژ اصلی — هیت‌سینک، برد و قطعات الکترونیکی",
     "desc": "ایستگاه ۲ — مسئول: اپراتور مونتاژ. خروجی فاز: هیت‌سینک نصب‌شده، برد الکترونیکی مونتاژشده، اتصالات لحیم‌شده و تست‌شده.",
     "owner": "اپراتور مونتاژ", "hold_point": "تست اتصالات و بازرسی لحیم‌ها"},
    {"letter": "C", "short_name": "آزمون و QC",
     "title": "فاز آزمون و کنترل کیفیت",
     "desc": "ایستگاه ۳ — مسئول: کارشناس کنترل کیفیت. خروجی فاز: درایور آزمون‌شده، نتایج ثبت‌شده در کارت QC، تأیید نهایی صادرشده.",
     "owner": "کارشناس کنترل کیفیت", "hold_point": "تست نهایی و صدور تأییدیه QC"},
    {"letter": "D", "short_name": "بسته‌بندی",
     "title": "فاز بسته‌بندی و ردیابی",
     "desc": "ایستگاه‌های ۴ و ۵ — مسئول: اپراتور بسته‌بندی و ردیابی. خروجی فاز: محصول در کارتن نهایی با لیبل ردیابی، آماده ارسال به انبار.",
     "owner": "اپراتور بسته‌بندی", "hold_point": "بازرسی نهایی بسته و لیبل ردیابی"},
]

# Per-letter defaults for steps
FAMILY_BY_LETTER = {"A": "ASM", "B": "ASM", "C": "QC", "D": "PKG"}
RESPONSIBLE_BY_LETTER = {
    "A": "اپراتور مونتاژ",
    "B": "اپراتور مونتاژ",
    "C": "کارشناس کنترل کیفیت",
    "D": "اپراتور بسته‌بندی",
}
STEP_VERSION_DEFAULT = "Rev.01"
STEP_DATE_DEFAULT = "۱۴۰۳/۰۷/۱۲"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def to_persian(value) -> str:
    return str(value).translate(str.maketrans(ENGLISH_DIGITS, PERSIAN_DIGITS))


def to_english(value: str) -> str:
    return str(value).translate(str.maketrans(PERSIAN_DIGITS, ENGLISH_DIGITS))


def _strip_tags(s: str) -> str:
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"</p>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = html_mod.unescape(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ---------------------------------------------------------------------------
# HTML parser
# ---------------------------------------------------------------------------
def _extract_step_card_blocks(text: str) -> list:
    """Find every <div class="step-card"> ... </div> top-level block."""
    blocks = []
    for m in re.finditer(r'<div\s+class="step-card"[^>]*>', text):
        i = m.end()
        depth = 1
        start_inner = i
        while i < len(text):
            om = re.search(r"<div\b", text[i:])
            cm = re.search(r"</div\s*>", text[i:])
            if not cm:
                break
            if om and om.start() < cm.start():
                depth += 1
                i += om.end()
            else:
                depth -= 1
                if depth == 0:
                    blocks.append((m.start(), text[start_inner:i]))
                    break
                i += cm.end()
    return blocks


def parse_step(block_inner: str) -> dict:
    s = {"step_number": 0, "phase": "A", "station": "",
         "title": "", "actions": [], "tools": "",
         "materials": "", "acceptance": ""}
    m = re.search(r'<div\s+class="step-number"[^>]*>([^<]+)</div>', block_inner)
    if m:
        s["step_number"] = int(to_english(m.group(1).strip()).split(".")[0].split()[0] or 0)
    m = re.search(r'<h4\s+class="step-title"[^>]*>(.*?)</h4>', block_inner, re.S)
    if m:
        s["title"] = _strip_tags(m.group(1))
    m = re.search(r'<div\s+class="step-station"[^>]*>(.*?)</div>', block_inner, re.S)
    if m:
        s["station"] = _strip_tags(m.group(1))
    actions = []
    for am in re.finditer(r'<li\s+class="numbered"[^>]*>(.*?)</li>', block_inner, re.S):
        t = _strip_tags(am.group(1))
        if t:
            actions.append(t)
    s["actions"] = actions
    resources = re.findall(
        r'<div\s+class="resource-block"[^>]*>\s*<div\s+class="label"[^>]*>(.*?)</div>\s*<div\s+class="value"[^>]*>(.*?)</div>',
        block_inner, re.S)
    for label, value in resources:
        lbl = _strip_tags(label)
        val = _strip_tags(value)
        if "ابزار" in lbl:
            s["tools"] = val
        elif "مواد" in lbl or "ملزومات" in lbl:
            s["materials"] = val
    am = re.search(r'<div\s+class="step-acceptance"[^>]*>(.*?)</div>\s*</div>\s*$', block_inner, re.S)
    if not am:
        am = re.search(r'<div\s+class="step-acceptance"[^>]*>(.*)$', block_inner, re.S)
    if am:
        acc_html = am.group(1)
        acc_html = re.sub(r'<span\s+class="label"[^>]*>[^<]*</span>', "", acc_html)
        s["acceptance"] = _strip_tags(acc_html)
    return s


def parse_phase_blocks(text: str) -> list:
    """Find each <div class='phase-block'> block and extract letter/title/desc."""
    out = []
    pat = re.compile(
        r'<div\s+class="phase-block"[^>]*>\s*'
        r'<div\s+class="phase-number"[^>]*>([^<]+)</div>\s*'
        r'<div\s+class="phase-text"[^>]*>(.*?)</div>\s*</div>',
        re.S)
    for m in pat.finditer(text):
        letter = m.group(1).strip().upper()[:1]
        inner = m.group(2)
        tm = re.search(r'<h3\s+class="phase-title"[^>]*>(.*?)</h3>', inner, re.S)
        dm = re.search(r'<div\s+class="phase-desc"[^>]*>(.*?)</div>', inner, re.S)
        title = _strip_tags(tm.group(1)) if tm else ""
        desc = _strip_tags(dm.group(1)) if dm else ""
        out.append({"pos": m.start(), "letter": letter, "title": title, "desc": desc})
    return out


def parse_html(html_path: str) -> tuple:
    text = Path(html_path).read_text(encoding="utf-8")
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.S | re.I)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.S | re.I)
    text = re.sub(r"<!--.*?-->", "", text, flags=re.S)

    phase_events = parse_phase_blocks(text)
    blocks = _extract_step_card_blocks(text)
    steps = []
    for pos, inner in blocks:
        s = parse_step(inner)
        if not s["step_number"]:
            continue
        last_letter = "A"
        for p in phase_events:
            if p["pos"] < pos:
                last_letter = p["letter"]
            else:
                break
        s["phase"] = last_letter
        steps.append(s)

    steps.sort(key=lambda d: d["step_number"])
    return phase_events, steps


# ---------------------------------------------------------------------------
# Excel styling
# ---------------------------------------------------------------------------
HEADER_FONT = Font(name="Tahoma", bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1E3A5F")
THIN_BORDER = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="top", wrap_text=True)


def style_header(ws, row: int = 1, ncols: int = 4) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER


def autosize(ws, widths: list) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def box_body(ws, max_col: int, max_row: int) -> None:
    for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.border = THIN_BORDER
            cell.alignment = ALIGN_RIGHT


# ---------------------------------------------------------------------------
# Workbook writers
# ---------------------------------------------------------------------------
def write_steps_workbook(out_path: str, steps: list) -> None:
    wb = openpyxl.Workbook()

    # --- Sheet 1: DocumentInfo ---
    ws = wb.active
    ws.title = "DocumentInfo"
    ws.append(["Key", "Value"])
    style_header(ws, 1, 2)
    for k, v in DOC_INFO:
        ws.append([k, v])
    box_body(ws, 2, ws.max_row)
    autosize(ws, [22, 50])

    # --- Sheet 2: Revisions ---
    ws = wb.create_sheet("Revisions")
    ws.append(["rev", "date", "author", "summary"])
    style_header(ws, 1, 4)
    for r in REVISIONS:
        ws.append(list(r))
    box_body(ws, 4, ws.max_row)
    autosize(ws, [10, 14, 24, 60])

    # --- Sheet 3: Phases ---
    ws = wb.create_sheet("Phases")
    ws.append(["letter", "short_name", "title", "desc", "owner", "hold_point"])
    style_header(ws, 1, 6)
    step_counts = {L: 0 for L in "ABCD"}
    for s in steps:
        step_counts[s["phase"]] = step_counts.get(s["phase"], 0) + 1
    for p in PHASES_META:
        ws.append([
            p["letter"], p["short_name"], p["title"], p["desc"],
            p["owner"], p["hold_point"],
        ])
    box_body(ws, 6, ws.max_row)
    autosize(ws, [8, 16, 50, 60, 24, 28])

    # --- Sheet 4: Steps (v12 schema) ---
    ws = wb.create_sheet("Steps")
    header = [
        "step_number", "phase", "station", "title",
        "action_1", "action_2", "action_3", "action_4",
        "action_5", "action_6", "action_7", "action_8",
        "tools", "materials", "acceptance_criteria",
        "process_family", "doc_number", "section_responsible",
        "step_version", "step_date",
    ]
    ws.append(header)
    style_header(ws, 1, len(header))

    for s in steps:
        ph = s["phase"]
        actions = (s["actions"] + [""] * 8)[:8]    # pad/truncate to 8
        doc_number = f"{s['step_number']:03d}"     # 001-038
        row = [
            s["step_number"],
            ph,
            s["station"] or "",
            s["title"],
            *actions,
            s["tools"],
            s["materials"],
            s["acceptance"],
            FAMILY_BY_LETTER.get(ph, "ASM"),
            doc_number,
            RESPONSIBLE_BY_LETTER.get(ph, ""),
            STEP_VERSION_DEFAULT,
            STEP_DATE_DEFAULT,
        ]
        ws.append(row)
    box_body(ws, len(header), ws.max_row)
    autosize(ws, [
        7,   # step_number
        6,   # phase
        12,  # station
        36,  # title
        30, 30, 30, 30, 30, 30, 30, 30,   # action_1..8
        28,  # tools
        28,  # materials
        40,  # acceptance_criteria
        10,  # process_family
        10,  # doc_number
        22,  # section_responsible
        8,   # step_version
        12,  # step_date
    ])
    ws.freeze_panes = "A2"
    wb["Phases"].freeze_panes = "A2"
    wb.save(out_path)


def write_content_workbook(out_path: str) -> None:
    """Create content.xlsx — all sections hidden (show=0)."""
    wb = openpyxl.Workbook()
    # Section definitions: (sheet_name, rows)
    sections = [
        ("Intro", [
            ("intro", "purpose",        "این دستورالعمل فرآیند تولید درایور را در چهار فاز شرح می‌دهد."),
            ("intro", "scope",          "تمامی خطوط تولید درایور در کارخانه."),
            ("intro", "responsibility", "سرپرست تولید مسئول اجرای این دستورالعمل است."),
        ]),
        ("Notes", [
            ("notes", "general", "یادداشت‌های آزاد برای توسعه آینده."),
        ]),
        ("Scope", [
            ("scope", "in_scope",   "خط تولید درایور شامل مونتاژ، آزمون و بسته‌بندی."),
            ("scope", "out_scope",  "تولید قطعات خام و PCB خارج از محدوده این سند است."),
        ]),
    ]
    for idx, (name, rows) in enumerate(sections):
        if idx == 0:
            ws = wb.active
            ws.title = name
        else:
            ws = wb.create_sheet(name)
        ws.append(["section", "key", "value", "show"])
        style_header(ws, 1, 4)
        for r in rows:
            ws.append([r[0], r[1], r[2], 0])    # show=0 (hidden)
        box_body(ws, 4, ws.max_row)
        autosize(ws, [12, 18, 60, 8])
        ws.sheet_state = "hidden"               # sheet hidden too

    # Add a visible "Index" sheet so the workbook opens cleanly
    ws_idx = wb.create_sheet("Index")
    ws_idx.append(["Sheet", "Purpose", "Visible"])
    style_header(ws_idx, 1, 3)
    for name, _ in sections:
        ws_idx.append([name, "متون پشتیبان (مخفی)", "hidden"])
    box_body(ws_idx, 3, ws_idx.max_row)
    autosize(ws_idx, [12, 40, 12])

    wb.save(out_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="Build SOP Portable v12 starter Excels")
    ap.add_argument("--html", default=DEFAULT_HTML, help="source SOP HTML path")
    ap.add_argument("--out", default=DEFAULT_OUT_DIR, help="output directory (edit/)")
    args = ap.parse_args()

    if not OPENPYXL_OK:
        sys.stderr.write("openpyxl not installed. Install with: pip install openpyxl\n")
        return 1

    html_path = Path(args.html)
    if not html_path.exists():
        sys.stderr.write(f"HTML not found: {html_path}\n")
        return 1

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Parsing: {html_path}")
    phase_events, steps = parse_html(str(html_path))
    print(f"  → phases: {len(phase_events)}")
    for p in phase_events:
        cnt = sum(1 for s in steps if s["phase"] == p["letter"])
        print(f"    {p['letter']}: {p['title'][:50]}... ({cnt} steps)")
    print(f"  → steps:  {len(steps)}")
    if steps:
        print(f"    first: #{steps[0]['step_number']} '{steps[0]['title'][:50]}'")
        print(f"    last:  #{steps[-1]['step_number']} '{steps[-1]['title'][:50]}'")

    steps_path = out_dir / "steps.xlsx"
    content_path = out_dir / "content.xlsx"

    print(f"\nWriting: {steps_path}")
    write_steps_workbook(str(steps_path), steps)

    print(f"Writing: {content_path}")
    write_content_workbook(str(content_path))

    print("\n✓ All workbooks created successfully.")
    print(f"  Output dir: {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
