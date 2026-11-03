#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOP Portable v12 — CLI PDF generator (no web UI, no server).

Usage (typical):
    python app/generate_sop.py \
        --steps     edit/steps.xlsx \
        --content   edit/content.xlsx \
        --images    images \
        --output    output/sop.pdf \
        --template-dir app

The launcher run.bat supplies all these paths automatically.

Pipeline
--------
1.  Read steps.xlsx  — DocumentInfo / Revisions / Phases / Steps sheets.
2.  Read content.xlsx (optional — currently informational only).
3.  Scan images/ for files matching ^step-(\\d+)-(\\d+)\\.(png|jpg|jpeg|webp)
4.  Normalize each image (EXIF transpose → RGB → center-crop 4:3 →
    resize if larger than 1920x1440 → JPEG q92) into output/assets/.
5.  Copy logo.png + fonts/ into output/ so the rendered HTML resolves
    every relative URL via the file:// protocol.
6.  Render template.html (Jinja2) with embedded styles.css → output/sop.html
7.  Render PDF with Playwright (A4, 10mm margins, print_background=True,
    wait_until="domcontentloaded" + 2s settle for fonts).
8.  If Playwright crashes (V8 string-size limit on long SOPs), fall back
    to chunked rendering (6 steps per chunk) stitched with pypdf.

Console progress is printed at every stage.
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

VERSION = "v12"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
ENGLISH_DIGITS = "0123456789"

DEFAULT_PHASES = [
    {"letter": "A", "short_name": "پیش‌مونتاژ",
     "title": "فاز پیش‌مونتاژ — آماده‌سازی سیم، قاب و کارتن"},
    {"letter": "B", "short_name": "مونتاژ اصلی",
     "title": "فاز مونتاژ اصلی — هیت‌سینک، برد و قطعات الکترونیکی"},
    {"letter": "C", "short_name": "آزمون و QC",
     "title": "فاز آزمون و کنترل کیفیت"},
    {"letter": "D", "short_name": "بسته‌بندی",
     "title": "فاز بسته‌بندی و ردیابی"},
]
PHASE_FAMILY_DEFAULT = {"A": "ASM", "B": "ASM", "C": "QC", "D": "PKG"}

DEFAULT_DOC = {
    "title":              "دستورالعمل استاندارد تولید درایور",
    "title_html":         "دستورالعمل استاندارد<br>تولید درایور",
    "doc_code":           "SOP-DRV-001",
    "revision":           "02",
    "issue_date":         "۱۴۰۳/۰۷/۱۲",
    "effective_date":     "۱۴۰۳/۰۷/۱۲",
    "department":         "تولید / مونتاژ",
    "process_owner":      "مهندسی فرآیند",
    "company":            "پیشگامان پاکران الکترونیک",
    "subtitle":           "سند کنترل‌شده فرآیند تولید.",
    "scope":              "خط تولید درایور",
    "doc_type_prefix":    "WI",
    "last_revision_date": "۱۴۰۳/۰۷/۱۲",
}

IMG_PATTERN = re.compile(r"^step-(\d+)-(\d+)\.(png|jpg|jpeg|webp)$", re.I)
MAX_IMG_W, MAX_IMG_H = 1920, 1440          # 4:3 cap
CHUNK_SIZE = 6                              # chunked fallback


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def to_persian_digits(value) -> str:
    s = str(value)
    return s.translate(str.maketrans(ENGLISH_DIGITS, PERSIAN_DIGITS))


def to_english_digits(value) -> str:
    s = str(value)
    return s.translate(str.maketrans(PERSIAN_DIGITS, ENGLISH_DIGITS))


def safe_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s in ("None", "nan", ""):
        return ""
    return s


def log(msg: str) -> None:
    print(f"  {msg}", flush=True)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class DocInfo:
    title: str = ""
    title_html: str = ""
    doc_code: str = ""
    revision: str = ""
    issue_date: str = ""
    effective_date: str = ""
    department: str = ""
    process_owner: str = ""
    company: str = "پیشگامان پاکران الکترونیک"
    subtitle: str = ""
    scope: str = ""
    doc_type_prefix: str = "WI"
    last_revision_date: str = ""


@dataclass
class ImageSlot:
    position: int
    src_path: str            # relative to output HTML (e.g. assets/step-01-01.jpg)
    alt_text: str = ""


@dataclass
class Step:
    step_number: int = 0
    phase: str = "A"
    station: str = ""
    title: str = ""
    actions: list = field(default_factory=list)         # up to 8 items
    tools: str = ""
    materials: str = ""
    acceptance: str = ""
    image_slots: list = field(default_factory=list)     # list[ImageSlot]
    process_family: str = "ASM"
    doc_number: str = ""              # e.g. "001"
    section_responsible: str = ""
    step_version: str = "Rev.01"
    step_date: str = "۱۴۰۳/۰۷/۱۲"

    @property
    def num_persian(self) -> str:
        return to_persian_digits(self.step_number)

    @property
    def doc_code(self) -> str:
        """{family}-{number} — template prepends doc_type_prefix (WI-)."""
        fam = self.process_family or "STEP"
        num = self.doc_number or f"{self.step_number:03d}"
        return f"{fam}-{num}"

    @property
    def gallery_html(self) -> str:
        if not self.image_slots:
            return ""
        cells = []
        for slot in self.image_slots:
            pos_p = to_persian_digits(slot.position)
            cells.append(
                f'<div class="img-filled" data-pos="{pos_p}">'
                f'<img src="{slot.src_path}" alt="{slot.alt_text}"></div>'
            )
        return (
            '<div class="step-gallery">'
            '<div class="gallery-header">'
            '<span class="gallery-label">مرجع تصویری مرحله</span>'
            '</div>'
            '<div class="gallery-grid">'
            + "".join(cells) +
            '</div></div>'
        )


# ---------------------------------------------------------------------------
# Excel reading
# ---------------------------------------------------------------------------
def _load_wb(path: str):
    from openpyxl import load_workbook
    return load_workbook(path, data_only=True)


def _rows(ws):
    out = []
    for r in ws.iter_rows(values_only=True):
        if all(c is None or (isinstance(c, str) and not c.strip()) for c in r):
            continue
        out.append(list(r))
    return out


def _find_sheet(wb, candidates):
    names = {n.lower(): n for n in wb.sheetnames}
    for cand in candidates:
        cl = cand.lower()
        for nl, real in names.items():
            if cl in nl:
                return wb[real]
    return None


def read_document_info(steps_path: str) -> DocInfo:
    info = DocInfo(**DEFAULT_DOC)
    try:
        wb = _load_wb(steps_path)
        ws = _find_sheet(wb, ["DocumentInfo", "DocInfo", "Info"])
        if not ws:
            return info
        for row in _rows(ws):
            if len(row) < 2:
                continue
            k = safe_text(row[0]).lower().replace("-", "_").replace(" ", "_")
            v = safe_text(row[1])
            if not k or not v:
                continue
            if hasattr(info, k):
                setattr(info, k, v)
    except Exception as e:
        sys.stderr.write(f"[read_document_info] WARN: {e}\n")
    return info


def read_phases(steps_path: str) -> list:
    """Return list of phase dicts (with steps_count filled later)."""
    out = [dict(p) for p in DEFAULT_PHASES]
    try:
        wb = _load_wb(steps_path)
        ws = _find_sheet(wb, ["Phases", "Phase"])
        if not ws:
            return out
        rows = _rows(ws)
        if not rows:
            return out
        header = [safe_text(h).lower() for h in rows[0]]
        ci = {h: i for i, h in enumerate(header)}
        parsed = {}
        for row in rows[1:]:
            letter = safe_text(row[ci.get("letter", 0)] if "letter" in ci else "").upper()[:1]
            if letter not in "ABCD":
                continue
            parsed[letter] = {
                "letter": letter,
                "short_name": safe_text(row[ci["short_name"]] if "short_name" in ci and ci["short_name"] < len(row) else "") or next((p["short_name"] for p in DEFAULT_PHASES if p["letter"] == letter), ""),
                "title": safe_text(row[ci["title"]] if "title" in ci and ci["title"] < len(row) else "") or next((p["title"] for p in DEFAULT_PHASES if p["letter"] == letter), ""),
            }
        # Merge parsed into out (preserve A-D order)
        out = []
        for letter in "ABCD":
            if letter in parsed:
                out.append(parsed[letter])
            else:
                out.append(next(p for p in DEFAULT_PHASES if p["letter"] == letter))
    except Exception as e:
        sys.stderr.write(f"[read_phases] WARN: {e}\n")
    return out


def read_steps(steps_path: str, phases: list) -> list:
    steps = []
    family_by_letter = {p["letter"]: PHASE_FAMILY_DEFAULT.get(p["letter"], "ASM") for p in phases}
    try:
        wb = _load_wb(steps_path)
        ws = _find_sheet(wb, ["Steps", "Step"])
        if not ws:
            return steps
        rows = _rows(ws)
        if not rows:
            return steps
        header = [safe_text(h).lower() for h in rows[0]]
        ci = {h: i for i, h in enumerate(header)}

        def cell(row, *names):
            for n in names:
                if n in ci and ci[n] < len(row):
                    return safe_text(row[ci[n]])
            return ""

        for row in rows[1:]:
            num_raw = cell(row, "step_number", "step", "number", "شماره")
            if not num_raw:
                continue
            num = int(to_english_digits(num_raw).split()[0].split(".")[0] or 0)
            if num == 0:
                continue
            phase = (cell(row, "phase", "فاز") or "A").upper()[:1] or "A"
            actions = []
            for i in range(1, 9):
                a = cell(row, f"action_{i}", f"action{i}", f"اقدام_{i}")
                if a:
                    actions.append(a)
            if not actions:
                # Fallback to legacy single "actions" column
                legacy = cell(row, "actions", "action", "اقدامات")
                actions = [a.strip() for a in re.split(r"[•·▪;\n\r]+", legacy) if a.strip()]
            family = cell(row, "process_family", "processfamily", "family") or family_by_letter.get(phase, "ASM")
            steps.append(Step(
                step_number=num,
                phase=phase,
                station=cell(row, "station", "ایستگاه"),
                title=cell(row, "title", "عنوان"),
                actions=actions[:8],
                tools=cell(row, "tools", "ابزار"),
                materials=cell(row, "materials", "مواد", "ملزومات"),
                acceptance=cell(row, "acceptance_criteria", "acceptance", "criteria", "پذیرش"),
                process_family=family,
                doc_number=cell(row, "doc_number", "docnumber", "docno"),
                section_responsible=cell(row, "section_responsible", "sectionresponsible", "responsible", "مسئول"),
                step_version=cell(row, "step_version", "stepversion", "version", "نسخه") or "Rev.01",
                step_date=cell(row, "step_date", "stepdate", "date", "تاریخ") or "۱۴۰۳/۰۷/۱۲",
            ))
    except Exception as e:
        sys.stderr.write(f"[read_steps] ERROR: {e}\n{traceback.format_exc()}\n")
    steps.sort(key=lambda s: s.step_number)
    # Update phase step counts
    for p in phases:
        nums = [s.step_number for s in steps if s.phase == p["letter"]]
        p["steps_count"] = len(nums)
    return steps


# ---------------------------------------------------------------------------
# Image processing
# ---------------------------------------------------------------------------
def process_image(src: Path, dst: Path) -> bool:
    """EXIF transpose → RGB → center-crop 4:3 → resize if >1920x1440 → JPEG q92."""
    try:
        from PIL import Image, ImageOps
        img = Image.open(src)
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        w, h = img.size
        target_ratio = 4.0 / 3.0
        cur_ratio = w / h
        if cur_ratio > target_ratio:
            new_w = int(h * target_ratio)
            left = (w - new_w) // 2
            img = img.crop((left, 0, left + new_w, h))
        elif cur_ratio < target_ratio:
            new_h = int(w / target_ratio)
            top = (h - new_h) // 2
            img = img.crop((0, top, w, top + new_h))
        # Resize if too large
        if img.size[0] > MAX_IMG_W or img.size[1] > MAX_IMG_H:
            img.thumbnail((MAX_IMG_W, MAX_IMG_H), Image.LANCZOS)
        dst.parent.mkdir(parents=True, exist_ok=True)
        img.save(dst, "JPEG", quality=92, optimize=True)
        return True
    except Exception as e:
        sys.stderr.write(f"[process_image] {src} → {dst}: {e}\n")
        return False


def collect_images(steps: list, images_dir: str, output_dir: str) -> None:
    """Scan images/ for step-NN-PP.* files; process into output/assets/."""
    img_root = Path(images_dir)
    assets = Path(output_dir) / "assets"
    assets.mkdir(parents=True, exist_ok=True)

    by_step = {s.step_number: s for s in steps}
    if img_root.is_dir():
        for f in sorted(img_root.iterdir()):
            m = IMG_PATTERN.match(f.name)
            if not m:
                continue
            step_num = int(m.group(1))
            position = int(m.group(2))
            step = by_step.get(step_num)
            if not step:
                continue
            # Output filename keeps zero-padded NN-PP, always .jpg
            dst_name = f"step-{step_num:02d}-{position:02d}.jpg"
            dst = assets / dst_name
            if process_image(f, dst):
                step.image_slots.append(ImageSlot(
                    position=position,
                    src_path=f"assets/{dst_name}",
                    alt_text=f"تصویر {to_persian_digits(position)} — مرحله {to_persian_digits(step_num)}",
                ))
    # Sort each step's slots by position
    for s in steps:
        s.image_slots.sort(key=lambda x: x.position)


# ---------------------------------------------------------------------------
# Output dir prep
# ---------------------------------------------------------------------------
def prepare_output(output_path: str, template_dir: str) -> Path:
    out_dir = Path(output_path).resolve().parent
    out_dir.mkdir(parents=True, exist_ok=True)
    tdir = Path(template_dir).resolve()
    # Copy logo
    src_logo = tdir / "logo.png"
    if src_logo.exists():
        shutil.copy(src_logo, out_dir / "logo.png")
    # Copy fonts
    src_fonts = tdir / "fonts"
    dst_fonts = out_dir / "fonts"
    if dst_fonts.exists():
        shutil.rmtree(dst_fonts)
    if src_fonts.is_dir():
        shutil.copytree(src_fonts, dst_fonts)
    else:
        dst_fonts.mkdir()
    return out_dir


# ---------------------------------------------------------------------------
# Jinja2 rendering
# ---------------------------------------------------------------------------
def render_html(output_dir: Path, template_dir: str,
                steps: list, doc: DocInfo, phases: list) -> str:
    from jinja2 import Environment, FileSystemLoader, select_autoescape
    env = Environment(
        loader=FileSystemLoader(template_dir),
        autoescape=select_autoescape(["html", "xml"]),
        trim_blocks=True, lstrip_blocks=True,
    )
    template = env.get_template("template.html")
    html = template.render(doc=doc, steps=steps, phases=phases)
    out_html = output_dir / "sop.html"
    out_html.write_text(html, encoding="utf-8")
    return str(out_html)


# ---------------------------------------------------------------------------
# Playwright PDF
# ---------------------------------------------------------------------------
def _find_chromium_exe() -> Optional[str]:
    """Look for portable Chromium under PLAYWRIGHT_BROWSERS_PATH."""
    env_path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
    if not env_path:
        return None
    base = Path(env_path)
    if not base.is_dir():
        return None
    cands = sorted(base.glob("chromium-*"), reverse=True)
    for c in cands:
        for sub in ("chrome-win64", "chrome-win", "chrome-linux", "chrome-mac"):
            d = c / sub
            if d.is_dir():
                for exe_name in ("chrome.exe", "chrome"):
                    exe = d / exe_name
                    if exe.exists():
                        return str(exe)
    return None


def playwright_pdf(html_path: str, output_path: str) -> bool:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        sys.stderr.write(f"[playwright_pdf] playwright not installed: {e}\n")
        return False
    chrome_exe = _find_chromium_exe()
    launch_args = ["--no-sandbox", "--disable-dev-shm-usage", "--font-render-hinting=none"]
    try:
        with sync_playwright() as p:
            if chrome_exe:
                log(f"  launching chromium: {chrome_exe}")
                browser = p.chromium.launch(executable_path=chrome_exe, args=launch_args, headless=True)
            else:
                log("  launching system chromium")
                browser = p.chromium.launch(args=launch_args, headless=True)
            context = browser.new_context(viewport={"width": 794, "height": 1123})
            page = context.new_page()
            file_url = "file://" + str(Path(html_path).resolve()).replace("\\", "/")
            log(f"  navigating: {file_url}")
            page.goto(file_url, wait_until="domcontentloaded")
            # Wait for fonts
            try:
                page.evaluate("document.fonts && document.fonts.ready")
            except Exception:
                pass
            page.wait_for_timeout(2000)
            log("  rendering PDF (A4, 10mm margins)...")
            page.pdf(
                path=output_path,
                format="A4",
                margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
                print_background=True,
                prefer_css_page_size=True,
            )
            browser.close()
        return True
    except Exception as e:
        sys.stderr.write(f"[playwright_pdf] ERROR: {e}\n{traceback.format_exc()}\n")
        return False


# ---------------------------------------------------------------------------
# Chunked fallback
# ---------------------------------------------------------------------------
def chunked_pdf(html_path: str, output_path: str) -> bool:
    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfWriter, PdfReader  # type: ignore
        except ImportError as e:
            sys.stderr.write(f"[chunked_pdf] no pypdf: {e}\n")
            return False

    src_html = Path(html_path).read_text(encoding="utf-8")
    m = re.search(r"<body[^>]*>(.*)</body>", src_html, re.S | re.I)
    if not m:
        sys.stderr.write("[chunked_pdf] no <body> in template\n")
        return False
    body_inner = m.group(1)
    parts = re.split(r'(<div class="step-page-wrapper")', body_inner)
    prefix = parts[0]
    chunks_markup = []
    cur, cur_count = [], 0
    for i in range(1, len(parts), 2):
        marker = parts[i]
        content = parts[i + 1] if i + 1 < len(parts) else ""
        cur.append(marker + content)
        cur_count += 1
        if cur_count >= CHUNK_SIZE:
            chunks_markup.append("".join(cur))
            cur, cur_count = [], 0
    if cur:
        chunks_markup.append("".join(cur))

    head_match = re.search(r"(.*?)<body", src_html, re.S | re.I)
    head_part = (head_match.group(1) if head_match else "<!DOCTYPE html><html><head></head>") + "<body>"

    total = len(chunks_markup)
    log(f"  chunked fallback: {total} chunks of ≤{CHUNK_SIZE} steps")
    staging = Path(html_path).parent
    chunk_pdfs = []
    for i, markup in enumerate(chunks_markup):
        chunk_html = head_part + prefix + markup + "</body></html>"
        chunk_html_path = staging / f"_chunk_{i}.html"
        chunk_pdf_path = staging / f"_chunk_{i}.pdf"
        chunk_html_path.write_text(chunk_html, encoding="utf-8")
        if not playwright_pdf(str(chunk_html_path), str(chunk_pdf_path)):
            sys.stderr.write(f"[chunked_pdf] chunk {i} failed\n")
            return False
        chunk_pdfs.append(str(chunk_pdf_path))
        log(f"  chunk {i + 1}/{total} done")
        # Cleanup chunk html
        try:
            chunk_html_path.unlink()
        except Exception:
            pass

    writer = PdfWriter()
    for p in chunk_pdfs:
        r = PdfReader(p)
        for pg in r.pages:
            writer.add_page(pg)
    with open(output_path, "wb") as f:
        writer.write(f)
    # Cleanup chunk PDFs
    for p in chunk_pdfs:
        try:
            Path(p).unlink()
        except Exception:
            pass
    return True


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=f"SOP Portable {VERSION} — CLI PDF generator")
    ap.add_argument("--steps", required=True, help="path to edit/steps.xlsx")
    ap.add_argument("--content", default="", help="path to edit/content.xlsx (optional)")
    ap.add_argument("--images", required=True, help="path to images/ folder")
    ap.add_argument("--output", required=True, help="output PDF path (e.g. output/sop.pdf)")
    ap.add_argument("--template-dir", default="", help="app/ directory (template.html, styles.css, fonts/, logo.png)")
    args = ap.parse_args()

    script_dir = Path(__file__).resolve().parent
    template_dir = args.template_dir or str(script_dir)

    print(f"SOP Portable {VERSION} — پیشگامان پاکران الکترونیک")
    print()
    print(f"  Steps:    {args.steps}")
    print(f"  Content:  {args.content or '(none)'}")
    print(f"  Images:   {args.images}")
    print(f"  Output:   {args.output}")
    print(f"  Template: {template_dir}")
    print()

    if not Path(args.steps).exists():
        print(f"  ERROR: steps file not found: {args.steps}")
        return 1
    if not Path(template_dir).exists():
        print(f"  ERROR: template_dir not found: {template_dir}")
        return 1

    # Step 1: read Excel
    log("step 1/7 — reading DocumentInfo / Phases / Steps")
    doc = read_document_info(args.steps)
    phases = read_phases(args.steps)
    steps = read_steps(args.steps, phases)
    if not steps:
        print("  ERROR: no steps found in workbook")
        return 1
    log(f"  document: {doc.title}")
    log(f"  phases:   {len(phases)}")
    log(f"  steps:    {len(steps)}")
    for p in phases:
        log(f"    {p['letter']}: {p['short_name']} — {p.get('steps_count', 0)} مرحله")

    # Step 2: read content.xlsx (optional)
    if args.content and Path(args.content).exists():
        log("step 2/7 — reading content.xlsx (informational)")
    else:
        log("step 2/7 — content.xlsx skipped (not provided)")

    # Step 3: prepare output dir + copy logo + fonts
    log("step 3/7 — preparing output dir (logo, fonts)")
    out_dir = prepare_output(args.output, template_dir)

    # Step 4: process images
    log("step 4/7 — scanning + processing images")
    collect_images(steps, args.images, str(out_dir))
    total_imgs = sum(len(s.image_slots) for s in steps)
    log(f"  processed: {total_imgs} images for {sum(1 for s in steps if s.image_slots)} steps")

    # Step 5: render HTML
    log("step 5/7 — rendering HTML with Jinja2")
    html_path = render_html(out_dir, template_dir, steps, doc, phases)
    log(f"  wrote: {html_path}")

    # Step 6: render PDF with Playwright
    log("step 6/7 — rendering PDF with Playwright")
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    ok = playwright_pdf(html_path, args.output)
    if not ok:
        log("  Playwright failed → chunked fallback")
        ok = chunked_pdf(html_path, args.output)
    if not ok:
        print("  ERROR: PDF generation failed (both Playwright and chunked fallback)")
        return 1

    # Step 7: done
    log(f"step 7/7 — PDF written: {args.output}")
    print()
    print(f"  ✓ PDF created: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
