#!/usr/bin/env python3
"""Deterministic DOCX/PDF builder for the JobFitSkill.

Usage:
    python3 build_resume.py --spec resume_spec.json --out out.docx --pdf
    python3 build_resume.py --spec resume_spec.json --out out.docx --pdf --max-pages 1

The spec is a JSON object. See SKILL.md §17.6. The script only handles layout;
all content decisions belong to the agent.
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

PAGE_SIZES_CM = {"A4": (21.0, 29.7), "LETTER": (21.59, 27.94)}

DEFAULTS = {
    "cjk_font": "Noto Sans CJK SC",
    "latin_font": "Liberation Serif",
    "body_size": 10.5,
    "date_size": 10,
    "heading_size": 11,
    "name_size": 19,
    "contact_size": 9.5,
    "line_spacing": 1.15,
    "space_after": 2,
    "heading_space_before": 8,
    "heading_space_after": 3,
    "bullet_indent_cm": 0.42,
    "bullet_char": "·",
    "align": "left",
    "accent_color": None,
}


def _rgb(value):
    """Return an RGB tuple from [r, g, b] / '#rrggbb' / 'rrggbb'."""
    if not value:
        return None
    if isinstance(value, (list, tuple)) and len(value) == 3:
        return tuple(int(v) for v in value)
    if isinstance(value, str):
        s = value.lstrip("#")
        if len(s) == 6:
            return tuple(int(s[i : i + 2], 16) for i in (0, 2, 4))
    return None


def _hex(value):
    rgb = _rgb(value)
    return "{:02X}{:02X}{:02X}".format(*rgb) if rgb else "808080"


def set_run_font(run, cjk, latin, size=None, bold=None, italic=None, color=None):
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), cjk)
    rfonts.set(qn("w:cs"), latin)


BOLD_RE = re.compile(r"(\*\*[^*]+\*\*)")


def add_rich_text(paragraph, text, *, cjk, latin, size=None, bold=None, italic=None, color=None):
    """Add text to paragraph, rendering **bold** segments."""
    for part in BOLD_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, cjk, latin, size=size, bold=True, italic=italic, color=color)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, cjk, latin, size=size, bold=bold, italic=italic, color=color)


def set_bottom_border(paragraph, sz_eighth_pt=6, color="808080"):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz_eighth_pt))
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    ppr.append(pbdr)


def add_tab_stop(paragraph, position_cm):
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Cm(position_cm), WD_TAB_ALIGNMENT.RIGHT
    )


def build(spec, out_path):
    style_cfg = {**DEFAULTS, **spec.get("style", {})}
    page_cfg = spec.get("page", {})
    margins = page_cfg.get("margins_cm", {})
    top = float(margins.get("top", 1.9))
    bottom = float(margins.get("bottom", 1.7))
    left = float(margins.get("left", 2.0))
    right = float(margins.get("right", 2.0))

    doc = Document()

    # Page setup
    section = doc.sections[0]
    size_name = str(page_cfg.get("size", "A4")).upper()
    width_cm, height_cm = PAGE_SIZES_CM.get(size_name, PAGE_SIZES_CM["A4"])
    section.page_width = Cm(width_cm)
    section.page_height = Cm(height_cm)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    content_width_cm = width_cm - left - right

    # Base style
    normal = doc.styles["Normal"]
    normal.font.name = style_cfg["latin_font"]
    normal.font.size = Pt(style_cfg["body_size"])
    normal.font.color.rgb = RGBColor(0, 0, 0)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), style_cfg["latin_font"])
    rfonts.set(qn("w:hAnsi"), style_cfg["latin_font"])
    rfonts.set(qn("w:eastAsia"), style_cfg["cjk_font"])
    rfonts.set(qn("w:cs"), style_cfg["latin_font"])

    base_fmt = normal.paragraph_format
    base_fmt.line_spacing = style_cfg["line_spacing"]
    base_fmt.space_before = Pt(0)
    base_fmt.space_after = Pt(style_cfg["space_after"])
    base_fmt.widow_control = True

    common = {
        "cjk": style_cfg["cjk_font"],
        "latin": style_cfg["latin_font"],
    }
    accent = _rgb(style_cfg.get("accent_color"))

    # Header
    header = spec.get("header") or {}
    photo = header.get("photo") or {}
    if photo.get("path") and Path(photo["path"]).exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1
        p.add_run().add_picture(
            photo["path"], width=Cm(float(photo.get("width_cm", 2.5)))
        )
    elif photo.get("path"):
        print(f"[WARN] photo not found: {photo['path']}", file=sys.stderr)
    if header.get("name"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        add_rich_text(
            p, header["name"], size=style_cfg["name_size"], bold=True, **common
        )
    if header.get("contact"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(style_cfg["space_after"] + 2)
        add_rich_text(p, header["contact"], size=style_cfg["contact_size"], **common)
    if header.get("divider"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        set_bottom_border(p, sz_eighth_pt=8, color="404040")

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }

    for block in spec.get("blocks", []):
        btype = block.get("type", "para")

        if btype == "heading":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(block.get("space_before", style_cfg["heading_space_before"]))
            pf.space_after = Pt(block.get("space_after", style_cfg["heading_space_after"]))
            pf.keep_with_next = True
            pf.keep_together = True
            add_rich_text(
                p,
                block.get("text", ""),
                size=style_cfg["heading_size"],
                bold=True,
                color=accent,
                **common,
            )
            if block.get("border", True):
                set_bottom_border(p, sz_eighth_pt=6, color=_hex(accent) if accent else "808080")

        elif btype == "entry":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(block.get("space_before", 0))
            pf.space_after = Pt(block.get("space_after", style_cfg["space_after"]))
            pf.keep_together = True
            add_tab_stop(p, content_width_cm)
            left_text = block.get("left", "")
            right_text = block.get("right", "")
            bold_left = block.get("bold_left", True)
            add_rich_text(
                p,
                left_text,
                size=style_cfg["body_size"],
                bold=bold_left,
                **common,
            )
            if right_text:
                p.add_run("\t")
                run = p.add_run(right_text)
                set_run_font(run, size=style_cfg["date_size"], bold=False, **common)

        elif btype == "bullet":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(style_cfg["bullet_indent_cm"])
            pf.first_line_indent = Cm(-style_cfg["bullet_indent_cm"])
            pf.space_before = Pt(0)
            pf.space_after = Pt(block.get("space_after", style_cfg["space_after"]))
            # Long bullets must be allowed to split; keep_together pushes them to next page.
            pf.keep_together = False
            p.alignment = align_map.get(block.get("align", style_cfg["align"]), WD_ALIGN_PARAGRAPH.LEFT)
            bullet_char = block.get("bullet", style_cfg["bullet_char"])
            run = p.add_run(f"{bullet_char} ")
            set_run_font(run, size=style_cfg["body_size"], **common)
            add_rich_text(
                p,
                block.get("text", ""),
                size=style_cfg["body_size"],
                **common,
            )

        elif btype == "kv":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(block.get("space_after", style_cfg["space_after"]))
            pf.keep_together = True
            label = block.get("label", "")
            if label:
                run = p.add_run(f"{label}：")
                set_run_font(
                    run, size=style_cfg["body_size"], bold=True, **common
                )
            add_rich_text(
                p,
                block.get("text", ""),
                size=style_cfg["body_size"],
                **common,
            )

        elif btype == "space":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(block.get("pt", 4))
            p.paragraph_format.line_spacing = 1

        else:  # para
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(block.get("space_before", 0))
            pf.space_after = Pt(block.get("space_after", style_cfg["space_after"]))
            pf.keep_together = False
            p.alignment = align_map.get(block.get("align", style_cfg["align"]), WD_ALIGN_PARAGRAPH.LEFT)
            add_rich_text(
                p,
                block.get("text", ""),
                size=block.get("size", style_cfg["body_size"]),
                bold=block.get("bold"),
                **common,
            )

    core = doc.core_properties
    core.title = header.get("name", "") + " 简历"
    core.author = header.get("name", "")

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def to_pdf(docx_path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[WARN] LibreOffice not found; DOCX generated, PDF skipped.", file=sys.stderr)
        return None
    outdir = str(Path(docx_path).parent)
    env = dict(os.environ)
    env.setdefault("HOME", "/tmp")
    proc = subprocess.run(
        [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)],
        capture_output=True,
        text=True,
        env=env,
    )
    pdf_path = Path(docx_path).with_suffix(".pdf")
    if not pdf_path.exists():
        print(proc.stdout, proc.stderr, file=sys.stderr)
        print("[ERROR] PDF conversion failed.", file=sys.stderr)
        return None
    return pdf_path


def report(pdf_path, max_pages=None, expected_pages=None):
    try:
        from pypdf import PdfReader
    except ImportError:
        print("[INFO] pypdf not installed; page count skipped.")
        return
    reader = PdfReader(str(pdf_path))
    pages = len(reader.pages)
    last_text = reader.pages[-1].extract_text() or ""
    lines = [ln for ln in last_text.splitlines() if ln.strip()]
    print(f"[CHECK] pages={pages} last_page_lines={len(lines)}")
    if expected_pages and pages != expected_pages:
        print(
            f"[WARN] pages={pages} != original page count {expected_pages}; "
            "see SKILL.md §11.6.",
            file=sys.stderr,
        )
    if max_pages and pages > max_pages:
        print(f"[WARN] pages={pages} exceeds max_pages={max_pages}; trim content or tighten spacing.", file=sys.stderr)
    if pages > 1 and len(lines) <= 3:
        print("[WARN] last page has <=3 lines (orphan page).", file=sys.stderr)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True, help="path to JSON spec")
    ap.add_argument("--out", required=True, help="output .docx path")
    ap.add_argument("--pdf", action="store_true", help="also convert to PDF via LibreOffice")
    ap.add_argument("--max-pages", type=int, default=None)
    args = ap.parse_args()

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)

    docx_path = build(spec, args.out)
    print(f"[OK] docx: {docx_path}")

    if args.pdf:
        pdf_path = to_pdf(docx_path)
        if pdf_path:
            print(f"[OK] pdf: {pdf_path}")
            expected = spec.get("page", {}).get("pages")
            report(pdf_path, args.max_pages or expected, expected)


if __name__ == "__main__":
    main()
