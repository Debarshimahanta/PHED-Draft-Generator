from __future__ import annotations

import html
import re
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .config import ASSETS, OUTPUTS, OFFICE_HEADER, COPY_CATEGORIES


A4_WIDTH = Cm(21.0)
A4_HEIGHT = Cm(29.7)


def _set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge not in kwargs:
            continue
        tag = f"w:{edge}"
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        for key in ("val", "sz", "space", "color"):
            if key in kwargs[edge]:
                element.set(qn(f"w:{key}"), str(kwargs[edge][key]))


def _set_paragraph_bottom_border(paragraph, size=8, color="000000", space=4):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)

    bottom = pBdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pBdr.append(bottom)

    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)


def _font(run, *, bold=False, size=12, name="Times New Roman", italic=False):
    run.bold = bold
    run.italic = italic
    run.font.name = name
    run.font.size = Pt(size)
    if run._element.rPr is None:
        run._element.get_or_add_rPr()
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def _para_format(
    p,
    *,
    before=0,
    after=0,
    line=1.0,
    align=None,
    first_line_cm=None,
    left_cm=None,
    right_cm=None,
):
    fmt = p.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    if align is not None:
        p.alignment = align
    if first_line_cm is not None:
        fmt.first_line_indent = Cm(first_line_cm)
    if left_cm is not None:
        fmt.left_indent = Cm(left_cm)
    if right_cm is not None:
        fmt.right_indent = Cm(right_cm)


def _copy_sentence(item: dict) -> str:
    designation = (item.get("designation") or "").strip()
    category = item.get("category") or "Custom"
    extra = (item.get("extra") or "").strip()
    ps_target = (item.get("ps_target") or "").strip()
    custom = (item.get("custom") or "").strip()

    if not designation:
        return ""

    if category == "PS Intimation":
        if ps_target:
            base = f"The {designation} for favour of kind intimation of the matter to {ps_target}."
        else:
            base = f"The {designation} for favour of kind intimation."
    elif category == "Custom":
        base = custom if custom else f"The {designation}"
    else:
        phrase = COPY_CATEGORIES.get(category, "") or ""
        base = f"The {designation} {phrase}".strip()

    if extra:
        if base and not base.endswith((".", "!", "?")):
            base += "."
        base += " " + extra

    return base.strip()


def _add_header(doc: Document):
    emblem = ASSETS / "assam_emblem.jpeg"

    if emblem.exists():
        p = doc.add_paragraph()
        _para_format(p, after=1, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run()
        r.add_picture(str(emblem), width=Cm(1.45))

    for idx, line in enumerate(OFFICE_HEADER):
        p = doc.add_paragraph()
        _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(line)
        if idx == 0:
            _font(r, bold=True, size=11.5)
        elif idx == 1:
            _font(r, bold=True, size=11.5)
        elif idx == 2:
            _font(r, bold=True, size=11)
        else:
            _font(r, bold=True, size=10.5)

    rule = doc.add_paragraph()
    _para_format(rule, before=3, after=8)
    _set_paragraph_bottom_border(rule, size=10, space=2)


def _add_final_signatory(doc: Document):
    # Only one signatory/seal block is used. If Copy To exists it appears after Copy To;
    # otherwise it appears directly after the body.
    p = doc.add_paragraph()
    _para_format(p, before=12, after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, right_cm=0.4)
    r = p.add_run("e-signed")
    _font(r, italic=True, size=10.5)

    stamp = ASSETS / "chief_engineer_stamp.png"
    if stamp.exists():
        p = doc.add_paragraph()
        _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.RIGHT, right_cm=0.1)
        r = p.add_run()
        # Width-only sizing preserves the original aspect ratio and prevents distortion.
        r.add_picture(str(stamp), width=Cm(5.25))
    else:
        p = doc.add_paragraph()
        _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.RIGHT, right_cm=0.4)
        r = p.add_run("Chief Engineer (PHE) Water,")
        _font(r, bold=True, size=11.5)
        p = doc.add_paragraph()
        _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.RIGHT, right_cm=0.4)
        r = p.add_run("Hengrabari, Assam")
        _font(r, bold=True, size=11.5)


def build_docx(data: dict, output_path: Path | None = None) -> Path:
    OUTPUTS.mkdir(exist_ok=True)
    output_path = output_path or OUTPUTS / "generated_letter.docx"

    doc = Document()
    section = doc.sections[0]
    section.page_width = A4_WIDTH
    section.page_height = A4_HEIGHT
    section.top_margin = Cm(1.15)
    section.bottom_margin = Cm(1.25)
    section.left_margin = Cm(1.65)
    section.right_margin = Cm(1.65)
    section.header_distance = Cm(0.4)
    section.footer_distance = Cm(0.5)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    _add_header(doc)

    meta = doc.add_table(rows=1, cols=2)
    meta.autofit = False
    meta.columns[0].width = Cm(9.0)
    meta.columns[1].width = Cm(8.2)

    for cell in meta.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        _set_cell_border(
            cell,
            top={"val": "nil"},
            left={"val": "nil"},
            bottom={"val": "nil"},
            right={"val": "nil"},
        )

    p = meta.cell(0, 0).paragraphs[0]
    _para_format(p, after=0)
    r = p.add_run("No. ")
    _font(r, bold=True, size=11.5)
    r = p.add_run(data.get("file_no") or "")
    _font(r, bold=True, size=11.5)

    p = meta.cell(0, 1).paragraphs[0]
    _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    letter_date = (data.get("date") or "").strip()
    r = p.add_run(f"Dated: {letter_date}" if letter_date else "")
    _font(r, size=11)

    doc.add_paragraph().paragraph_format.space_after = Pt(1)

    p = doc.add_paragraph()
    _para_format(p, after=1)
    r = p.add_run("To,")
    _font(r, bold=True, size=11.5)

    for line in (data.get("addressee") or "").splitlines():
        p = doc.add_paragraph()
        _para_format(p, after=0)
        r = p.add_run(line)
        _font(r, size=11.5)

    p = doc.add_paragraph()
    _para_format(p, before=8, after=8)
    r = p.add_run("Sub: ")
    _font(r, bold=True, size=11.5)
    r = p.add_run(data.get("subject") or "")
    _font(r, size=11.5)

    p = doc.add_paragraph()
    _para_format(p, after=7)
    r = p.add_run("Sir,")
    _font(r, size=11.5)

    body = data.get("body") or ""
    blocks = re.split(r"\n\s*\n", body)

    for block in blocks:
        if block == "" and len(blocks) == 1:
            continue
        p = doc.add_paragraph()
        _para_format(
            p,
            after=7,
            line=1.05,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_cm=0.85,
        )
        r = p.add_run(block)
        _font(r, size=11.5)

    copies = [x for x in data.get("copies", []) if (x.get("designation") or "").strip()]

    if copies:
        p = doc.add_paragraph()
        _para_format(p, before=10, after=5)
        r = p.add_run("Copy to:")
        _font(r, bold=True, size=11.5)

        for i, item in enumerate(copies, start=1):
            text = _copy_sentence(item)
            p = doc.add_paragraph()
            _para_format(
                p,
                after=3,
                line=1.0,
                align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                left_cm=0.75,
            )
            p.paragraph_format.first_line_indent = Cm(-0.48)
            r = p.add_run(f"{i}.  {text}")
            _font(r, size=10.8)

    _add_final_signatory(doc)

    doc.save(output_path)
    return output_path


def build_pdf(docx_path: Path) -> Path | None:
    out_dir = docx_path.parent

    commands = [
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
    ]

    for cmd in commands:
        try:
            subprocess.run(
                cmd,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60,
            )
            pdf = out_dir / f"{docx_path.stem}.pdf"
            if pdf.exists():
                return pdf
        except Exception:
            continue

    return None


def preview_html(data: dict) -> str:
    def esc(value):
        return html.escape(value or "").replace("\n", "<br>")

    body = data.get("body") or ""
    body_html = "".join(
        f'<p>{html.escape(block).replace(chr(10), "<br>")}</p>'
        for block in re.split(r"\n\s*\n", body)
        if block != ""
    )

    copies = []
    for i, item in enumerate(data.get("copies", []), start=1):
        if (item.get("designation") or "").strip():
            copies.append(
                f'<div class="copy"><span>{i}.</span><div>{html.escape(_copy_sentence(item))}</div></div>'
            )

    copy_html = "".join(copies)

    return f'''<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
*{{box-sizing:border-box}}
body{{font-family:"Times New Roman",serif;background:#eceff2;margin:0;padding:22px;color:#111}}
.page{{
    background:#fff;
    width:794px;
    min-height:1123px;
    margin:auto;
    padding:44px 62px 48px;
    box-shadow:0 3px 16px rgba(0,0,0,.18);
}}
.emblem{{text-align:center;line-height:1}}
.emblem img{{width:54px;height:auto}}
.header{{text-align:center;font-weight:700;line-height:1.12;font-size:15px}}
.header .small{{font-size:13px}}
.rule{{border-top:1.5px solid #111;margin:8px 0 14px}}
.meta{{display:flex;justify-content:space-between;font-size:14px;margin-top:2px}}
.to{{font-size:14px;margin-top:17px;line-height:1.22}}
.subject{{font-size:14px;margin-top:20px;line-height:1.28}}
.salutation{{font-size:14px;margin-top:18px}}
.body p{{
    font-size:14px;
    line-height:1.30;
    text-align:justify;
    text-indent:34px;
    margin:9px 0;
}}
.copyhead{{font-size:14px;font-weight:700;margin-top:22px;margin-bottom:6px}}
.copy{{display:grid;grid-template-columns:24px 1fr;gap:5px;margin:4px 10px;font-size:13.4px;line-height:1.25;text-align:justify}}
.sign-wrap{{margin-top:30px;text-align:right;padding-right:6px}}
.esigned{{font-size:13px;font-style:italic;margin-right:25px;margin-bottom:3px}}
.stamp img{{width:198px;height:auto;object-fit:contain}}
</style>
</head>
<body>
<div class="page">
    <div class="emblem"><img src="data:image/jpeg;base64,{{EMBLEM}}"></div>
    <div class="header">
        GOVERNMENT OF ASSAM<br>
        OFFICE OF THE CHIEF ENGINEER (PHE) WATER, ASSAM<br>
        HENGRABARI, GUWAHATI - 36<br>
        <span class="small">Email - asphe@rediffmail.com</span>
    </div>
    <div class="rule"></div>

    <div class="meta">
        <div><b>No. {esc(data.get("file_no"))}</b></div>
        <div>{'Dated: ' + esc(data.get("date")) if data.get("date") else ''}</div>
    </div>

    <div class="to"><b>To,</b><br>{esc(data.get("addressee"))}</div>
    <div class="subject"><b>Sub:</b> {esc(data.get("subject"))}</div>
    <div class="salutation">Sir,</div>
    <div class="body">{body_html}</div>

    {'<div class="copyhead">Copy to:</div>' + copy_html if copy_html else ''}

    <div class="sign-wrap">
        <div class="esigned">e-signed</div>
        <div class="stamp"><img src="data:image/png;base64,{{STAMP}}"></div>
    </div>
</div>
</body>
</html>'''