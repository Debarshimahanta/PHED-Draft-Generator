from __future__ import annotations

import html
import re
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .config import ASSETS, OUTPUTS, OFFICE_HEADER, SIGNATORY_LINES, COPY_CATEGORIES


def _set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ['val', 'sz', 'space', 'color']:
                if key in kwargs[edge]:
                    element.set(qn('w:{}'.format(key)), str(kwargs[edge][key]))


def _font(run, *, bold=False, size=12, name='Times New Roman', italic=False):
    run.bold = bold
    run.italic = italic
    run.font.name = name
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)


def _para_format(p, before=0, after=0, line=1.0, align=None):
    fmt = p.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    if align is not None:
        p.alignment = align


def _copy_sentence(item: dict) -> str:
    designation = (item.get('designation') or '').strip()
    category = item.get('category') or 'Custom'
    extra = (item.get('extra') or '').strip()
    ps_target = (item.get('ps_target') or '').strip()
    custom = (item.get('custom') or '').strip()

    if not designation:
        return ''

    if category == 'PS Intimation':
        base = f"The {designation} for favour of kind intimation of the matter to {ps_target}." if ps_target else f"The {designation} for favour of kind intimation."
    elif category == 'Custom':
        base = custom if custom else f"The {designation}"
    else:
        phrase = COPY_CATEGORIES.get(category, '') or ''
        base = f"The {designation} {phrase}".strip()

    if extra:
        if base and not base.endswith(('.', '!', '?')):
            base += '.'
        base += ' ' + extra
    return base.strip()


def build_docx(data: dict, output_path: Path | None = None) -> Path:
    OUTPUTS.mkdir(exist_ok=True)
    output_path = output_path or OUTPUTS / 'generated_letter.docx'

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.25)
    section.right_margin = Cm(1.25)

    p = doc.add_paragraph()
    _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    emblem = ASSETS / 'assam_emblem.jpeg'
    if emblem.exists():
        r = p.add_run()
        r.add_picture(str(emblem), width=Cm(1.7))

    for idx, line in enumerate(OFFICE_HEADER):
        p = doc.add_paragraph()
        _para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(line)
        _font(r, bold=True, size=12 if idx < 3 else 11)

    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=2)
    table.autofit = True
    for cell in table.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        _set_cell_border(cell, top={'val':'nil'}, left={'val':'nil'}, bottom={'val':'nil'}, right={'val':'nil'})

    p = table.cell(0,0).paragraphs[0]
    r = p.add_run('No.\n' + (data.get('file_no') or ''))
    _font(r, bold=True, size=11)

    p = table.cell(0,1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    letter_date = (data.get('date') or '').strip()
    r = p.add_run(f'Dated: {letter_date}' if letter_date else '')
    _font(r, size=11)

    doc.add_paragraph()

    p = doc.add_paragraph()
    _para_format(p, after=0)
    r = p.add_run('To,')
    _font(r, bold=True, size=12)

    for line in (data.get('addressee') or '').splitlines():
        p = doc.add_paragraph()
        _para_format(p, after=0)
        r = p.add_run(line)
        _font(r, size=12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    _para_format(p, after=8)
    r = p.add_run('Sub: ')
    _font(r, bold=True, size=12)
    r = p.add_run(data.get('subject') or '')
    _font(r, size=12)

    p = doc.add_paragraph()
    _para_format(p, after=8)
    r = p.add_run('Sir,')
    _font(r, size=12)

    body = data.get('body') or ''
    for block in re.split(r'\n\s*\n', body):
        p = doc.add_paragraph()
        _para_format(p, after=7, line=1.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        r = p.add_run(block)
        _font(r, size=12)

    doc.add_paragraph()
    for line in SIGNATORY_LINES:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _para_format(p, after=0)
        r = p.add_run(line)
        _font(r, bold=True, size=12)

    copies = [x for x in data.get('copies', []) if (x.get('designation') or '').strip()]
    if copies:
        doc.add_paragraph()
        p = doc.add_paragraph()
        _para_format(p, after=5)
        r = p.add_run('Copy to:')
        _font(r, bold=True, size=12)

        for i, item in enumerate(copies, start=1):
            text = _copy_sentence(item)
            p = doc.add_paragraph(style=None)
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.first_line_indent = Cm(-0.45)
            _para_format(p, after=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            r = p.add_run(f'{i}.  {text}')
            _font(r, size=11.5)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run('e-signed')
    _font(r, italic=True, size=11)

    stamp = ASSETS / 'chief_engineer_stamp.png'
    if stamp.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run()
        r.add_picture(str(stamp), width=Cm(6.6))
    else:
        for line in SIGNATORY_LINES:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _para_format(p, after=0)
            r = p.add_run(line)
            _font(r, bold=True, size=12)

    doc.save(output_path)
    return output_path


def build_pdf(docx_path: Path) -> Path | None:
    out_dir = docx_path.parent
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', str(out_dir), str(docx_path)],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
        )
        pdf = out_dir / (docx_path.stem + '.pdf')
        return pdf if pdf.exists() else None
    except Exception:
        return None


def preview_html(data: dict) -> str:
    def esc(value):
        return html.escape(value or '').replace('\n', '<br>')

    copies = []
    for i, item in enumerate(data.get('copies', []), start=1):
        if (item.get('designation') or '').strip():
            copies.append(f'<div class="copy"><span>{i}.</span><div>{html.escape(_copy_sentence(item))}</div></div>')

    copy_html = ''.join(copies)
    body = data.get('body') or ''
    body_html = ''.join(
        f'<p>{html.escape(block).replace(chr(10), "<br>")}</p>'
        for block in re.split(r'\n\s*\n', body)
    )

    return f'''<!doctype html><html><head><meta charset="utf-8"><style>
    body{{font-family:"Times New Roman",serif;background:#efefef;margin:0;padding:24px;color:#111}}
    .page{{background:#fff;max-width:780px;min-height:1020px;margin:auto;padding:38px 44px;box-shadow:0 2px 12px #aaa}}
    .center{{text-align:center;font-weight:700;line-height:1.15}} .header{{font-size:16px}}
    .emblem{{text-align:center;min-height:10px}} .emblem img{{width:58px}}
    .meta{{display:flex;justify-content:space-between;margin-top:26px;font-size:15px}}
    .to{{margin-top:24px;line-height:1.25}} .subject{{margin-top:22px}}
    .body p{{text-align:justify;text-indent:42px;line-height:1.25;margin:12px 0}}
    .sign{{text-align:right;font-weight:700;margin:34px 38px 0 0;line-height:1.25}}
    .copyhead{{font-weight:700;margin-top:26px}}
    .copy{{display:grid;grid-template-columns:28px 1fr;gap:4px;margin:6px 18px;line-height:1.25;text-align:justify}}
    .esigned{{text-align:right;font-style:italic;margin:40px 55px 0 0}}
    </style></head><body><div class="page">
    <div class="emblem"><img src="data:image/jpeg;base64,{{EMBLEM}}"></div>
    <div class="center header">GOVERNMENT OF ASSAM<br>OFFICE OF THE CHIEF ENGINEER (PHE) WATER, ASSAM<br>HENGRABARI, GUWAHATI - 36<br>Email - asphe@rediffmail.com</div>
    <div class="meta"><div><b>No.</b><br><b>{esc(data.get('file_no'))}</b></div><div>{'Dated: ' + esc(data.get('date')) if data.get('date') else ''}</div></div>
    <div class="to"><b>To,</b><br>{esc(data.get('addressee'))}</div>
    <div class="subject"><b>Sub:</b> {esc(data.get('subject'))}</div>
    <div style="margin-top:18px">Sir,</div><div class="body">{body_html}</div>
    <div class="sign">Chief Engineer (PHE) Water,<br>Hengrabari, Assam</div>
    {'<div class="copyhead">Copy to:</div>' + copy_html if copy_html else ''}
    <div class="esigned">e-signed</div>
    <div class="sign">Chief Engineer (PHE) Water,<br>Hengrabari, Assam</div>
    </div></body></html>'''
