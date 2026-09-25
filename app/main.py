from __future__ import annotations

import base64
from datetime import date
from pathlib import Path
import sys

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.config import ASSETS, TEMPLATE_NAME
from app.generator import build_docx, build_pdf, preview_html

st.set_page_config(page_title='PHED Letter Draft Generator', page_icon='📝', layout='wide')

st.markdown("""
<style>
.block-container{padding-top:0.8rem;padding-bottom:1.5rem;max-width:1600px}
.app-title{background:linear-gradient(90deg,#12579c,#2479ba);color:white;padding:18px 24px;border-radius:10px;margin-bottom:14px}
.app-title h1{font-size:28px;margin:0}.app-title p{margin:2px 0 0;font-size:14px;opacity:.95}
.section{font-weight:700;font-size:17px;margin:10px 0 6px}.copy-card{border:1px solid #d9e2ec;border-radius:10px;padding:10px;margin:8px 0;background:#fafcff}
.stDownloadButton>button,.stButton>button{border-radius:8px}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="app-title"><h1>PUBLIC HEALTH ENGINEERING DEPARTMENT — Letter Draft Generator</h1><p>CE_1 · Exact user text · Controlled official formatting</p></div>', unsafe_allow_html=True)

if 'copies' not in st.session_state:
    st.session_state.copies = [
        {'designation':'Addl. Chief Engineer (PHE), LAZ','category':'Info & Action','extra':'','ps_target':'','custom':''},
        {'designation':'Superintending Engineer (PHE), Guwahati Circle','category':'Info & Action','extra':'','ps_target':'','custom':''},
    ]

left, right = st.columns([1.05, 1], gap='large')

with left:
    st.selectbox('Template', [TEMPLATE_NAME], index=0, disabled=True)
    c1, c2 = st.columns([1,1])
    file_no = c1.text_input('File No.', value='')
    dt = c2.date_input('Date', value=date.today())
    addressee = st.text_area('To (Addressee)', height=92, placeholder='The Executive Engineer (PHE)\nGuwahati Division I')
    subject = st.text_area('Subject', height=92)
    body = st.text_area('Body — reproduced exactly as entered', height=300)

    st.markdown('<div class="section">Copy To</div>', unsafe_allow_html=True)
    st.caption('Add as many copies as required. The standard phrase is generated from the selected category; optional instruction remains exactly as typed.')

    delete_idx = None
    for i, item in enumerate(st.session_state.copies):
        with st.container(border=True):
            a,b = st.columns([2.2,1.1])
            item['designation'] = a.text_input('Designation / Address', value=item.get('designation',''), key=f'des_{i}')
            item['category'] = b.selectbox('Category', ['Info & Action','Kind Information','PS Intimation','Custom'], index=['Info & Action','Kind Information','PS Intimation','Custom'].index(item.get('category','Info & Action')), key=f'cat_{i}')

            if item['category'] == 'PS Intimation':
                item['ps_target'] = st.text_input('Intimation to designation', value=item.get('ps_target',''), key=f'ps_{i}', placeholder="Special Chief Secretary, PHED")
            elif item['category'] == 'Custom':
                item['custom'] = st.text_area('Exact copy text', value=item.get('custom',''), key=f'custom_{i}', height=70)

            item['extra'] = st.text_input('Additional instruction (optional)', value=item.get('extra',''), key=f'extra_{i}', placeholder='He is also requested to attend the aforesaid review meeting.')

            if st.button('Remove copy', key=f'rm_{i}'):
                delete_idx = i

    if delete_idx is not None:
        st.session_state.copies.pop(delete_idx)
        st.rerun()

    if st.button('＋ Add Copy', use_container_width=True):
        st.session_state.copies.append({'designation':'','category':'Info & Action','extra':'','ps_target':'','custom':''})
        st.rerun()

    data = {
        'file_no': file_no,
        'date': dt.strftime('%d/%m/%Y') if dt else '',
        'addressee': addressee,
        'subject': subject,
        'body': body,
        'copies': st.session_state.copies,
    }

    b1,b2,b3 = st.columns(3)

    if b1.button('Generate DOCX', type='primary', use_container_width=True):
        path = build_docx(data, ROOT/'outputs'/'PHED_Letter.docx')
        st.session_state.docx_path = str(path)
        st.success('DOCX generated.')

    if b2.button('Generate PDF', use_container_width=True):
        path = build_docx(data, ROOT/'outputs'/'PHED_Letter.docx')
        pdf = build_pdf(path)
        if pdf:
            st.session_state.pdf_path = str(pdf)
            st.success('PDF generated.')
        else:
            st.error('PDF conversion is unavailable on this system. DOCX was generated successfully.')

    if b3.button('Clear All', use_container_width=True):
        for k in list(st.session_state.keys()):
            if k != 'copies':
                del st.session_state[k]
        st.session_state.copies = []
        st.rerun()

    if st.session_state.get('docx_path') and Path(st.session_state.docx_path).exists():
        with open(st.session_state.docx_path,'rb') as f:
            st.download_button('Download DOCX', f.read(), file_name='PHED_Letter.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', use_container_width=True)

    if st.session_state.get('pdf_path') and Path(st.session_state.pdf_path).exists():
        with open(st.session_state.pdf_path,'rb') as f:
            st.download_button('Download PDF', f.read(), file_name='PHED_Letter.pdf', mime='application/pdf', use_container_width=True)

with right:
    st.markdown('### Live Preview')
    html_preview = preview_html(data)

    emblem = ASSETS / 'assam_emblem.jpeg'
    if emblem.exists():
        encoded = base64.b64encode(emblem.read_bytes()).decode('ascii')
        html_preview = html_preview.replace('{EMBLEM}', encoded)
    else:
        html_preview = html_preview.replace(
            '<div class="emblem"><img src="data:image/jpeg;base64,{EMBLEM}"></div>',
            '<div class="emblem"></div>'
        )

    stamp = ASSETS / 'chief_engineer_stamp.png'
    if stamp.exists():
        stamp_encoded = base64.b64encode(stamp.read_bytes()).decode('ascii')
        html_preview = html_preview.replace('{STAMP}', stamp_encoded)
    else:
        html_preview = html_preview.replace(
            '<div class="stamp"><img src="data:image/png;base64,{STAMP}"></div>',
            '<div class="stamp"><b>Chief Engineer (PHE) Water,<br>Hengrabari, Assam</b></div>'
        )

    st.components.v1.html(html_preview, height=1120, scrolling=True)
