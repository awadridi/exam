import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Mm
import io
import os
from datetime import datetime
from copy import deepcopy

# =====================================
# 1. نظام تسجيل الدخول باستخدام Secrets
# =====================================
def login():
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
        st.session_state['username'] = ""

    if not st.session_state['logged_in']:
        st.markdown("<h2 style='text-align: center;'>🔐 نظام تكليفات المكتب - دخول</h2>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("login_form"):
                user = st.text_input("اسم المستخدم").lower().strip()
                pw = st.text_input("كلمة المرور", type="password").strip()
                submit = st.form_submit_button("دخول")
                
                if submit:
                    try:
                        valid_password = st.secrets[f"password_{user}"]
                        if pw == valid_password:
                            st.session_state['logged_in'] = True
                            st.session_state['username'] = user
                            st.rerun()
                        else:
                            st.error("❌ كلمة المرور غير صحيحة")
                    except KeyError:
                        st.error("❌ اسم المستخدم غير معرف في Secrets")
        return False
    return True

if not login():
    st.stop()

# =====================================
# 2. إعدادات الواجهة وقاعدة البيانات
# =====================================
st.set_page_config(page_title="نظام تكليف المراقبة", layout="wide")

# عرض اسم المستخدم وزر تسجيل الخروج في الشريط الجانبي
with st.sidebar:
    st.markdown(f"### 👤 المستخدم: {st.session_state.username}")
    if st.button("🚪 تسجيل الخروج"):
        st.session_state['logged_in'] = False
        st.session_state['username'] = ""
        st.rerun()
    st.markdown("---")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    [data-testid="stSidebar"] { direction: rtl; text-align: right; }
    .main-info-box {
        background-color: #1e2129;
        padding: 15px;
        border-radius: 8px;
        border-right: 5px solid #00ffcc;
        margin-bottom: 15px;
    }
    .data-label { color: #888; font-size: 0.9rem; }
    .data-value { color: #fff; font-weight: bold; margin-left: 15px; }
    button[key^="save_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .right-align { text-align: right; direction: rtl; width: 100%; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect("data_system_v26.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
conn.commit()

# =====================================
# 3. وظائف معالجة الملفات
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9: phone_val = '0' + phone_val
    
    repls = {
        'ZNAME': str(row.get('name', '')), 
        'ZID': str(row.get('id', '')), 
        'ZJOB': str(row.get('role', '') or 'مراقب'), 
        'ZHALL': str(h_name or ''), 
        'ZLOC': str(h_city or ''), 
        'ZWORK': str(row.get('school', '')), 
        'ZCITY': str(row.get('city', '')),
        'ZPHONE': phone_val
    }
    
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)
    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in repls.items():
                        if k in p.text:
                            for run in p.runs:
                                if k in run.text:
                                    run.text = run.text.replace(k, v)
    return doc_obj

def generate_bulk_word(df, h_name):
    if not os.path.exists("template.docx") or df.empty: return None
    final_doc = Document("template.docx")
    final_doc._body.clear_content()
    for idx, row in df.iterrows():
        temp_doc = Document("template.docx")
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        for element in temp_doc.element.body:
            if element.tag.endswith('sectPr'): continue
            final_doc.element.body.append(deepcopy(element))
        if idx < len(df) - 1:
            final_doc.add_page_break()
    out = io.BytesIO()
    final_doc.save(out); out.seek(0)
    return out

def generate_single_doc(row):
    if
