import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
import io
import os
import re

# =====================================
# 1. إعدادات الواجهة وتثبيت الألوان (CSS)
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    div[data-testid="stExpander"] summary p { color: #ffffff !important; font-weight: bold !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; width: 100% !important; }
    label, .stMarkdown p { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات
db_path = "data_system_v16.db"
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# روابط Google Sheets
TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 2. الوظائف المساعدة (Functions)
# =====================================

def apply_smart_bold_replace(paragraph, data_map):
    text = paragraph.text
    if any(key in text for key in data_map):
        original_size = Pt(14)
        if paragraph.runs and paragraph.runs[0].font.size:
            original_size = paragraph.runs[0].font.size
        for run in paragraph.runs: run.text = ""
        parts = re.split(r'(<[^>]+>)', text)
        for part in parts:
            run = paragraph.add_run()
            if part in data_map:
                run.text = str(data_map[part]) if data_map[part] else ""
                run.bold = True
            else:
                run.text = part
            run.font.size = original_size

def generate_from_template(row):
    try:
        doc = Document("template.docx")
        replacements = {
            '<NAME>': str(row.get('name', '')),
            '<ID>': str(row.get('id', row.get('id_number', ''))),
            '<JOB>': str(row.get('role', row.get('job', ''))),
            '<HALL_NAME>': str(row.get('hall', '')),
            '<HALL_LOCATION>': str(row.get('hall_city', '')),
            '<WORKPLACE>': str(row.get('school', '')),
            '<CITY>': str(row.get('city', ''))
        }
        for p in doc.paragraphs: apply_smart_bold_replace(p, replacements)
        for table in doc.tables:
            for r in table.rows:
                for cell in r.cells:
                    for p in cell.paragraphs: apply_smart_bold_replace(p, replacements)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except: return None

def generate_bulk_docs(df, h_name):
    template_p = "template.docx"
    if not os.path.exists(template_p): return None
    final_doc = Document(template_p)
    final_doc._body.clear_content()
    for idx, r in df.iterrows():
        temp_doc = Document(template_p)
        repls = {
            '<NAME>': str(r.get('name', '')),
            '<ID>': str(r.get('id', '')),
            '<JOB>': str(r.get('role', '')),
            '<HALL_NAME>': str(h_name),
            '<HALL_LOCATION>': str(r.get('hall_city', '')),
            '<WORKPLACE>': str(r.get('school', '')),
            '<CITY>': str(r.get('city', ''))
        }
        # استبدال بسيط للسرعة في الدمج الجماعي
        for p in temp_doc.paragraphs:
            for k, v in repls.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text: run.text = run.text.replace(k, v); run.bold = True
        for table in temp_doc.tables:
            for row_t in table.rows:
                for cell in row_t.cells:
                    for p in cell.paragraphs:
                        for k, v in repls.items():
                            if k in p.text:
                                for run in p.runs:
                                    if k in run.text: run.text = run.text.replace(k, v); run.bold = True
        if idx > 0: final_doc.add_page_break()
        for element in temp_doc.element.body:
            if not element.tag.endswith('sectPr'):
                final_doc.element.body.append(element)
    out = io.BytesIO()
    final_doc.save(out)
    out.seek(0)
    return out

def sync_data():
    try:
        df_t = pd.read_csv(TEACHERS_URL)
        df_t.columns = df_t.columns.str.strip().str.lower()
        if 'id_number' in df_t.columns and 'id' not in df_t.columns:
            df_t.rename(columns={'id_number': 'id'}, inplace=True)
        for col in ['hall', 'role', 'hall_city']:
            if col not in df_t.columns: df_t[col] = ""
        df_t.to_sql('teachers', conn, if_exists='replace', index=False)
        
        df_h = pd.read_csv(HALLS_URL)
        df_h.columns = df_h.columns.str.strip().str.lower()
        df_h.to_sql('halls', conn, if_exists='replace', index=False)
        st.success("✅ تم التحديث بنجاح!")
        st.rerun()
    except Exception as e: st.error(f"❌ خطأ: {e}")

# =====================================
# 3. بناء التبويبات (Tabs)
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

# --- تبويب البحث ---
with tab_search:
    st.subheader("إدارة الموظفين الفردية")
    try:
        df_halls = pd.read_sql("SELECT * FROM halls", conn)
        hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    except: hall_map = {}
    
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو الهوية")
    if q:
        df_t = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for i, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | {row['role'] or '-'} | {row['hall'] or '-'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox("القاعة", hall_list, index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox("الوظيفة", role_list, index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("💾 حفظ", key=f"btn_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, hall_map.get(sel_hall, ""), row['id']))
                        conn.commit(); st.success("تم!"); st.rerun()
                    if row['role'] or row['hall']:
                        if st.button("❌ إلغاء", key=f"del_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit(); st.rerun()

# --- تبويب رفع الملفات ---
with tab_upload:
    st.subheader("تزامن البيانات مع Google Sheets")
    st.info("سيتم سحب بيانات المعلمين والقاعات من الروابط المبرمجة مسبقاً.")
    if st.button("🔄 تحديث الأسماء والقاعات الآن"):
        sync_data()

# --- تبويب الإدارة ---
with tab_manage:
    st.subheader("إدارة القاعات والتحميل الجماعي")
    df_active_halls = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    
    if not df_active_halls.empty:
        hall_to_manage = st.selectbox("اختر قاعة للعرض:", [""] + sorted(df_active_halls['hall'].tolist()))
        if hall_to_manage:
            df_members = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(hall_to_manage,))
            st.write(f"👥 عدد المكلفين: **{len(df_members)}**")
            
            # زر التحميل الجماعي
            if st.button(f"📄 تجهيز ملف تكليفات قاعة {hall_to_manage}"):
                with st.spinner("جاري الدمج..."):
                    out = generate_bulk_docs(df_members, hall_to_manage)
                    if out:
                        st.download_button("📥 تحميل الملف الآن", data=out, file_name=f"تكليفات_{hall_to_manage}.docx")

            st.divider()
            for idx, row in df_members.iterrows():
                col1, col2 = st.columns([3, 1])
                col1.write(f"**{row['name']}** - {row['role']}")
                if col2.button("🗑️ حذف", key=f"manage_del_{row['id']}"):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                    conn.commit(); st.rerun()
    
    st.divider()
    if st.button("⚠️ مسح جميع التكليفات الحالية", type="secondary"):
        c.execute("UPDATE teachers SET hall='', role='', hall_city=''")
        conn.commit(); st.success("تم تصفير النظام"); st.rerun()
