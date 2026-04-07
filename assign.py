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
    
    /* تنسيق مستطيل الاسم (Expander) */
    div[data-testid="stExpander"] {
        border: 1px solid #444 !important;
        background-color: #1a1c23 !important;
    }
    div[data-testid="stExpander"] summary p {
        color: #ffffff !important;
        font-weight: bold !important;
    }

    /* ألوان الأزرار إجبارية */
    /* حفظ - أخضر */
    button[key^="btn_"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
    }
    /* إلغاء - أحمر */
    button[key^="del_"] {
        background-color: #dc3545 !important;
        color: white !important;
        border: none !important;
    }
    /* تحميل - أزرق */
    .stDownloadButton button {
        background-color: #007bff !important;
        color: white !important;
        width: 100% !important;
    }
    
    label, .stMarkdown p { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات - محاولة جعلها ثابتة قدر الإمكان
db_path = "data_system_v16.db"
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة الاستبدال (البولد الذكي)
# =====================================
def generate_from_template(row):
    try:
        doc = Document("template.docx")
        replacements = {
            '<NAME>': str(row['name']), '<ID>': str(row['id']),
            '<JOB>': str(row['role']), '<HALL_NAME>': str(row['hall']),
            '<HALL_LOCATION>': str(row['hall_city']), '<WORKPLACE>': str(row['school']),
            '<CITY>': str(row['city'])
        }
        def apply_smart_bold_replace(paragraph, data_map):
            text = paragraph.text
            if any(key in text for key in data_map):
                for run in paragraph.runs: run.text = ""
                parts = re.split(r'(<[^>]+>)', text)
                for part in parts:
                    if part in data_map:
                        run = paragraph.add_run(data_map[part] if data_map[part] else "")
                        run.bold = True
                        run.font.size = Pt(14)
                    else:
                        paragraph.add_run(part)
                        run.font.size = Pt(14)

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

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("إدارة الموظفين")
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for i, row in results.iterrows():
            title = f"👤 {row['name']} | {row['role'] if row['role'] else '-'} | {row['hall'] if row['hall'] else '-'}"
            with st.expander(title):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"اختر القاعة لـ {row['id']}", hall_list, index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"اختر الوظيفة لـ {row['id']}", role_list, index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    col_save, col_del = st.columns(2)
                    with col_save:
                        if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                            h_city = hall_map.get(sel_hall, "")
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                            conn.commit(); st.success("تم الحفظ!"); st.rerun()
                    with col_del:
                        if row['role'] or row['hall']:
                            if st.button("❌ إلغاء التكليف", key=f"del_{row['id']}"):
                                c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                                conn.commit(); st.warning("تم الإلغاء"); st.rerun()
                    
                    if row['hall'] and row['role']:
                        f_data = generate_from_template(row)
                        if f_data:
                            st.download_button("📥 تحميل التكليف", data=f_data, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# تبويب الرفع (تم إرجاع زر القاعات)
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        f_teachers = st.file_uploader("رفع ملف الموظفين (Excel)", type="xlsx")
        if f_teachers and st.button("تثبيت قائمة الموظفين"):
            df = pd.read_excel(f_teachers)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
            conn.commit()
            st.success("تم رفع الموظفين بنجاح")
            
    with col_u2:
        f_halls = st.file_uploader("رفع ملف القاعات (Excel)", type="xlsx")
        if f_halls and st.button("تثبيت قائمة القاعات"):
            dfh = pd.read_excel(f_halls)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم رفع القاعات بنجاح")

    # تأكد أن هذه الأسطر تبدأ من نفس مستوى col_u1 (خارج الـ with col_u2)
    st.divider() 
    st.subheader("📄 رفع نموذج كتاب التكليف")
    f_template = st.file_uploader("ارفع ملف الوورد (template.docx)", type="docx", key="u_docx")
    if f_template and st.button("تثبيت النموذج الجديد"):
        with open("template.docx", "wb") as f:
            f.write(f_template.getbuffer())
        st.success("✅ تم تحديث نموذج كتاب التكليف بنجاح!")

with tab_manage:
    if st.button("⚠️ مسح شامل لقاعدة البيانات"):
        c.execute("DELETE FROM teachers"); c.execute("DELETE FROM halls"); conn.commit(); st.rerun()
