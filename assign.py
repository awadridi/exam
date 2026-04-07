import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import re

# =====================================
# 1. إعدادات الواجهة وتنسيق الأزرار القوي
# =====================================
st.set_page_config(page_title="نظام التكليفات الذكي 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] {
        border: 1px solid #444; border-radius: 10px;
        background-color: #262730; margin-bottom: 10px;
    }
    
    /* زر حفظ البيانات - أخضر */
    button[kind="secondary"]:has(div:contains("حفظ")) {
        background-color: #28a745 !important;
        color: white !important;
        border: 1px solid #1e7e34 !important;
    }

    /* زر إلغاء التكليف - أحمر */
    button[kind="secondary"]:has(div:contains("إلغاء")) {
        background-color: #dc3545 !important;
        color: white !important;
        border: 1px solid #bd2130 !important;
    }

    /* زر تحميل التكليف - أزرق */
    button[kind="primary"] {
        background-color: #007bff !important;
        color: white !important;
        border: 1px solid #0069d9 !important;
    }
    
    /* تحسين شكل النصوص داخل الأزرار */
    .stButton p, .stDownloadButton p {
        font-weight: bold !important;
        font-size: 16px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات
db_path = os.path.join(os.getcwd(), "final_fix_v13.db")
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
            '<NAME>': str(row['name']),
            '<ID>': str(row['id']),
            '<JOB>': str(row['role']),
            '<HALL_NAME>': str(row['hall']),
            '<HALL_LOCATION>': str(row['hall_city']),
            '<WORKPLACE>': str(row['school']),
            '<CITY>': str(row['city'])
        }

        def apply_smart_bold_replace(paragraph, data_map):
            text = paragraph.text
            if any(key in text for key in data_map):
                for run in paragraph.runs:
                    run.text = ""
                parts = re.split(r'(<[^>]+>)', text)
                for part in parts:
                    if part in data_map:
                        run = paragraph.add_run(data_map[part] if data_map[part] else "")
                        run.bold = True
                    else:
                        paragraph.add_run(part)

        for p in doc.paragraphs:
            apply_smart_bold_replace(p, replacements)
        for table in doc.tables:
            for r in table.rows:
                for cell in r.cells:
                    for p in cell.paragraphs:
                        apply_smart_bold_replace(p, replacements)

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception:
        return None

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("إدارة التكليفات")
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for i, row in results.iterrows():
            title = f"👤 {row['name']}"
            if row['role']: title += f" | {row['role']}"
            if row['hall']: title += f" | {row['hall']}"
            
            with st.expander(title):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"القاعة لـ {row['id']}", hall_list, index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"الوظيفة لـ {row['id']}", role_list, index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    b1, b2 = st.columns(2)
                    with b1:
                        if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                            h_city = hall_map.get(sel_hall, "")
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                            conn.commit()
                            st.success("تم الحفظ!")
                            st.rerun()
                    with b2:
                        if row['role'] or row['hall']:
                            if st.button(f"❌ إلغاء التكليف", key=f"del_{row['id']}"):
                                c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                                conn.commit()
                                st.warning("تم الحذف")
                                st.rerun()
                    
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button(f"📥 تحميل التكليف", data=file_data, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# التبويبات الأخرى كما هي
with tab_upload:
    cu1, cu2 = st.columns(2)
    with cu1:
        f_t = st.file_uploader("ملف الموظفين", type="xlsx")
        if f_t and st.button("تأكيد الرفع", key="confirm_t"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
            conn.commit()
            st.success("تم الرفع")
    with cu2:
        f_h = st.file_uploader("ملف القاعات", type="xlsx")
        if f_h and st.button("رفع القاعات", key="confirm_h"):
            dfh = pd.read_excel(f_h)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم الرفع")

with tab_manage:
    if st.button("🗑️ مسح شامل للبيانات", key="wipe_all"):
        c.execute("DELETE FROM teachers"); c.execute("DELETE FROM halls"); conn.commit(); st.rerun()
