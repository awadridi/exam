import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import re

# =====================================
# 1. إعدادات الواجهة وتثبيت الألوان (CSS)
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    /* تنسيق التطبيق العام */
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    
    /* تنسيق مستطيل الاسم (Expander) */
    .st-emotion-cache-p4m61c { background-color: #1a1c23 !important; border: 1px solid #3d3d3d !important; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    div[data-testid="stExpander"] summary p { color: #ffffff !important; font-weight: bold !important; font-size: 1.1rem !important; }

    /* الحل الوسط: استهداف الأزرار من خلال النوع والمحتوى */
    /* زر الحفظ - أخضر */
    div.stButton > button:has(div:contains("حفظ")) {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
    }
    /* زر الإلغاء - أحمر */
    div.stButton > button:has(div:contains("إلغاء")) {
        background-color: #dc3545 !important;
        color: white !important;
        border: none !important;
    }
    /* زر التحميل - أزرق */
    div.stDownloadButton > button {
        background-color: #007bff !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
    }
    
    /* إجبار لون النصوص الداخلية لتكون بيضاء */
    .stMarkdown p, label { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات
db_path = os.path.join(os.getcwd(), "final_stable_v15.db")
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
                for run in paragraph.runs: run.text = ""
                parts = re.split(r'(<[^>]+>)', text)
                for part in parts:
                    if part in data_map:
                        run = paragraph.add_run(data_map[part] if data_map[part] else "")
                        run.bold = True
                    else:
                        paragraph.add_run(part)

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
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث", "📥 الرفع", "⚙️ الإدارة"])

with tab_search:
    q = st.text_input("ابحث عن الاسم أو الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for i, row in results.iterrows():
            title = f"👤 {row['name']} | {row['role'] if row['role'] else 'لم يعين'} | {row['hall'] if row['hall'] else '-'}"
            with st.expander(title):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"القاعة {row['id']}", [""] + list(pd.read_sql("SELECT hall_name FROM halls", conn)['hall_name']), 
                                          index=0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"الوظيفة {row['id']}", ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"], 
                                          index=0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة الأصلية: {row['school']}")
                    if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                        h_info = pd.read_sql(f"SELECT city FROM halls WHERE hall_name='{sel_hall}'", conn)
                        h_city = h_info.iloc[0]['city'] if not h_info.empty else ""
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                        conn.commit(); st.success("تم الحفظ!"); st.rerun()
                    
                    if row['role'] or row['hall']:
                        if st.button(f"❌ إلغاء التكليف لـ {row['name']}", key=f"del_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit(); st.warning("تم الحذف"); st.rerun()
                    
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button(f"📥 تحميل الكتاب الرسمي", data=file_data, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# التبويبات الأخرى مختصرة للحفاظ على استقرار الكود
with tab_upload:
    f_t = st.file_uploader("ملف الموظفين", type="xlsx")
    if f_t and st.button("تثبيت الموظفين"):
        df = pd.read_excel(f_t)
        for _, r in df.iterrows():
            c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                      (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
        conn.commit(); st.success("تم")

with tab_manage:
    if st.button("🗑️ مسح البيانات"):
        c.execute("DELETE FROM teachers"); c.execute("DELETE FROM halls"); conn.commit(); st.rerun()
