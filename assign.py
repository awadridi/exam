import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os

# =====================================
# 1. إعدادات الواجهة والتصميم
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] {
        border: 1px solid #444; border-radius: 10px;
        background-color: #262730; margin-bottom: 10px;
    }
    div[data-testid="stExpander"] p, div[data-testid="stExpander"] label {
        color: #ffffff !important; font-weight: 500;
    }
    .stButton>button {
        width: 100%; background-color: #28a745; color: white;
        border-radius: 8px; font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات
db_path = os.path.join(os.getcwd(), "exam_data_2026.db")
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة التعبئة المطورة (لضمان استبدال <JOB>)
# =====================================
def generate_from_template(row):
    try:
        doc = Document("template.docx")
        
        # خريطة الاستبدال
        data_map = {
            '<NAME>': row['name'],
            '<ID>': row['id'],
            '<JOB>': row['role'],  # هذه القيمة التي يتم حفظها من القائمة المنسدلة
            '<HALL_NAME>': row['hall'],
            '<HALL_LOCATION>': row['hall_city'],
            '<WORKPLACE>': row['school'],
            '<CITY>': row['city']
        }

        # دالة الاستبدال المحسنة
        def smart_replace(doc_obj, replacements):
            for p in doc_obj.paragraphs:
                for key, val in replacements.items():
                    if key in p.text:
                        # استبدال النص داخل الـ runs للحفاظ على التنسيق قدر الإمكان
                        for run in p.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, str(val))
            
            # نفس العملية للجداول
            for table in doc_obj.tables:
                for row_obj in table.rows:
                    for cell in row_obj.cells:
                        smart_replace(cell, replacements)

        smart_replace(doc, data_map)

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        return None

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("تعيين الموظفين")
    
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            # تحسين العناوين لتظهر الوظيفة المختارة فقط
            disp_hall = row['hall'] if row['hall'] else "لم تُحدد"
            disp_role = row['role'] if row['role'] else "لم تُحدد"
            
            with st.expander(f"👤 {row['name']} | القاعة: {disp_hall} | الوظيفة: {disp_role}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"القاعة لـ {row['id']}", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"الوظيفة لـ {row['id']}", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                        h_city = hall_map.get(sel_hall, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ!")
                        st.rerun()
                    
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button(f"📥 تحميل تكليف {row['role']}", data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# تبويبات الرفع والإدارة (نفس الكود السابق لضمان الثبات)
with tab_upload:
    cu1, cu2 = st.columns(2)
    with cu1:
        f_t = st.file_uploader("xlsx - معلمين", key="u_t")
        if f_t and st.button("رفع المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
            conn.commit()
            st.success("تم")
    with cu2:
        f_h = st.file_uploader("xlsx - قاعات", key="u_h")
        if f_h and st.button("رفع القاعات"):
            dfh = pd.read_excel(f_h)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم")

with tab_manage:
    if st.button("🗑️ مسح الكل"):
        c.execute("DELETE FROM teachers"); c.execute("DELETE FROM halls"); conn.commit(); st.rerun()
