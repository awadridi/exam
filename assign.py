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

# قاعدة البيانات الثابتة
db_path = os.path.join(os.getcwd(), "final_system_v5.db")
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة الاستبدال الاحترافية (حل مشكلة JOB)
# =====================================
def generate_from_template(row):
    try:
        doc = Document("template.docx")
        
        # القائمة التي سيتم استبدالها في ملف الورد
        replacements = {
            '<NAME>': str(row['name']),
            '<ID>': str(row['id']),
            '<JOB>': str(row['role']), # المسمى الوظيفي الذي اخترته وحفظته
            '<HALL_NAME>': str(row['hall']),
            '<HALL_LOCATION>': str(row['hall_city']),
            '<WORKPLACE>': str(row['school']),
            '<CITY>': str(row['city'])
        }

        # دالة استبدال متطورة تبحث في كل الفقرات والجداول
        def docx_replace(doc_obj, data):
            # الاستبدال في الفقرات
            for p in doc_obj.paragraphs:
                for key, val in data.items():
                    if key in p.text:
                        # هذه الطريقة تضمن استبدال الكلمة حتى لو كان تنسيقها Bold
                        inline = p.runs
                        for i in range(len(inline)):
                            if key in inline[i].text:
                                text = inline[i].text.replace(key, val if val else "")
                                inline[i].text = text

            # الاستبدال داخل الجداول
            for table in doc_obj.tables:
                for row_cell in table.rows:
                    for cell in row_cell.cells:
                        docx_replace(cell, data)

        docx_replace(doc, replacements)

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
    st.subheader("تعيين الموظفين")
    
    # جلب بيانات القاعات
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            # بناء العنوان بحيث يظهر الاسم فقط، ثم يظهر الباقي بعد الحفظ
            title_parts = [f"👤 {row['name']}"]
            if row['role']: title_parts.append(f"الوظيفة: {row['role']}")
            if row['hall']: title_parts.append(f"القاعة: {row['hall']}")
            
            with st.expander(" | ".join(title_parts)):
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
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", 
                                 (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ بنجاح!")
                        st.rerun()
                    
                    # لا يظهر التحميل إلا إذا تم اختيار قاعة ووظيفة
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button(f"📥 تحميل تكليف ({row['role']})", data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# التبويبات الأخرى (الرفع والإدارة)
with tab_upload:
    cu1, cu2 = st.columns(2)
    with cu1:
        f_t = st.file_uploader("ملف الموظفين", type="xlsx", key="u_t")
        if f_t and st.button("تأكيد رفع المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                # إفراغ خانات الوظيفة والقاعة عند الرفع الجديد لضمان عدم ظهور مسميات قديمة
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
            conn.commit()
            st.success("تم الرفع")
    with cu2:
        f_h = st.file_uploader("ملف القاعات", type="xlsx", key="u_h")
        if f_h and st.button("تأكيد رفع القاعات"):
            dfh = pd.read_excel(f_h)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم الرفع")

with tab_manage:
    if st.button("🗑️ مسح شامل للبيانات"):
        c.execute("DELETE FROM teachers"); c.execute("DELETE FROM halls"); conn.commit(); st.rerun()
