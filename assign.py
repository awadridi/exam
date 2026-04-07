import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os

# =====================================
# 1. إعدادات الواجهة والتصميم (ثيم مظلم مع خطوط واضحة)
# =====================================
st.set_page_config(page_title="نظام تكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    
    /* تنسيق المستطيلات لتبدو واضحة جداً في البحث */
    div[data-testid="stExpander"] {
        border: 1px solid #444;
        border-radius: 10px;
        background-color: #1e1e1e;
        margin-bottom: 10px;
        color: white;
    }
    
    /* جعل الخط أبيض وعريض داخل النتائج */
    div[data-testid="stExpander"] p, 
    div[data-testid="stExpander"] span,
    div[data-testid="stExpander"] label,
    div[data-testid="stExpander"] div {
        color: #ffffff !important;
        font-weight: 500;
    }

    /* تحسين الأزرار */
    .stButton>button {
        width: 100%;
        background-color: #28a745;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 2. إدارة قاعدة البيانات (دائمة)
# =====================================
# استخدام مسار مطلق لضمان عدم ضياع الملف عند تحديث الصفحة
db_path = os.path.join(os.getcwd(), "database_2026.db")
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()

def init_db():
    c.execute('''CREATE TABLE IF NOT EXISTS teachers 
                 (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS halls 
                 (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
    conn.commit()

init_db()

# =====================================
# 3. وظيفة تعبئة الوورد (مع الحفاظ على Bold ووسم JOB)
# =====================================
def generate_docx(row):
    try:
        doc = Document("template.docx")
        
        def replace_in_container(container, search, replace):
            for p in container.paragraphs:
                for run in p.runs:
                    if search in run.text:
                        run.text = run.text.replace(search, str(replace))
            for table in container.tables:
                for r in table.rows:
                    for cell in r.cells:
                        replace_in_container(cell, search, replace)

        # ربط البيانات بالوسوم الموجودة في ملفك
        mapping = {
            '<NAME>': row['name'],
            '<ID>': row['id'],
            '<JOB>': row['role'],
            '<HALL_NAME>': row['hall'],
            '<HALL_LOCATION>': row['hall_city'],
            '<WORKPLACE>': row['school'],
            '<CITY>': row['city']
        }

        for key, val in mapping.items():
            replace_in_container(doc, key, val if val else "")

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except:
        return None

# =====================================
# 4. واجهة التطبيق
# =====================================
t1, t2, t3 = st.tabs(["🔍 البحث والتعيين", "📥 رفع ملفات Excel", "⚙️ الإدارة"])

# --- التبويب الأول: البحث ---
with t1:
    st.subheader("إدارة تكليفات المعلمين")
    
    # جلب القاعات المتاحة
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    h_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    h_list = [""] + list(h_map.keys())
    roles = ["رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    search_q = st.text_input("ابحث عن المعلم بالاسم أو رقم الهوية")
    df_teachers = pd.read_sql("SELECT * FROM teachers", conn)

    if search_q and not df_teachers.empty:
        results = df_teachers[df_teachers['name'].str.contains(search_q, na=False) | 
                             df_teachers['id'].astype(str).str.contains(search_q)]
        
        for i, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall']} | الوظيفة: {row['role']}"):
                col1, col2 = st.columns(2)
                with col1:
                    s_hall = st.selectbox("اختر القاعة", h_list, 
                                        index=h_list.index(row['hall']) if row['hall'] in h_list else 0, 
                                        key=f"sh_{row['id']}")
                    s_role = st.selectbox("اختر الوظيفة (JOB)", roles, 
                                        index=roles.index(row['role']) if row['role'] in roles else 0, 
                                        key=f"sr_{row['id']}")
                with col2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("💾 حفظ التعديلات", key=f"save_{row['id']}"):
                        h_city = h_map.get(s_hall, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", 
                                 (s_hall, s_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ بنجاح")
                        st.rerun()
                    
                    if row['hall'] and row['role']:
                        doc_file = generate_docx(row)
                        if doc_file:
                            st.download_button("📥 تحميل كتاب التكليف", doc_file, 
                                             file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# --- التبويب الثاني: الرفع ---
with t2:
    st.info("ارفع الملفات هنا مرة واحدة وسيقوم النظام بحفظها للأبد.")
    c1, c2 = st.columns(2)
    with c1:
        f_teachers = st.file_uploader("ملف المعلمين (Excel)", type="xlsx")
        if f_teachers and st.button("تثبيت المعلمين في النظام"):
            df = pd.read_excel(f_teachers)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), "", ""))
            conn.commit()
            st.success("تم الحفظ بنجاح")

    with c2:
        f_halls = st.file_uploader("ملف القاعات (Excel)", type="xlsx")
        if f_halls and st.button("تثبيت القاعات في النظام"):
            dfh = pd.read_excel(f_halls)
            c.execute("DELETE FROM halls") # تنظيف القائمة لرفع قائمة جديدة
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم تحديث القاعات بنجاح")

# --- التبويب الثالث: الإدارة ---
with t3:
    st.warning("⚠️ تحذير: هذه الأزرار ستمسح البيانات المخزنة نهائياً.")
    if st.button("🗑️ حذف جميع المعلمين والقاعات"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
