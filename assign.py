import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile

# =====================================
# 1. إعدادات الصفحة
# =====================================
st.set_page_config(page_title="نظام إدارة التكليفات", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    .stButton>button { width: 100%; border-radius: 8px; background-color: #28a745; color: white; }
    div[data-testid="stExpander"] { background-color: #f8f9fa; border: 1px solid #ddd; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 2. قاعدة البيانات
# =====================================
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect("data.db", check_same_thread=False)
    return conn

conn = get_db_connection()
c = conn.cursor()

# إنشاء الجداول
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, accept TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 3. دالة توليد Word
# =====================================
def create_docx(row):
    doc = Document()
    doc.add_paragraph("وزارة التربية والتعليم").alignment = WD_ALIGN_PARAGRAPH.CENTER
    title = doc.add_heading('بطاقة تكليف رسمي', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    content = f"""
    الاسم: {row['name']}
    رقم الهوية: {row['id']}
    المهمة: {row['role']}
    مكان التكليف: {row['hall']}
    المنطقة: {row['city']}
    """
    run = p.add_run(content)
    run.font.size = Pt(13)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# 4. تسجيل الدخول
# =====================================
if "logged" not in st.session_state:
    st.session_state.logged = False

if not st.session_state.logged:
    st.title("🔐 دخول النظام")
    pwd = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if pwd == "1234":
            st.session_state.logged = True
            st.rerun()
    st.stop()

# =====================================
# 5. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع ملفات Excel", "⚙️ الإدارة"])

# --- التبويب الأول: البحث والتعيين الفردي ---
with tab_search:
    st.subheader("إدارة تعيينات الموظفين")
    
    # جلب بيانات القاعات والمهام
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_list = [""] + (df_halls['hall_name'].tolist())
    role_list = ["مراقب", "رئيس قاعة", "مساعد رئيس قاعة", "آذن", "عضو"]

    search_q = st.text_input("ابحث عن موظف لتعديل تكليفه...")
    
    df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
    
    if search_q and not df_teachers.empty:
        results = df_teachers[df_teachers['name'].str.contains(search_q, na=False) | df_teachers['id'].astype(str).str.contains(search_q)]
        
        for i, row in results.iterrows():
            with st.expander(f"👤 {row['name']} (الحالة الحالية: {row['hall'] if row['hall'] else 'غير معين'})"):
                col1, col2 = st.columns(2)
                with col1:
                    new_hall = st.selectbox(f"اختر القاعة لـ {row['name']}", hall_list, index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    new_role = st.selectbox(f"حدد المهمة", role_list, index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                
                with col2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("✅ حفظ التعيين", key=f"save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (new_hall, new_role, row['id']))
                        conn.commit()
                        st.success("تم التحديث")
                        st.rerun()
                    
                    if row['hall']:
                        doc = create_docx(row)
                        st.download_button("📄 تحميل الكتاب", data=doc, file_name=f"تكليف_{row['id']}.docx", key=f"dl_{row['id']}")

# --- التبويب الثاني: رفع الملفات (المعلمين والقاعات) ---
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        st.subheader("1. ملف المعلمين")
        file_t = st.file_uploader("ارفع ملف الموظفين (xlsx)", type="xlsx", key="u_t")
        if file_t and st.button("حفظ الموظفين"):
            df = pd.read_excel(file_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, accept) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), str(r.get('hall','')), str(r.get('accept or r','نعم'))))
            conn.commit()
            st.success("تم رفع بيانات المعلمين")

    with col_u2:
        st.subheader("2. ملف القاعات")
        file_h = st.file_uploader("ارفع ملف القاعات (xlsx)", type="xlsx", key="u_h")
        if file_h and st.button("حفظ القاعات"):
            dfh = pd.read_excel(file_h)
            for _, r in dfh.iterrows():
                # تأكد أن أسماء الأعمدة في ملف القاعات هي number, hall, city
                c.execute("INSERT OR REPLACE INTO halls (number, hall_name, city) VALUES (?,?,?)",
                          (str(r.get('number','')), str(r.get('hall','')), str(r.get('city',''))))
            conn.commit()
            st.success("تم رفع بيانات القاعات")

# --- التبويب الثالث: الإدارة ---
with tab_manage:
    if st.button("🗑️ مسح كل البيانات والبدء من جديد"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
