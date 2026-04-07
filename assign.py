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
st.set_page_config(page_title="نظام التكليفات", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    .stButton>button { width: 100%; border-radius: 8px; background-color: #28a745; color: white; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 2. قاعدة البيانات
# =====================================
@st.cache_resource
def get_db_connection():
    # استخدام اسم قاعدة بيانات جديد لضمان التحديث
    return sqlite3.connect("final_exam_data.db", check_same_thread=False)

conn = get_db_connection()
c = conn.cursor()

def init_db():
    c.execute('''CREATE TABLE IF NOT EXISTS teachers 
                 (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, accept TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS halls 
                 (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
    conn.commit()

init_db()

def find_col(df, keywords):
    for col in df.columns:
        if any(key in str(col).lower() for key in keywords):
            return col
    return None

# =====================================
# 3. دالة Word
# =====================================
def create_docx(row):
    doc = Document()
    doc.add_paragraph("دولة فلسطين\nوزارة التربية والتعليم").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading('بطاقة تكليف رسمي', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    content = f"الاسم: {row['name']}\nرقم الهوية: {row['id']}\nالمهمة: {row['role']}\nالقاعة: {row['hall']}\nالمدينة: {row['city']}"
    run = p.add_run(content)
    run.font.size = Pt(13)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# 4. تسجيل الدخول
# =====================================
if "logged" not in st.session_state: st.session_state.logged = False
if not st.session_state.logged:
    pwd = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if pwd == "1234":
            st.session_state.logged = True
            st.rerun()
    st.stop()

# =====================================
# 5. الواجهة
# =====================================
tab1, tab2, tab3 = st.tabs(["🔍 البحث", "📥 الرفع", "⚙️ الإدارة"])

with tab1:
    df_h = pd.read_sql("SELECT * FROM halls", conn)
    h_list = [""] + df_h['hall_name'].tolist() if not df_h.empty else [""]
    roles = ["مراقب", "رئيس قاعة", "مساعد", "آذن"]
    
    q = st.text_input("ابحث بالاسم أو الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    if q and not df_t.empty:
        res = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for _, r in res.iterrows():
            with st.expander(f"👤 {r['name']}"):
                c1, c2 = st.columns(2)
                new_h = c1.selectbox("القاعة", h_list, index=h_list.index(r['hall']) if r['hall'] in h_list else 0, key=f"h{r['id']}")
                new_r = c1.selectbox("المهمة", roles, index=roles.index(r['role']) if r['role'] in roles else 0, key=f"r{r['id']}")
                if c2.button("حفظ", key=f"s{r['id']}"):
                    c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (new_h, new_r, r['id']))
                    conn.commit()
                    st.success("تم الحفظ")
                    st.rerun()
                if r['hall']: st.download_button("تحميل", create_docx(r), f"{r['id']}.docx", key=f"d{r['id']}")

with tab2:
    st.subheader("رفع الملفات")
    f_h = st.file_uploader("1. ملف القاعات", type="xlsx")
    if f_h and st.button("تثبيت القاعات"):
        df = pd.read_excel(f_h)
        # حذف وإعادة إنشاء الجدول لتجنب OperationalError
        c.execute("DROP TABLE IF EXISTS halls")
        c.execute("CREATE TABLE halls (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)")
        
        c_n = find_col(df, ['number', 'رقم'])
        c_h = find_col(df, ['hall', 'قاعة'])
        c_c = find_col(df, ['city', 'مدينة', 'سكن'])
        
        for _, r in df.iterrows():
            c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r[c_n]), str(r[c_h]), str(r[c_c]) if c_c else ""))
        conn.commit()
        st.success("تم رفع القاعات بنجاح ✅")

    f_t = st.file_uploader("2. ملف الموظفين", type="xlsx")
    if f_t and st.button("تثبيت الموظفين"):
        df = pd.read_excel(f_t)
        c.execute("DROP TABLE IF EXISTS teachers")
        c.execute("CREATE TABLE teachers (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, accept TEXT)")
        
        cid = find_col(df, ['id', 'هوية'])
        cnm = find_col(df, ['name', 'اسم'])
        
        for _, r in df.iterrows():
            c.execute("INSERT INTO teachers VALUES (?,?,?,?,?,?,?,?)", 
                      (str(r[cid]), str(r[cnm]), str(df.columns[2]), "", "", "", "", "نعم"))
        conn.commit()
        st.success("تم رفع الموظفين بنجاح ✅")

with tab3:
    if st.button("🗑️ تفريغ النظام"):
        c.execute("DROP TABLE IF EXISTS teachers")
        c.execute("DROP TABLE IF EXISTS halls")
        conn.commit()
        st.rerun()
