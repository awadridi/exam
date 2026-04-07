import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile

# =====================================
# 1. إعدادات الصفحة والتنسيق العربي
# =====================================
st.set_page_config(page_title="نظام إدارة التكليفات الرقمي", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    th, td { text-align: right !important; }
    .stButton>button { width: 100%; border-radius: 8px; background-color: #007bff; color: white; height: 3em; }
    div[data-testid="stExpander"] p { font-size: 1.1rem; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 2. إدارة قاعدة البيانات
# =====================================
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect("data.db", check_same_thread=False)
    return conn

conn = get_db_connection()
c = conn.cursor()

# إنشاء الجدول بالأعمدة الأساسية
c.execute('''
CREATE TABLE IF NOT EXISTS teachers (
    id TEXT PRIMARY KEY,
    name TEXT,
    school TEXT,
    city TEXT,
    phone TEXT,
    role TEXT,
    hall TEXT,
    accept TEXT
)
''')
conn.commit()

# =====================================
# 3. دالة توليد ملف الـ Word
# =====================================
def create_docx(row):
    doc = Document()
    
    # ترويسة افتراضية - يمكنك تعديل النص بين الاقتباسات
    header = doc.add_paragraph("دولة فلسطين\nوزارة التربية والتعليم")
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    title = doc.add_heading('بطاقة تكليف لمهمة امتحان', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    content = f"""
    بناءً على مقتضيات المصلحة العامة، فقد تقرر تكليفكم بالمهمة المذكورة أدناه:
    
    • الاسم الكامل: {row['name']}
    • رقم الهوية: {row['id']}
    • المدرسة الأصلية: {row['school']}
    • المهمة الموكلة: {row['role']}
    • مكان العمل (القاعة): {row['hall']}
    • المدينة/المنطقة: {row['city']}
    
    نتمنى لكم التوفيق والسداد في أداء الأمانة.
    """
    run = p.add_run(content)
    run.font.size = Pt(13)
    
    doc.add_paragraph("\nتوقيع جهة الاختصاص\n...........................").alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# 4. نظام الدخول
# =====================================
if "logged" not in st.session_state:
    st.session_state.logged = False

if not st.session_state.logged:
    st.title("🔐 تسجيل الدخول")
    col_l, col_r = st.columns([1, 2])
    with col_l:
        pwd = st.text_input("أدخل كلمة المرور", type="password")
        if st.button("دخول"):
            if pwd == "1234": # غيرها لاحقاً
                st.session_state.logged = True
                st.rerun()
    st.stop()

# =====================================
# 5. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والطباعة", "📥 استيراد من إكسل", "⚙️ الإدارة"])

# --- التبويب الأول: البحث ---
with tab_search:
    st.subheader("قائمة المعلمين والتكليفات")
    search_q = st.text_input("ابحث عن اسم، هوية، أو مدرسة...")
    
    df_data = pd.read_sql("SELECT * FROM teachers", conn)
    
    if not df_data.empty:
        if search_q:
            df_data = df_data[
                df_data['name'].str.contains(search_q, na=False) | 
                df_data['id'].astype(str).str.contains(search_q) |
                df_data['school'].str.contains(search_q, na=False)
            ]

        for i, row in df_data.iterrows():
            with st.expander(f"👤 {row['name']} - {row['role']} (قاعة: {row['hall'] if row['hall'] else 'لم تُعين'})"):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**رقم الهوية:** {row['id']}")
                    st.write(f"**المدينة:** {row['city']}")
                with c2:
                    if row['hall']:
                        doc_file = create_docx(row)
                        st.download_button(
                            label=f"📄 تحميل كتاب {row['name']}",
                            data=doc_file,
                            file_name=f"تكليف_{row['id']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{row['id']}"
                        )
                    else:
