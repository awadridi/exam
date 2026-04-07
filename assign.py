import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile

# =====================================
# 1. إعدادات الصفحة والتنسيق
# =====================================
st.set_page_config(page_title="نظام إدارة التكليفات", page_icon="🎓", layout="wide")

# تنسيق الواجهة لتدعم اللغة العربية (RTL)
st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] div[role="button"] p { font-size: 1.2rem; font-weight: bold; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 2. قاعدة البيانات SQLite
# =====================================
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect("data.db", check_same_thread=False)
    return conn

conn = get_db_connection()
c = conn.cursor()

# إنشاء الجداول (مطابق لأعمدة ملف الاكسل الخاص بك)
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
# 3. وظائف النظام (Word & Logic)
# =====================================

def create_docx(row):
    """توليد ملف Word احترافي لكل معلم"""
    doc = Document()
    
    # ترويسة الكتاب
    header = doc.add_paragraph("دولة فلسطين\nوزارة التربية والتعليم\nمديرية التربية والتعليم")
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('كتاب تكليف رسمي', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    content = f"""
    التاريخ: 2026/04/07م
    
    إلى الزميل/ة: {row['name']} المحترم/ة
    رقم الهوية: {row['id']}
    المدرسة الأصلية: {row['school']}
    
    نحييكم أطيب تحية، ونعلمكم بأنه قد تقرر تكليفكم بالعمل في امتحانات الثانوية العامة وفق البيانات التالية:
    
    • المهمة الموكلة إليكم: {row['role']}
    • مكان التكليف (القاعة): {row['hall']}
    • المدينة/المنطقة: {row['city']}
    
    يرجى التواجد في القاعة المذكورة أعلاه في تمام الساعة الثامنة صباحاً.
    
    نتمنى لكم التوفيق في مهمتكم.
    """
    run = p.add_run(content)
    run.font.size = Pt(13)
    
    doc.add_paragraph("\nتوقيع مدير التربية والتعليم\n...........................").alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# 4. نظام تسجيل الدخول
# =====================================
if "logged" not in st.session_state:
    st.session_state.logged = False

if not st.session_state.logged:
    st.title("🔐 تسجيل الدخول للنظام")
    pwd = st.text_input("أدخل كلمة المرور", type="password")
    if st.button("دخول"):
        if pwd == "1234": # يمكنك تغييرها لاحقاً
            st.session_state.logged = True
            st.rerun()
    st.stop()

# =====================================
# 5. واجهة المستخدم الرئيسية (Tabs)
# =====================================
tab_search, tab_upload, tab_reports = st.tabs(["🔍 البحث والطباعة", "📥 رفع بيانات Excel", "📊 تقارير وتحميل جماعي"])

# --- التبويب الأول: البحث والتعيين ---
with tab_search:
    st.subheader("البحث عن المعلمين المكلفين")
    search_q = st.text_input("أدخل الاسم أو رقم الهوية للبحث...")
    
    # جلب البيانات الحالية
    df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
    
    if search_q:
        results = df_teachers[
            df_teachers['name'].str.contains(search_q, na=False) | 
            df_teachers['id'].astype(str).str.contains(search_q)
        ]
        
        if not results.empty:
            for i, row in results.iterrows():
                with st.expander(f"👤 {row['name']} - {row['role']}"):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write(f"**رقم الهوية:** {row['id']}")
                        st.write(f"**المدرسة:** {row['school']}")
                        st.write(f"**القاعة:** {row['hall'] if row['hall'] else '❌ غير محددة'}")
                    
                    with col2:
                        if row['hall'] and row['hall'] != "":
                            doc_bytes = create_docx(row)
                            st.download_button(
                                label="📄 تحميل كتاب التكليف (Word)",
                                data=doc_bytes,
                                file_name=f"تكليف_{row['name']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"btn_{row['id']}"
                            )
                        else:
                            st.warning("يجب تحديد قاعة للمعلم لتتمكن من طباعة الكتاب.")
        else:
            st.info("لا توجد نتائج مطابقة للبحث.")

# --- التبويب الثاني: استيراد البيانات ---
with tab_upload:
    st.subheader("استيراد البيانات من ملف Excel")
    st.info("تأكد أن ملف الاكسل يحتوي على الأعمدة التالية بنفس المسميات: id, name, school, city, phone, role, hall, accept or r")
    
    up_file = st.file_uploader("اختر ملف XLSX", type="xlsx")
    
    if up_file:
        try:
            df_excel = pd.read_excel(up_file)
            st.dataframe(df_excel.head(10))
            
            if st.button("🚀 اعتماد البيانات وحفظها في النظام"):
                with st.spinner("جاري الحفظ..."):
                    for _, row in df_excel.iterrows():
                        c.execute('''
                        INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, accept)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            str(row['id']), 
                            row['name'], 
                            row['school'], 
                            row['city'], 
                            str(row['phone']), 
                            row['role'], 
                            row['hall'], 
                            row['accept or r']
                        ))
                    conn.commit()
                st.success(f"تم بنجاح! تم استيراد {len(df_excel)} سجل.")
                st.rerun()
        except Exception as e:
            st.error(f"حدث خطأ في قراءة الملف: {e}")

# --- التبويب الثالث: التحميل الجماعي ---
with tab_reports:
    st.subheader("أدوات الإدارة الجماعية")
    
    # تصفية المعلمين الذين لديهم قاعات فقط
    ready_teachers = pd.read_sql("SELECT * FROM teachers WHERE hall IS NOT NULL AND hall != ''", conn)
    
    st.write(f"عدد المعلمين الجاهزين للطباعة (لديهم قاعات): {len(ready_teachers)}")
    
    if st.button("📦 تحميل كافة الكتب في ملف ZIP"):
        if not ready_teachers.empty:
            zip_buffer
