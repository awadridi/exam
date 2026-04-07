import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile

# =====================================
# إعدادات الصفحة والستايل
# =====================================
st.set_page_config(page_title="نظام إدارة التكليف المتكامل", page_icon="📄", layout="centered")

# جعل النصوص من اليمين لليسار في الواجهة
st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# قاعدة البيانات
# =====================================
@st.cache_resource
def get_db_connection():
    return sqlite3.connect("data.db", check_same_thread=False)

conn = get_db_connection()
c = conn.cursor()

c.execute('''
CREATE TABLE IF NOT EXISTS teachers (
    id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, 
    phone TEXT, role TEXT, hall TEXT, accept TEXT DEFAULT 'نعم'
)
''')
c.execute('CREATE TABLE IF NOT EXISTS halls (number TEXT PRIMARY KEY, hall TEXT, city TEXT)')
conn.commit()

# =====================================
# وظائف توليد الملفات (Word & ZIP)
# =====================================

def create_docx(teacher_row):
    """إنشاء ملف Word لكتاب التكليف"""
    doc = Document()
    
    # إعدادات الصفحة (اختياري: إضافة ترويسة)
    header = doc.add_heading('وزارة التربية والتعليم', 0)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    title = doc.add_paragraph('كتاب تكليف رسمي')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(20)
    title.runs[0].bold = True

    doc.add_paragraph(f"\nالتاريخ: 2024/05/20م") # يمكنك تغيير التاريخ تلقائياً
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"\nإلى الزميل/ة: {teacher_row['name']}")
    run.font.size = Pt(14)
    run.bold = True
    
    content = f"""
نحييكم أطيب تحية، ونعلمكم بأنه قد تقرر تكليفكم بالعمل في امتحانات الثانوية العامة وفق البيانات التالية:

• رقم الهوية: {teacher_row['id']}
• المدرسة الأصلية: {teacher_row['school']}
• المهمة الموكلة إليكم: {teacher_row['role']}
• مكان التكليف (القاعة): {teacher_row['hall']}
• المدينة/المنطقة: {teacher_row['city']}

يرجى التواجد في القاعة المذكورة أعلاه في تمام الساعة الثامنة صباحاً، مع الالتزام بكافة التعليمات الصادرة عن رئيس القاعة.

نتمنى لكم التوفيق في مهمتكم.
    """
    para_content = doc.add_paragraph(content)
    para_content.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # التوقيع
    sign = doc.add_paragraph("\n\nتوقيع مدير التربية والتعليم\n...........................")
    sign.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # حفظ الملف في ذاكرة مؤقتة
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# تسجيل الدخول (مبسط حسب طلبك)
# =====================================
if "logged" not in st.session_state:
    st.session_state.logged = False

if not st.session_state.logged:
    st.title("🔐 تسجيل الدخول")
    password = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if password == "1234": # سيتم تغييرها لاحقاً
            st.session_state.logged = True
            st.rerun()
    st.stop()

# =====================================
# الواجهة الرئيسية
# =====================================
st.title("🎓 نظام إدارة التكليفات وطباعة الكتب")

def get_teachers():
    return pd.read_sql("SELECT * FROM teachers WHERE accept='نعم'", conn)

def get_halls():
    return pd.read_sql("SELECT * FROM halls", conn)

teachers = get_teachers()
halls = get_halls()

tab1, tab2, tab3, tab4 = st.tabs(["🔍 التكليف والطباعة", "👨‍🏫 إضافة معلم", "🏢 القاعات", "📊 التقارير العامة"])

# --- تبويب البحث والتكليف ---
with tab1:
    search = st.text_input("🔍 ابحث عن معلم (اسم، هوية، مدرسة)")
    if search:
        result = teachers[
            teachers['name'].str.contains(search, na=False) | 
            teachers['id'].str.contains(search, na=False)
        ]
        
        if not result.empty:
            for i, row in result.iterrows():
                with st.container():
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.write(f"**{row['name']}** - {row['role']} (قاعة: {row['hall'] if row['hall'] else 'غير معين'})")
                    with col2:
                        if row['hall']:
                            # زر توليد الكتاب وتحميله
                            doc_file = create_docx(row)
                            st.download_button(
                                label="📄 تحميل الكتاب",
                                data=doc_file,
                                file_name=f"تكليف_{row['name']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_{row['id']}"
                            )
                    
                    # نموذج التعيين
                    with st.expander("تعديل التكليف"):
                        h_options = [""] + [f"{r['number']} - {r['hall']}" for _, r in halls.iterrows()]
                        new_hall = st.selectbox("اختر القاعة", h_options, key=f"h_{row['id']}")
                        new_role = st.selectbox("المهمة", ["مراقب", "رئيس قاعة", "آذن", "مساعد"], key=f"r_{row['id']}")
                        if st.button("تحديث البيانات", key=f"b_{row['id']}"):
                            h_name = new_hall.split(" - ")[1] if new_hall else ""
                            c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (h_name, new_role, row['id']))
                            conn.commit()
                            st.success("تم التحديث!")
                            st.rerun()
                    st.divider()

# --- تبويب التقارير والتحميل الجماعي ---
with tab4:
    st.subheader("تحميل جميع كتب التكليف")
    assigned_teachers = teachers[teachers['hall'] != ""]
    
    st.write(f"عدد المعلمين المعينين حالياً: {len(assigned_teachers)}")
    
    if st.button("📦 توليد وتحميل جميع الكتب (ZIP)"):
        if not assigned_teachers.empty:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for _, row in assigned_teachers.iterrows():
                    doc_bytes = create_docx(row)
                    zip_file.writestr(f"تكليف_{row['name']}.docx", doc_bytes.getvalue())
            
            zip_buffer.seek(0)
            st.download_button(
                label="📥 تحميل الملف المضغوط الآن",
                data=zip_buffer,
                file_name="جميع_كتب_التكليف.zip",
                mime="application/zip"
            )
        else:
            st.error("لا يوجد معلمين معينين لتحميل كتبهم!")

# --- باقي التبويبات (إضافة معلم وقاعات) تُترك كما كانت في الكود السابق لتوفير المساحة ---
with tab2:
    # (كود إضافة معلم كما في السابق)
    st.subheader("إضافة معلم جديد")
    with st.form("add_teacher"):
        c1, c2 = st.columns(2)
        name = c1.text_input("الاسم")
        idd = c2.text_input("الهوية")
        school = c1.text_input("المدرسة")
        phone = c2.text_input("الجوال")
        if st.form_submit_button("حفظ"):
            c.execute("INSERT INTO teachers (id, name, school, phone) VALUES (?,?,?,?)", (idd, name, school, phone))
            conn.commit()
            st.rerun()

with tab3:
    st.subheader("إضافة قاعة")
    with st.form("add_hall"):
        h_n = st.text_input("رقم القاعة")
        h_m = st.text_input("اسم القاعة")
        if st.form_submit_button("إضافة"):
            c.execute("INSERT INTO halls VALUES (?,?,?)", (h_n, h_m, ""))
            conn.commit()
            st.rerun()
