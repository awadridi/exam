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
            if pwd == "1234": 
                st.session_state.logged = True
                st.rerun()
            else:
                st.error("كلمة المرور غير صحيحة")
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
                df_data['name'].astype(str).str.contains(search_q, na=False) | 
                df_data['id'].astype(str).str.contains(search_q) |
                df_data['school'].astype(str).str.contains(search_q, na=False)
            ]

        for i, row in df_data.iterrows():
            with st.expander(f"👤 {row['name']} - {row['role']} (قاعة: {row['hall'] if row['hall'] else 'لم تُعين'})"):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**رقم الهوية:** {row['id']}")
                    st.write(f"**المدينة:** {row['city']}")
                with c2:
                    if row['hall'] and str(row['hall']).strip() != "":
                        doc_file = create_docx(row)
                        st.download_button(
                            label=f"📄 تحميل كتاب {row['name']}",
                            data=doc_file,
                            file_name=f"تكليف_{row['id']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{row['id']}"
                        )
                    else:
                        st.warning("⚠️ لا توجد قاعة معينة حالياً.")
    else:
        st.info("النظام فارغ. يرجى رفع ملف الإكسل أولاً.")

# --- التبويب الثاني: الرفع ---
with tab_upload:
    st.subheader("تحميل البيانات من ملف Excel")
    up_file = st.file_uploader("اختر ملف الإكسل", type="xlsx")
    
    if up_file:
        df_excel = pd.read_excel(up_file)
        st.write("معاينة أولية للبيانات:")
        st.dataframe(df_excel.head())

        if st.button("🚀 تأكيد حفظ البيانات في النظام"):
            with st.spinner("جاري استيراد البيانات..."):
                try:
                    def get_col(df, keywords):
                        for col in df.columns:
                            if any(key in str(col).lower() for key in keywords):
                                return col
                        return None

                    col_id = get_col(df_excel, ['id', 'هوية'])
                    col_name = get_col(df_excel, ['name', 'اسم'])
                    col_school = get_col(df_excel, ['school', 'مدرسة'])
                    col_city = get_col(df_excel, ['city', 'سكن', 'مدينة'])
                    col_phone = get_col(df_excel, ['phone', 'جوال'])
                    col_role = get_col(df_excel, ['role', 'مهمة'])
                    col_hall = get_col(df_excel, ['hall', 'قاعة'])
                    col_accept = get_col(df_excel, ['accept', 'r'])

                    for _, row in df_excel.iterrows():
                        c.execute('''
                        INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, accept)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            str(row[col_id]) if col_id else "",
                            str(row[col_name]) if col_name else "",
                            str(row[col_school]) if col_school else "",
                            str(row[col_city]) if col_city else "",
                            str(row[col_phone]) if col_phone else "",
                            str(row[col_role]) if col_role else "",
                            str(row[col_hall]) if col_hall else "",
                            str(row[col_accept]) if col_accept else "نعم"
                        ))
                    conn.commit()
                    st.success("تم التحديث بنجاح! ✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ: {e}")

# --- التبويب الثالث: الإدارة ---
with tab_manage:
    st.subheader("الإجراءات الجماعية")
    col_a, col_b = st.columns(2)
    
    with col_a:
        ready = pd.read_sql("SELECT * FROM teachers WHERE hall IS NOT NULL AND hall != ''", conn)
        if st.button(f"📦 تحميل ({len(ready)}) كتاب في ملف ZIP"):
            if not ready.empty:
                zip_io = io.BytesIO()
                with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as zf:
                    for _, r in ready.iterrows():
                        d_io = create_docx(r)
                        zf.writestr(f"تكليف_{r['name']}.docx", d_io.getvalue())
                zip_io.seek(0)
                st.download_button("📥 بدء التحميل", data=zip_io, file_name="all_assignments.zip")
            else:
                st.error("لا يوجد معلمون معينون بقاعات.")

    with col_b:
        if st.button("🗑️ مسح جميع البيانات الحالية"):
            c.execute("DELETE FROM teachers")
            conn.commit()
            st.warning("تم إفراغ قاعدة البيانات.")
            st.rerun()
