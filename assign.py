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

# إعادة إنشاء الجداول لضمان مطابقتها للكود
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, accept TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# دالة ذكية للبحث عن الأعمدة
def find_column(df, keywords):
    for col in df.columns:
        if any(key in str(col).lower() for key in keywords):
            return col
    return None

# =====================================
# 3. دالة توليد Word
# =====================================
def create_docx(row):
    doc = Document()
    doc.add_paragraph("دولة فلسطين\nوزارة التربية والتعليم").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading('بطاقة تكليف رسمي', level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    content = f"""
    الاسم: {row['name']}
    رقم الهوية: {row['id']}
    المهمة: {row['role']}
    مكان العمل (القاعة): {row['hall']}
    المنطقة/المدينة: {row['city']}
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

# --- التبويب الأول: البحث والتعيين ---
with tab_search:
    st.subheader("إدارة التعيينات")
    
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_list = [""] + (df_halls['hall_name'].tolist() if not df_halls.empty else [])
    role_list = ["مراقب", "رئيس قاعة", "مساعد رئيس قاعة", "آذن", "عضو"]

    search_q = st.text_input("ابحث عن اسم الموظف أو الهوية...")
    df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
    
    if search_q and not df_teachers.empty:
        results = df_teachers[df_teachers['name'].str.contains(search_q, na=False) | df_teachers['id'].astype(str).str.contains(search_q)]
        
        for i, row in results.iterrows():
            with st.expander(f"👤 {row['name']} (القاعة الحالية: {row['hall'] if row['hall'] else 'غير معين'})"):
                col1, col2 = st.columns(2)
                with col1:
                    new_hall = st.selectbox(f"اختر القاعة", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, 
                                          key=f"h_{row['id']}")
                    new_role = st.selectbox(f"المهمة", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, 
                                          key=f"r_{row['id']}")
                
                with col2:
                    st.info(f"المدرسة: {row['school']}")
                    if st.button("💾 حفظ التعديلات", key=f"sv_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (new_hall, new_role, row['id']))
                        conn.commit()
                        st.success("تم الحفظ")
                        st.rerun()
                    
                    if row['hall']:
                        doc = create_docx(row)
                        st.download_button("📥 تحميل الكتاب", data=doc, file_name=f"تكليف_{row['id']}.docx", key=f"dl_{row['id']}")

# --- التبويب الثاني: رفع الملفات ---
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        st.subheader("1. ملف المعلمين")
        file_t = st.file_uploader("ارفع ملف الموظفين (xlsx)", type="xlsx", key="t_file")
        if file_t and st.button("تثبيت الموظفين"):
            df = pd.read_excel(file_t)
            c_id = find_column(df, ['id', 'هوية'])
            c_name = find_column(df, ['name', 'اسم'])
            c_school = find_column(df, ['school', 'مدرسة'])
            
            if c_id and c_name:
                for _, r in df.iterrows():
                    c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, accept) VALUES (?,?,?,?,?,?,?,?)",
                              (str(r[c_id]), str(r[c_name]), str(r.get(c_school, '')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), str(r.get('hall','')), 'نعم'))
                conn.commit()
                st.success("✅ تم تحديث المعلمين")
            else:
                st.error("تأكد من وجود عمود 'id' و 'name' في الملف")

    with col_u2:
        st.subheader("2. ملف القاعات")
        file_h = st.file_uploader("ارفع ملف القاعات (xlsx)", type="xlsx", key="h_file")
        if file_h and st.button("تثبيت القاعات"):
            dfh = pd.read_excel(file_h)
            c_num = find_column(dfh, ['number', 'رقم'])
            c_hall = find_column(dfh, ['hall', 'قاعة', 'اسم'])
            c_city = find_column(dfh, ['city', 'بلد', 'سكن'])
            
            if c_num and c_hall:
                for _, r in dfh.iterrows():
                    # تأكد من تحويل القيم لنصوص لتفادي أخطاء SQLite
                    val_num = str(r[c_num])
                    val_name = str(r[c_hall])
                    val_city = str(r[c_city]) if c_city else ""
                    
                    c.execute("INSERT OR REPLACE INTO halls (number, hall_name, city) VALUES (?,?,?)",
                              (val_num, val_name, val_city))
                conn.commit()
                st.success("✅ تم تحديث القاعات")
                st.rerun()
            else:
                st.error("تأكد من وجود عمود 'number' و 'hall' في الملف")

# --- التبويب الثالث: الإدارة ---
with tab_manage:
    if st.button("🗑️ مسح النظام بالكامل"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
