import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io

# =====================================
# 1. إعدادات قاعدة البيانات والتنسيق
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    .stButton>button { width: 100%; background-color: #28a745; color: white; border-radius: 8px; }
    div[data-testid="stExpander"] { background-color: #f9f9f9; border: 1px solid #eee; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect("exam_data_2026.db", check_same_thread=False)
c = conn.cursor()

# إنشاء الجداول
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة تعبئة القالب (Template)
# =====================================
def generate_from_template(row):
    try:
        # يفتح الملف المسمى template.docx الموجود في مشروعك
        doc = Document("template.docx")
        
        # استبدال العلامات داخل الفقرات
        for p in doc.paragraphs:
            if '<NAME>' in p.text: p.text = p.text.replace('<NAME>', str(row['name']))
            if '<ID>' in p.text: p.text = p.text.replace('<ID>', str(row['id']))
            if '<HALL_NAME>' in p.text: p.text = p.text.replace('<HALL_NAME>', str(row['hall']))
            if '<HALL_LOCATION>' in p.text: p.text = p.text.replace('<HALL_LOCATION>', str(row['hall_city']))
            if '<WORKPLACE>' in p.text: p.text = p.text.replace('<WORKPLACE>', str(row['school']))
            if '<CITY>' in p.text: p.text = p.text.replace('<CITY>', str(row['city']))
        
        # استبدال العلامات داخل الجداول (إذا كانت موجودة في جداول)
        for table in doc.tables:
            for row_obj in table.rows:
                for cell in row_obj.cells:
                    for p in cell.paragraphs:
                        if '<NAME>' in p.text: p.text = p.text.replace('<NAME>', str(row['name']))
                        if '<ID>' in p.text: p.text = p.text.replace('<ID>', str(row['id']))
                        if '<HALL_NAME>' in p.text: p.text = p.text.replace('<HALL_NAME>', str(row['hall']))

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        return None

# =====================================
# 3. الواجهة الرئيسية (Tabs)
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

# --- التبويب الأول: البحث والتعيين (التصميم الأول) ---
with tab_search:
    st.subheader("تعيين الموظفين على القاعات والمهام")
    
    # جلب بيانات القاعات والمهام
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {row['hall_name']: row['city'] for _, row in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["مراقب", "رئيس قاعة", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية...")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            with st.expander(f"👤 {row['name']} (القاعة الحالية: {row['hall'] if row['hall'] else 'لم تُعين'})"):
                c1, c2 = st.columns(2)
                with c1:
                    # ميزة اختيار الوظيفة والقاعة يدوياً
                    sel_hall = st.selectbox(f"اختر القاعة لـ {row['name']}", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"حدد الوظيفة", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                
                with c2:
                    st.write(f"المدرسة الأصلية: {row['school']}")
                    if st.button("✅ حفظ وتثبيت", key=f"btn_{row['id']}"):
                        h_city = hall_map.get(sel_hall, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", 
                                 (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ بنجاح!")
                        st.rerun()
                    
                    if row['hall']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button("📄 تحميل الكتاب الرسمي (Word)", 
                                             data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", 
                                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                             key=f"dl_{row['id']}")
                        else:
                            st.error("⚠️ ملف template.docx مفقود")

# --- التبويب الثاني: رفع الملفات ---
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.write("### 1. ملف المعلمين")
        f_t = st.file_uploader("ارفع ملف الموظفين (xlsx)", type="xlsx", key="upl_t")
        if f_t and st.button("تأكيد رفع المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), "", ""))
            conn.commit()
            st.success("تم الرفع")

    with col_u2:
        st.write("### 2. ملف القاعات")
        f_h = st.file_uploader("ارفع ملف القاعات (xlsx)", type="xlsx", key="upl_h")
        if f_h and st.button("تتبيث القاعات"):
            dfh = pd.read_excel(f_h)
            for _, r in dfh.iterrows():
                # الأعمدة: رقم القاعة، اسم القاعة، مدينة القاعة
                c.execute("INSERT OR REPLACE INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم رفع القاعات")

# --- التبويب الثالث: الإدارة ---
with tab_manage:
    if st.button("🗑️ إفراغ قاعدة البيانات بالكامل"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
