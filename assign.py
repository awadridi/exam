import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io

# =====================================
# 1. إعدادات الواجهة والتنسيق (حل مشكلة اللون الأبيض)
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    /* تعديل لون الخط داخل المستطيلات ليظهر بوضوح */
    div[data-testid="stExpander"] div[role="button"] p { color: #000000 !important; font-weight: bold; }
    div[data-testid="stExpander"] div[data-testid="stMarkdownContainer"] p { color: #000000 !important; }
    .stButton>button { width: 100%; background-color: #28a745; color: white; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

# الاتصال بقاعدة البيانات
conn = sqlite3.connect("exam_final_2026.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة التعبئة مع الحفاظ على التنسيق (Bold)
# =====================================
def generate_from_template(row):
    try:
        doc = Document("template.docx")
        
        # دالة داخلية للبحث والاستبدال داخل الـ Runs للحفاظ على التنسيق (Bold, Font size, etc.)
        def replace_text_preserve_format(container, search_str, replace_str):
            for p in container.paragraphs:
                for run in p.runs:
                    if search_str in run.text:
                        run.text = run.text.replace(search_str, str(replace_str))
            
            # البحث أيضاً داخل الجداول إذا وجدت
            for table in container.tables:
                for r in table.rows:
                    for cell in r.cells:
                        replace_text_preserve_format(cell, search_str, replace_str)

        # تنفيذ الاستبدال لكل العلامات
        data_map = {
            '<NAME>': row['name'],
            '<ID>': row['id'],
            '<HALL_NAME>': row['hall'],
            '<HALL_LOCATION>': row['hall_city'],
            '<WORKPLACE>': row['school'],
            '<CITY>': row['city']
        }

        for key, value in data_map.items():
            replace_text_preserve_format(doc, key, value if value else "")

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception as e:
        return None

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("تعيين الموظفين")
    
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {row['hall_name']: row['city'] for _, row in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["مراقب", "رئيس قاعة", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            # المستطيل الذي كان لونه أبيض (Expander)
            with st.expander(f"👤 الموظف: {row['name']} | التكليف الحالي: {row['hall'] if row['hall'] else 'لم يحدد'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"اختر القاعة", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"حدد الوظيفة", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("✅ حفظ التثبيت", key=f"btn_{row['id']}"):
                        h_city = hall_map.get(sel_hall, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", 
                                 (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ!")
                        st.rerun()
                    
                    if row['hall']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button("📥 تحميل الكتاب الرسمي", 
                                             data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", 
                                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                             key=f"dl_{row['id']}")

# التبويبات الأخرى (الرفع والإدارة) بقيت كما هي لضمان عدم حذف أي ميزة
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.write("### 1. ملف المعلمين")
        f_t = st.file_uploader("xlsx", key="upl_t")
        if f_t and st.button("تأكيد المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), "", ""))
            conn.commit()
            st.success("تم")

    with col_u2:
        st.write("### 2. ملف القاعات")
        f_h = st.file_uploader("xlsx ", key="upl_h")
        if f_h and st.button("تأكيد القاعات"):
            dfh = pd.read_excel(f_h)
            for _, r in dfh.iterrows():
                c.execute("INSERT OR REPLACE INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم")

with tab_manage:
    if st.button("🗑️ مسح الكل"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
