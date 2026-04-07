import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io

# =====================================
# 1. إعدادات الواجهة (التصميم والتنسيق)
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    
    /* تصميم المستطيلات (Expander) مع ضمان ظهور الخط الأبيض */
    div[data-testid="stExpander"] {
        border: 1px solid #444;
        border-radius: 10px;
        background-color: #262730;
        margin-bottom: 10px;
    }
    
    /* إجبار جميع النصوص داخل نتائج البحث على اللون الأبيض */
    div[data-testid="stExpander"] p, 
    div[data-testid="stExpander"] span,
    div[data-testid="stExpander"] label,
    div[data-testid="stExpander"] div {
        color: #ffffff !important;
        font-weight: 500;
    }

    /* تصميم الأزرار باللون الأخضر */
    .stButton>button {
        width: 100%;
        background-color: #28a745;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# الاتصال بقاعدة البيانات
conn = sqlite3.connect("final_system_v3.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة التعبئة (دعم JOB والحفاظ على Bold)
# =====================================
def generate_from_template(row):
    try:
        # تأكد أن اسم الملف في مجلد المشروع هو template.docx
        doc = Document("template.docx")
        
        def replace_text(container, search_str, replace_str):
            # الاستبدال في الفقرات مع الحفاظ على التنسيق
            for p in container.paragraphs:
                for run in p.runs:
                    if search_str in run.text:
                        run.text = run.text.replace(search_str, str(replace_str))
            
            # الاستبدال داخل الجداول
            for table in container.tables:
                for r in table.rows:
                    for cell in r.cells:
                        replace_text(cell, search_str, replace_str)

        # خريطة البيانات (ربط وسوم الورد بأعمدة قاعدة البيانات)
        data_map = {
            '<NAME>': row['name'],
            '<ID>': row['id'],
            '<JOB>': row['role'],        # الوسم الجديد الذي أضفته في ملف الورد
            '<HALL_NAME>': row['hall'],
            '<HALL_LOCATION>': row['hall_city'],
            '<WORKPLACE>': row['school'],
            '<CITY>': row['city']
        }

        for key, value in data_map.items():
            replace_text(doc, key, value if value else "")

        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except Exception:
        return None

# =====================================
# 3. الواجهة الرئيسية (تبويبات)
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("تعيين الموظفين وإصدار الكتب")
    
    # جلب بيانات القاعات
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(row['hall_name']): str(row['city']) for _, row in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    
    # قائمة الوظائف (ستظهر مكان <JOB>)
    role_list = ["رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            # عرض البيانات في المستطيل (الخط أبيض)
            with st.expander(f"👤 الموظف: {row['name']} | القاعة: {row['hall'] if row['hall'] else 'لم تُحدد'} | الوظيفة: {row['role'] if row['role'] else 'لم تُحدد'}"):
                c1, c2 = st.columns(2)
                with c1:
                    # اختيار القاعة
                    sel_hall = st.selectbox(f"اختر القاعة لـ {row['name']}", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, 
                                          key=f"h_{row['id']}")
                    
                    # اختيار الوظيفة (JOB)
                    sel_role = st.selectbox(f"حدد الوظيفة (تظهر في الورد)", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, 
                                          key=f"r_{row['id']}")
                
                with c2:
                    st.write(f"المدرسة الأصلية: {row['school']}")
                    if st.button("✅ حفظ البيانات", key=f"btn_{row['id']}"):
                        h_city = hall_map.get(sel_hall, "")
                        # تحديث القاعدة (القاعة + الوظيفة + مدينة القاعة)
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", 
                                 (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ بنجاح!")
                        st.rerun()
                    
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data:
                            st.download_button("📥 تحميل الكتاب الرسمي", 
                                             data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", 
                                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                             key=f"dl_{row['id']}")

# تبويب رفع الملفات
with tab_upload:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.write("### 1. ملف المعلمين")
        f_t = st.file_uploader("ارفع ملف الموظفين (xlsx)", type="xlsx", key="upl_t")
        if f_t and st.button("رفع المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), str(r.get('role','')), "", ""))
            conn.commit()
            st.success("تم الرفع بنجاح")

    with col_u2:
        st.write("### 2. ملف القاعات")
        f_h = st.file_uploader("ارفع ملف القاعات (xlsx)", type="xlsx", key="upl_h")
        if f_h and st.button("رفع القاعات"):
            dfh = pd.read_excel(f_h)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم رفع القاعات بنجاح")

# تبويب الإدارة
with tab_manage:
    if st.button("🗑️ مسح جميع البيانات"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
