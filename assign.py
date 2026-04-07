import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os

# =====================================
# 1. إعدادات الواجهة (التصميم الاحترافي)
# =====================================
st.set_page_config(page_title="نظام التكليفات الذكي 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] {
        border: 1px solid #444; border-radius: 10px;
        background-color: #262730; margin-bottom: 10px;
    }
    div[data-testid="stExpander"] p, div[data-testid="stExpander"] label {
        color: #ffffff !important; font-weight: 500;
    }
    .stButton>button {
        width: 100%; background-color: #28a745; color: white;
        border-radius: 8px; font-weight: bold;
    }
    .info-box {
        padding: 10px; background-color: #1e1e1e; border-right: 5px solid #28a745;
        margin-bottom: 20px; color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# الاتصال بقاعدة البيانات
db_name = "exam_system_permanent.db"
conn = sqlite3.connect(db_name, check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة التعبئة (دعم JOB و Bold)
# =====================================
def generate_from_template(row):
    try:
        if not os.path.exists("template.docx"):
            return "error_no_file"
        doc = Document("template.docx")
        def replace_text(container, search_str, replace_str):
            for p in container.paragraphs:
                for run in p.runs:
                    if search_str in run.text:
                        run.text = run.text.replace(search_str, str(replace_str))
            for table in container.tables:
                for r in table.rows:
                    for cell in r.cells:
                        replace_text(cell, search_str, replace_str)

        data_map = {
            '<NAME>': row['name'], '<ID>': row['id'], '<JOB>': row['role'],
            '<HALL_NAME>': row['hall'], '<HALL_LOCATION>': row['hall_city'],
            '<WORKPLACE>': row['school'], '<CITY>': row['city']
        }
        for key, value in data_map.items():
            replace_text(doc, key, value if value else "")
        
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except: return None

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_backup = st.tabs(["🔍 البحث والتعيين", "📥 رفع الإكسل", "💾 النسخ الاحتياطي (هام)"])

with tab_search:
    st.subheader("تعيين الموظفين وإصدار الكتب")
    
    df_halls = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو رقم الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for i, row in results.iterrows():
            h_status = row['hall'] if row['hall'] else "غير محدد"
            r_status = row['role'] if row['role'] else "غير محدد"
            
            with st.expander(f"👤 {row['name']} | القاعة: {h_status} | الوظيفة: {r_status}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"القاعة لـ {row['id']}", hall_list, 
                                          index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"الوظيفة لـ {row['id']}", role_list, 
                                          index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                        h_city = hall_map.get(sel_hall, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                        conn.commit()
                        st.success("تم الحفظ!")
                        st.rerun()
                    
                    if row['hall'] and row['role']:
                        file_data = generate_from_template(row)
                        if file_data == "error_no_file":
                            st.error("ملف template.docx غير موجود")
                        elif file_data:
                            st.download_button(f"📥 تحميل تكليف {row['role']}", data=file_data, 
                                             file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

with tab_upload:
    st.markdown('<div class="info-box">ارفع ملفات الإكسل هنا لمرة واحدة فقط.</div>', unsafe_allow_html=True)
    cu1, cu2 = st.columns(2)
    with cu1:
        f_t = st.file_uploader("ملف الموظفين (xlsx)", key="u_t")
        if f_t and st.button("تأكيد رفع المعلمين"):
            df = pd.read_excel(f_t)
            for _, r in df.iterrows():
                c.execute("INSERT OR REPLACE INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                          (str(r.get('id','')), str(r.get('name','')), str(r.get('school','')), str(r.get('city','')), str(r.get('phone','')), "", "", ""))
            conn.commit()
            st.success("تم الرفع")

    with cu2:
        f_h = st.file_uploader("ملف القاعات (xlsx)", key="u_h")
        if f_h and st.button("تأكيد رفع القاعات"):
            dfh = pd.read_excel(f_h)
            c.execute("DELETE FROM halls")
            for _, r in dfh.iterrows():
                c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
            conn.commit()
            st.success("تم الرفع")

# --- التبويب الجديد لحماية بياناتك من الضياع ---
with tab_backup:
    st.subheader("⚙️ حماية البيانات من الضياع")
    st.write("بما أن الاستضافة السحابية قد تحذف البيانات عند التحديث، استخدم هذه الأدوات:")
    
    # 1. تصدير البيانات لملف إكسل واحد يحتوي كل تعييناتك
    if st.button("📥 تحميل نسخة احتياطية (Excel)"):
        all_data = pd.read_sql("SELECT * FROM teachers", conn)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            all_data.to_excel(writer, index=False)
        st.download_button("اضغط هنا لتحميل النسخة الاحتياطية", data=buffer, file_name="backup_data_2026.xlsx")

    st.divider()
    
    # 2. مسح شامل
    if st.button("🗑️ مسح قاعدة البيانات بالكامل"):
        c.execute("DELETE FROM teachers")
        c.execute("DELETE FROM halls")
        conn.commit()
        st.rerun()
