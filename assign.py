import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
from copy import deepcopy

# =====================================
# 1. نظام تسجيل الدخول (حسب كودك الأصلي)
# =====================================
def login():
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
        st.session_state['username'] = ""

    if not st.session_state['logged_in']:
        st.markdown("<h2 style='text-align: center;'>🔐 نظام تكليفات المكتب - دخول</h2>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("login_form"):
                user = st.text_input("اسم المستخدم").lower().strip()
                pw = st.text_input("كلمة المرور", type="password").strip()
                submit = st.form_submit_button("دخول")
                if submit:
                    try:
                        valid_password = st.secrets[f"password_{user}"]
                        if pw == valid_password:
                            st.session_state['logged_in'] = True
                            st.session_state['username'] = user
                            st.rerun()
                        else:
                            st.error("❌ كلمة المرور غير صحيحة")
                    except KeyError:
                        st.error("❌ اسم المستخدم غير معرف في Secrets")
        return False
    return True

if not login():
    st.stop()

# =====================================
# 2. إعدادات الواجهة والستايل (CSS)
# =====================================
st.set_page_config(page_title="نظام تكليف المراقبة", layout="wide")

# شريط جانبي مع معلومات المستخدم وزر الخروج
with st.sidebar:
    st.markdown(f"### 👤 المستخدم: {st.session_state.username}")
    if st.button("🚪 تسجيل الخروج"):
        st.session_state['logged_in'] = False
        st.session_state['username'] = ""
        st.rerun()
    st.markdown("---")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0b0e14; }
    .employee-card {
        background-color: #161b22;
        padding: 20px;
        border-radius: 6px;
        border-right: 5px solid #00f2ea;
        margin-bottom: 15px;
        border: 1px solid #30363d;
    }
    .card-title { color: #f0f6fc; font-size: 1.2rem; font-weight: bold; margin-bottom: 10px; }
    .label { color: #8b949e; font-size: 0.9rem; }
    .value { color: #58a6ff; font-weight: bold; margin-left: 15px; }
    .validity-alert { color: #d29922; font-weight: bold; margin-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 3. قاعدة البيانات والوظائف التقنية
# =====================================
conn = sqlite3.connect("data_system_v26.db", check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
conn.commit()

def fill_template(row):
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    # استبدال النصوص بناءً على القالب الخاص بك
    replacements = {
        "سارة سامر محمود عابد": str(row['name']),
        "404493474": str(row['id']),
        "مراقب": str(row['role'] or "مراقب"),
        "عبد الرحيم عودة محمود": str(row['hall'] or ""),
        "يتما": str(row['hall_city'] or ""),
        "الراشد الاساسية للبنبن": str(row['school'] or ""),
        "قبلان": str(row['city'] or "")
    }
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text: p.text = p.text.replace(k, v)
    for table in doc.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in replacements.items():
                        if k in p.text: p.text = p.text.replace(k, v)
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# =====================================
# 4. التبويبات الرئيسية
# =====================================
tab_search, tab_auto, tab_manage, tab_upload = st.tabs([
    "🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📊 الإدارة", "📥 الرفع"
])

with tab_search:
    st.markdown("### 🔍 إدارة الموظفين والتعيين اليدوي")
    q = st.text_input("ابحث بالاسم أو الهوية:")
    
    h_df = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in h_df.iterrows()}

    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            st.markdown(f"""
            <div class="employee-card">
                <div class="card-title">👤 {row['name']} | القاعة: {row['hall'] or '---'}</div>
                <div style="display: flex; flex-wrap: wrap; gap: 20px;">
                    <div><span class="label">🆔 الهوية:</span> <span class="value">{row['id']}</span></div>
                    <div><span class="label">📱 الجوال:</span> <span class="value">{row['phone']}</span></div>
                    <div><span class="label">🏫 المدرسة:</span> <span class="value">{row['school']}</span></div>
                </div>
                <div class="validity-alert">⚠️ صلاحية المراقبة: {row['ability'] or 'يصلح'}</div>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                new_h = st.selectbox("تعيين لقاعة:", [""] + list(hall_map.keys()), key=f"h_{row['id']}")
            with c2:
                new_r = st.selectbox("المهمة:", ["مراقب", "رئيس قاعة", "مساعد"], key=f"r_{row['id']}")
            with c3:
                if st.button("💾 حفظ", key=f"sv_{row['id']}", use_container_width=True):
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                             (new_h, new_r, hall_map.get(new_h, ""), st.session_state.username, row['id']))
                    conn.commit(); st.rerun()
            
            if row['hall']:
                doc_file = fill_template(row)
                if doc_file:
                    st.download_button(f"📥 تحميل تكليف {row['name']}", data=doc_file, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")
            st.divider()

with tab_auto:
    st.markdown("### 🤖 التوزيع التلقائي الذكي")
    # منطق التوزيع التلقائي (كما في كودك الأصلي)
    df_avail = pd.read_sql("SELECT * FROM teachers WHERE (hall = '' OR hall IS NULL)", conn)
    if not df_avail.empty:
        target_h = st.selectbox("القاعة المستهدفة:", list(hall_map.keys()))
        count = st.number_input("العدد:", min_value=1, max_value=len(df_avail), value=1)
        if st.button("🚀 بدء التوزيع"):
            pool = df_avail.sample(n=int(count))
            for _, r in pool.iterrows():
                c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=? WHERE id=?", (target_h, hall_map[target_h], r['id']))
            conn.commit(); st.success("تم التوزيع"); st.rerun()

with tab_manage:
    st.markdown("### 📊 إحصائيات القاعات")
    df_all = pd.read_sql("SELECT * FROM teachers WHERE hall != ''", conn)
    st.dataframe(df_all[['name', 'hall', 'role', 'school']], use_container_width=True)

with tab_upload:
    st.markdown("### 📥 رفع البيانات")
    up_docx = st.file_uploader("ارفع قالب الورد (template.docx):", type="docx")
    if up_docx:
        with open("template.docx", "wb") as f: f.write(up_docx.getbuffer())
        st.success("تم تحديث القالب")
