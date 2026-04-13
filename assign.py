import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import time
from datetime import datetime
import copy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# =====================================
# 1. نظام تسجيل الدخول باستخدام Secrets
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
# 2. إعدادات الحالة والتبديل
# =====================================
if 'popover_counter' not in st.session_state:
    st.session_state.popover_counter = 0

if 'system_mode' not in st.session_state:
    st.session_state['system_mode'] = "tawjihi"

def switch_system(mode):
    st.session_state['system_mode'] = mode
    st.cache_data.clear()
    st.rerun()

# تخصيص الإعدادات بناءً على النظام المختار
if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام تكليفات الثانوية العامة"
elif st.session_state['system_mode'] == "tawzif":
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام تكليفات التوظيف"
else: # تصحيح الثانوية العامة (New)
    DB_NAME = "data_tasheeh.db"
    TEMPLATE_NAME = "template_tasheeh.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=0&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=1885970999&single=true&output=csv"
    PAGE_TITLE = "نظام تصحيح الثانوية العامة"

st.set_page_config(page_title=PAGE_TITLE, layout="wide", initial_sidebar_state="collapsed")

# CSS وتنسيق الواجهة
st.markdown("""
    <style>
        .custom-header {
            position: fixed; top: 0; left: 0; width: 100%; background-color: #1a1c23; color: white;
            text-align: center; padding: 15px 0; z-index: 999999; border-bottom: 2px solid #00ffcc;
            line-height: 1.5; direction: rtl; box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
        }
        .stApp { margin-top: 80px; }
        header {visibility: hidden;}
        .main, .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
        .move-to-right { text-align: right !important; direction: rtl !important; display: block; width: 100%; color: white; }
        [data-testid="stSidebar"] { display: none; }
    </style>
    <div class="custom-header">
        <div style="font-weight: bold; font-size: 1.2rem;">إعداد وتصميم : عوض نعمان ريده</div>
        <div style="font-size: 1rem; color: #00ffcc;">قسم الامتحانات - مديرية التربية والتعليم جنوب نابلس</div>
    </div>
    """, unsafe_allow_html=True)

# قاعدة البيانات
conn = sqlite3.connect(DB_NAME, check_same_thread=False, timeout=30)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT,
             relative TEXT, relative_exam TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
conn.commit()

@st.cache_data(ttl=10)
def get_cached_teachers():
    return pd.read_sql("SELECT * FROM teachers", conn)

@st.cache_data(ttl=60)
def get_cached_halls():
    return pd.read_sql("SELECT * FROM halls", conn)

def add_log(action, details):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO logs (user, action, details, timestamp) VALUES (?, ?, ?, ?)", (st.session_state.username, action, details, now))
    conn.commit()

# =====================================
# 3. معالجة الملفات (Docx)
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    repls = {
        'ZNAME': str(row.get('name', '')),
        'ZID': str(row.get('id', '')),
        'ZPHONE': str(row.get('phone', '')),
        'ZJOB': str(row.get('role', '') or '---'),
        'ZHALL': str(h_name) if h_name else "---",
        'ZLOC': str(h_city) if h_city else "---",
        'ZWORK': str(row.get('school', '')),
        'ZCITY': str(row.get('city', ''))
    }
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)
                        run.bold = True
    return doc_obj

# =====================================
# 4. واجهة التبديل والأزرار
# =====================================
header_col1, header_col2 = st.columns([4, 1])
with header_col1:
    btn_col1, btn_col2, btn_col3, btn_spacer = st.columns([1, 1, 1.2, 1])
    with btn_col1:
        if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
            switch_system("tawjihi")
    with btn_col2:
        if st.button("👨‍🏫 التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
            switch_system("tawzif")
    with btn_col3:
        if st.button("🖋️ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="tasheeh" else "secondary"):
            switch_system("tasheeh")

with header_col2:
    if st.button("🚪 خروج", key="logout_btn", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

# التبويبات الرئيسية (نفس الوظائف لجميع الأنظمة)
tab_search, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "📥 رفع البيانات", "📊 الإحصائيات", "📜 السجل"])

with tab_search:
    st.markdown(f'<h2 class="move-to-right">إدارة البيانات - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
    q = st.text_input("ابحث عن الاسم أو الهوية")
    if q:
        df_t = get_cached_teachers()
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for idx, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | {row['hall'] or 'غير محدد'}"):
                st.write(f"المدرسة: {row['school']} | السكن: {row['city']}")
                # (باقي كود التكليف والحفظ...)

with tab_upload:
    st.markdown("### تحديث البيانات والقوالب")
    up_tpl = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
    if up_tpl:
        with open(TEMPLATE_NAME, "wb") as f: f.write(up_tpl.getbuffer())
        st.success("تم التحديث")

    if st.button("🔄 تحديث من Google Sheets (الرابط المخصص)"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
            dft.columns = dft.columns.str.strip().str.lower()
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            
            dfh = pd.read_csv(HALLS_URL)
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            
            st.success("✅ تم تحديث البيانات بنجاح من الروابط المحددة لهذا النظام")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"خطأ: {e}")

# استكمال باقي أقسام tab_manage و tab_logs...
