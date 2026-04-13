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
# 2. إعدادات الحالة والتبديل (تم إضافة نظام التصحيح هنا)
# =====================================
if 'popover_counter' not in st.session_state:
    st.session_state.popover_counter = 0

if 'system_mode' not in st.session_state:
    st.session_state['system_mode'] = "tawjihi"

def switch_system(mode):
    st.session_state['system_mode'] = mode
    st.cache_data.clear()
    st.rerun()

# توزيع الإعدادات بناءً على النظام المختار
if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان الثانوية العامة"
elif st.session_state['system_mode'] == "correction":
    DB_NAME = "data_correction.db"
    TEMPLATE_NAME = "template_correction.docx"
    # استبدل الروابط أدناه بروابط الشيت الخاصة بالتصحيح عند توفرها
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=0&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=111&single=true&output=csv"
    PAGE_TITLE = "نظام تكليفات تصحيح الثانوية العامة"
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان التوظيف"

st.set_page_config(page_title=PAGE_TITLE, layout="wide", initial_sidebar_state="collapsed")

# تصميم الواجهة CSS
st.markdown("""
    <style>
        .custom-header {
            position: fixed; top: 0; left: 0; width: 100%;
            background-color: #1a1c23; color: white; text-align: center;
            padding: 15px 0; z-index: 999999; border-bottom: 2px solid #00ffcc;
            line-height: 1.5; direction: rtl; box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
        }
        .stApp { margin-top: 80px; }
        header {visibility: hidden;}
        .main, .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
        .user-box { background-color: #1a1c23; padding: 5px 15px; border-radius: 8px; border-right: 5px solid #00ffcc; display: inline-block; float: right; }
        div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect { direction: rtl !important; text-align: right !important; }
        .move-to-right { text-align: right !important; direction: rtl !important; display: block; width: 100%; color: white; }
    </style>
    <div class="custom-header">
        <div style="font-weight: bold; font-size: 1.2rem;">إعداد وتصميم : عوض نعمان ريده</div>
        <div style="font-size: 1rem; color: #00ffcc;">قسم الامتحانات - مديرية التربية والتعليم جنوب نابلس</div>
    </div>
    """, unsafe_allow_html=True)

# الاتصال بقاعدة البيانات وإعداد الجداول
conn = sqlite3.connect(DB_NAME, check_same_thread=False, timeout=30)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT,
             relative TEXT, relative_exam TEXT)''')

c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS logs 
             (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
conn.commit()

@st.cache_data(ttl=10)
def get_cached_teachers():
    return pd.read_sql("SELECT * FROM teachers", conn)

@st.cache_data(ttl=60)
def get_cached_halls():
    return pd.read_sql("SELECT * FROM halls", conn)

def add_log(action, details):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO logs (user, action, details, timestamp) VALUES (?, ?, ?, ?)", 
              (st.session_state.username, action, details, now))
    conn.commit()
    st.cache_data.clear()

# =====================================
# 3. وظائف معالجة الملفات والوورد
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9:
        phone_val = '0' + phone_val
    
    h_name_final = str(h_name) if h_name and str(h_name).lower() != 'nan' else "---"
    h_city_final = str(h_city) if h_city and str(h_city).lower() != 'nan' else "---"
        
    repls = {
        'ZNAME': str(row.get('name', '')),
        'ZID': str(row.get('id', '')),
        'ZPHONE': phone_val,
        'ZJOB': str(row.get('role', '') or '---'),
        'ZHALL': h_name_final,
        'ZLOC': h_city_final,
        'ZWORK': str(row.get('school', '')),
        'ZCITY': str(row.get('city', '')),
        'ZREL': str(row.get('relative', 'لا يوجد')),
        'ZRELEXAM': str(row.get('relative_exam', '---'))
    }

    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)
                        run.bold = True

    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in repls.items():
                        if k in p.text:
                            for run in p.runs:
                                if k in run.text:
                                    run.text = run.text.replace(k, v)
                                    run.bold = True
    return doc_obj

def generate_single_doc(row):
    if not os.path.exists(TEMPLATE_NAME):
        st.error(f"❌ ملف القالب '{TEMPLATE_NAME}' غير موجود")
        return None
    doc = Document(TEMPLATE_NAME)
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# =====================================
# 4. الواجهة الرئيسية وتبديل الأنظمة
# =====================================
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    st.markdown(f"""
        <div class="user-box">
            <span style="color: #bbb;">👤 الموظف الحالي:</span> 
            <strong style="color: white; font-size: 1.1rem;">{st.session_state.username}</strong>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("") 
    # شريط تبديل الأنظمة - تمت إضافة زر التصحيح
    btn_col1, btn_col2, btn_col3, btn_spacer = st.columns([1, 1, 1.2, 1])
    with btn_col1:
        if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
            switch_system("tawjihi")
    with btn_col2:
        if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
            switch_system("tawzif")
    with btn_col3:
        if st.button("🖊️ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="correction" else "secondary"):
            switch_system("correction")

with header_col2:
    if st.button("🚪 خروج", key="logout_btn", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

# التبويبات الرئيسية
tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإحصائيات", "📜 سجل العمليات"])

# --- تبويب البحث ---
with tab_search:
    st.markdown(f'<h2 class="move-to-right">إدارة الموظفين - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        
        for idx, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | التكليف: {row['hall'] or 'غير مكلف'}"):
                # عرض البيانات الأساسية
                st.write(f"المدرسة: {row['school']} | السكن: {row['city']}")
                
                # نموذج التعيين
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("اختر المركز/القاعة", [""] + list(hall_map.keys()), key=f"h_{row['id']}_{idx}")
                with c2:
                    roles_list = ["", "مصحح", "رئيس ديوان", "عضو ديوان", "مراقب تصحيح"] if st.session_state.system_mode == "correction" else ["", "رئيس قاعة", "مساعد رئيس", "مراقب", "آذن"]
                    sel_r = st.selectbox("المهمة", roles_list, key=f"r_{row['id']}_{idx}")
                
                if st.button("💾 حفظ", key=f"s_{row['id']}_{idx}"):
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                             (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                    conn.commit()
                    st.success("تم الحفظ")
                    st.rerun()

# --- تبويب رفع البيانات ---
with tab_upload:
    st.markdown(f"### تحديث بيانات {PAGE_TITLE}")
    up_tpl = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
    if up_tpl:
        with open(TEMPLATE_NAME, "wb") as f:
            f.write(up_tpl.getbuffer())
        st.success("تم تحديث القالب")
    
    if st.button("🔄 تحديث من Google Sheets الآن"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
            dft.to_sql('teachers_temp', conn, if_exists='replace', index=False)
            # تحديث مع الحفاظ على التكليفات الحالية
            c.execute("INSERT OR IGNORE INTO teachers (id, name) SELECT id, name FROM teachers_temp")
            c.execute("""UPDATE teachers SET 
                        name=(SELECT name FROM teachers_temp WHERE teachers_temp.id=teachers.id),
                        phone=(SELECT phone FROM teachers_temp WHERE teachers_temp.id=teachers.id),
                        school=(SELECT school FROM teachers_temp WHERE teachers_temp.id=teachers.id),
                        city=(SELECT city FROM teachers_temp WHERE teachers_temp.id=teachers.id),
                        current_job=(SELECT current_job FROM teachers_temp WHERE teachers_temp.id=teachers.id)
                        WHERE id IN (SELECT id FROM teachers_temp)""")
            
            dfh = pd.read_csv(HALLS_URL)
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            conn.commit()
            st.success("✅ تم تحديث البيانات بنجاح")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"خطأ: {e}")

# (بقية التبويبات تعمل بشكل تلقائي بناءً على DB_NAME المختار لكل نظام)
