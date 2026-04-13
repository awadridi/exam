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
# 1. نظام تسجيل الدخول
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
                        st.error("❌ اسم المستخدم غير معرف")
        return False
    return True

if not login():
    st.stop()

# =====================================
# 2. إعدادات الأنظمة (إضافة التصحيح)
# =====================================
if 'system_mode' not in st.session_state:
    st.session_state['system_mode'] = "tawjihi"
if 'popover_counter' not in st.session_state:
    st.session_state.popover_counter = 0

def switch_system(mode):
    st.session_state['system_mode'] = mode
    st.cache_data.clear()
    st.rerun()

if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان الثانوية العامة"
elif st.session_state['system_mode'] == "correction":
    DB_NAME = "data_correction.db"
    TEMPLATE_NAME = "template_correction.docx"
    # روابط افتراضية للتصحيح (يمكنك تغييرها لاحقاً)
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام تكليفات تصحيح الثانوية العامة"
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان التوظيف"

st.set_page_config(page_title=PAGE_TITLE, layout="wide")

# =====================================
# 3. قاعدة البيانات والوظائف
# =====================================
conn = sqlite3.connect(DB_NAME, check_same_thread=False)
c = conn.cursor()

def init_db():
    c.execute('''CREATE TABLE IF NOT EXISTS teachers 
                 (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
                 role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
                 preference TEXT, current_job TEXT, ability TEXT,
                 relative TEXT, relative_exam TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS logs 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
    conn.commit()

init_db()

@st.cache_data(ttl=5)
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

# --- وظائف معالجة المستندات (نفس منطقك الأصلي) ---
def process_doc(doc_obj, row, h_name, h_city):
    repls = {
        'ZNAME': str(row.get('name', '')), 'ZID': str(row.get('id', '')),
        'ZJOB': str(row.get('role', '') or '---'), 'ZHALL': str(h_name or '---'),
        'ZWORK': str(row.get('school', '')), 'ZCITY': str(row.get('city', ''))
    }
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text: run.text = run.text.replace(k, v)
    return doc_obj

# =====================================
# 4. واجهة المستخدم الرأسية
# =====================================
st.markdown(f'<h1 style="text-align:center;">{PAGE_TITLE}</h1>', unsafe_allow_html=True)
col_h1, col_h2, col_h3 = st.columns(3)
with col_h1:
    if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
        switch_system("tawjihi")
with col_h2:
    if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
        switch_system("tawzif")
with col_h3:
    if st.button("🖊️ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="correction" else "secondary"):
        switch_system("correction")

st.divider()

# =====================================
# 5. التبويبات (كاملة الوظائف)
# =====================================
tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث", "🤖 التوزيع التلقائي", "📥 الرفع", "📊 الإحصائيات", "📜 السجل"])

# --- تبويب البحث ---
with tab_search:
    q = st.text_input("ابحث بالاسم أو الهوية")
    df_t = get_cached_teachers()
    h_data = get_cached_halls()
    h_map = {r['hall_name']: r['city'] for _, r in h_data.iterrows()}

    if q:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].str.contains(q, na=False)]
        for idx, row in results.iterrows():
            with st.expander(f"👤 {row['name']} - {row['hall'] or 'غير مكلف'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(h_map.keys()), key=f"h_{row['id']}")
                with c2:
                    roles = ["مراقب", "رئيس قاعة"] if st.session_state.system_mode != "correction" else ["مصحح", "عضو ديوان"]
                    sel_r = st.selectbox("المهمة", roles, key=f"r_{row['id']}")
                if st.button("حفظ التكليف", key=f"btn_{row['id']}"):
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                             (sel_h, sel_r, h_map.get(sel_h, ""), st.session_state.username, row['id']))
                    conn.commit()
                    add_log("تعيين يدوي", f"تكليف {row['name']} في {sel_h}")
                    st.success("تم الحفظ")
                    st.rerun()

# --- تبويب التوزيع التلقائي (مستعاد) ---
with tab_auto:
    df_all = get_cached_teachers()
    df_pool = df_all[(df_all['hall'] == '') | (df_all['hall'].isna())]
    st.info(f"عدد المتاحين للتوزيع: {len(df_pool)}")
    
    target_h = st.selectbox("القاعة المستهدفة", [""] + list(get_cached_halls()['hall_name']))
    num = st.number_input("العدد المطلوب", min_value=0, max_value=len(df_pool), value=0)
    
    if st.button("🚀 بدء التوزيع") and target_h and num > 0:
        sample = df_pool.sample(n=num)
        h_city = h_map.get(target_h, "")
        for _, r in sample.iterrows():
            c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=?, updated_by='توزيع تلقائي' WHERE id=?", 
                     (target_h, h_city, r['id']))
        conn.commit()
        add_log("توزيع تلقائي", f"توزيع {num} موظف على {target_h}")
        st.success("تم التوزيع بنجاح")
        st.rerun()

# --- تبويب الإحصائيات (مستعاد) ---
with tab_manage:
    df_stat = get_cached_teachers()
    total = len(df_stat)
    assigned = len(df_stat[df_stat['hall'].str.len() > 0])
    
    m1, m2, m3 = st.columns(3)
    m1.metric("الإجمالي", total)
    m2.metric("المكلفين", assigned)
    m3.metric("المتبقي", total - assigned)
    
    st.write("### توزيع الموظفين حسب القاعات")
    if not df_stat.empty:
        st.bar_chart(df_stat['hall'].value_counts())

# --- تبويب سجل العمليات (مستعاد) ---
with tab_logs:
    st.write("### سجل العمليات الأخير")
    df_logs = pd.read_sql("SELECT * FROM logs ORDER BY id DESC LIMIT 50", conn)
    st.table(df_logs)

# --- تبويب الرفع ---
with tab_upload:
    if st.button("🔄 تحديث البيانات من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str})
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL)
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            add_log("تحديث بيانات", "تم استيراد بيانات جديدة من جوجل")
            st.success("تم التحديث")
            st.rerun()
        except Exception as e:
            st.error(f"فشل التحديث: {e}")
