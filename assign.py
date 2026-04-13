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
# 2. إعدادات الأنظمة والتبديل
# =====================================
if 'system_mode' not in st.session_state:
    st.session_state['system_mode'] = "tawjihi"

def switch_system(mode):
    st.session_state['system_mode'] = mode
    st.cache_data.clear()
    st.rerun()

if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات - الثانوية العامة"
elif st.session_state['system_mode'] == "correction":
    DB_NAME = "data_correction.db"
    TEMPLATE_NAME = "template_correction.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=0&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=111&single=true&output=csv"
    PAGE_TITLE = "نظام تكليفات - تصحيح الثانوية"
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات - امتحان التوظيف"

st.set_page_config(page_title=PAGE_TITLE, layout="wide")

# إجبار التنسيق من اليمين لليسار (RTL)
st.markdown("""
    <style>
    .main, .stApp { direction: rtl; text-align: right; }
    div[data-testid="stSidebar"] { direction: rtl; }
    .stTabs [data-baseweb="tab-list"] { direction: rtl; }
    </style>
    """, unsafe_allow_html=True)

# =====================================
# 3. قاعدة البيانات والعمليات
# =====================================
conn = sqlite3.connect(DB_NAME, check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT,
             relative TEXT, relative_exam TEXT, subject TEXT)''')
# إضافة عمود المبحث إذا لم يكن موجوداً
try: c.execute("ALTER TABLE teachers ADD COLUMN subject TEXT")
except: pass

c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS logs 
             (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
conn.commit()

@st.cache_data(ttl=5)
def get_cached_teachers():
    return pd.read_sql("SELECT * FROM teachers", conn)

def add_log(action, details):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO logs (user, action, details, timestamp) VALUES (?, ?, ?, ?)", 
              (st.session_state.username, action, details, now))
    conn.commit()

# =====================================
# 4. الواجهة الرأسية
# =====================================
st.markdown(f'<h2 style="text-align:right;">{PAGE_TITLE}</h2>', unsafe_allow_html=True)
c_btn1, c_btn2, c_btn3 = st.columns(3)
with c_btn1:
    if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
        switch_system("tawjihi")
with c_btn2:
    if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
        switch_system("tawzif")
with c_btn3:
    if st.button("🖊️ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="correction" else "secondary"):
        switch_system("correction")

st.divider()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإحصائيات", "📜 سجل العمليات"])

# --- البحث ---
with tab_search:
    q = st.text_input("ابحث عن الاسم أو الهوية")
    if q:
        df = get_cached_teachers()
        results = df[df['name'].str.contains(q, na=False) | df['id'].str.contains(q, na=False)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']}"):
                st.write(f"المبحث: {row['subject'] or 'غير محدد'}")
                # نموذج الحفظ اليدوي...

# --- التوزيع التلقائي (تعديل المبحث) ---
with tab_auto:
    st.markdown("### التوزيع التلقائي حسب المبحث")
    df_auto = get_cached_teachers()
    # جلب المباحث الفريدة المتوفرة في قاعدة البيانات
    available_subjects = df_auto['subject'].dropna().unique().tolist()
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_h = st.selectbox("القاعة المستهدفة", [""] + pd.read_sql("SELECT hall_name FROM halls", conn)['hall_name'].tolist())
    with col_a2:
        selected_sub = st.selectbox("اختر المبحث المراد توزيعه", [""] + available_subjects)

    if st.button("🚀 بدء التوزيع بناءً على المبحث"):
        if target_h and selected_sub:
            # تصفية المعلمين حسب المبحث والذين ليس لديهم تكليف حالي
            pool = df_auto[(df_auto['subject'] == selected_sub) & ((df_auto['hall'] == '') | (df_auto['hall'].isna()))]
            if not pool.empty:
                for _, r in pool.iterrows():
                    c.execute("UPDATE teachers SET hall=?, role='مكلف', updated_by='توزيع تلقائي' WHERE id=?", (target_h, r['id']))
                conn.commit()
                add_log("توزيع تلقائي", f"توزيع مبحث {selected_sub} على قاعة {target_h}")
                st.success(f"تم توزيع جميع معلمي {selected_sub} بنجاح!")
                st.rerun()
            else:
                st.warning("لا يوجد معلمون غير مكلفين لهذا المبحث")

# --- رفع البيانات (استعادة الوورد) ---
with tab_upload:
    st.markdown("### رفع القوالب والبيانات")
    up_docx = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
    if up_docx:
        with open(TEMPLATE_NAME, "wb") as f:
            f.write(up_docx.getbuffer())
        st.success("تم تحديث ملف القالب بنجاح")
    
    st.divider()
    if st.button("🔄 تحديث من Google Sheets"):
        # منطق الاستيراد...
        st.info("جاري التحديث...")

# --- سجل العمليات (إضافة الحذف) ---
with tab_logs:
    col_l1, col_l2 = st.columns([4, 1])
    with col_l1:
        st.markdown("### سجل العمليات الأخير")
    with col_l2:
        if st.button("🗑️ مسح السجل", type="secondary"):
            c.execute("DELETE FROM logs")
            conn.commit()
            st.success("تم مسح السجل")
            st.rerun()
    
    df_logs = pd.read_sql("SELECT * FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_logs, use_container_width=True)

# --- الإحصائيات ---
with tab_manage:
    df_stat = get_cached_teachers()
    if not df_stat.empty:
        st.metric("إجمالي الموظفين", len(df_stat))
        st.write("التوزيع حسب المبحث:")
        st.bar_chart(df_stat['subject'].value_counts())
