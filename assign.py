import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import time
from datetime import datetime

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
# 2. إعدادات الواجهة وقاعدة البيانات
# =====================================
st.set_page_config(page_title="نظام تكليف المراقبة", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    .editor-info { color: #ffc107 !important; font-size: 0.9rem; font-weight: bold; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; color: #00ffcc !important; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect("data_system_v26.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT)''')

c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS logs 
             (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
conn.commit()

def add_log(action, details):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO logs (user, action, details, timestamp) VALUES (?, ?, ?, ?)", 
              (st.session_state.username, action, details, now))
    conn.commit()

TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 3. وظائف معالجة الملفات
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9: phone_val = '0' + phone_val
    repls = {
        '<NAME>': str(row.get('name', '')), '<ID>': str(row.get('id', '')), 
        '<PHONE>': phone_val, '<JOB>': str(row.get('role', '')), 
        '<HALL_NAME>': str(h_name), '<HALL_LOCATION>': str(h_city), 
        '<WORKPLACE>': str(row.get('school', '')), '<CITY>': str(row.get('city', ''))
    }
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text: run.text = run.text.replace(k, v)
    return doc_obj

def generate_bulk_word(df, h_name):
    if not os.path.exists("template.docx"): return None
    final_doc = Document("template.docx"); final_doc._body.clear_content()
    for idx, row in df.iterrows():
        temp_doc = Document("template.docx")
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        if idx > 0: final_doc.add_page_break()
        for element in temp_doc.element.body:
            if not element.tag.endswith('sectPr'): final_doc.element.body.append(element)
    out = io.BytesIO(); final_doc.save(out); out.seek(0)
    return out

# =====================================
# 4. الواجهة والتابات
# =====================================
tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs([
    "🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"
])

# --- تبويب البحث (مختصر للعرض) ---
with tab_search:
    st.subheader("إدارة الموظفين والتعيين اليدوي")
    # (هنا يوضع كود البحث كما في النسخ السابقة)

# --- تبويب التوزيع التلقائي (المعدل لإظهار الإحصائيات) ---
with tab_auto:
    st.subheader("🤖 التوزيع التلقائي للمراقبين")
    
    # 1. جلب البيانات المتاحة حالياً (المراقبين فقط)
    query_avail = """
        SELECT * FROM teachers 
        WHERE current_job = 'معلم' 
        AND preference = 'يرغب' 
        AND ability = 'يصلح' 
        AND (hall = '' OR hall IS NULL OR hall = 'nan')
    """
    df_avail = pd.read_sql(query_avail, conn)
    
    # 2. عرض إحصائيات المتاحين حسب المنطقة السكنية (طلبك)
    st.info("📊 أعداد المعلمين المتاحين للتوزيع حالياً (حسب السكن):")
    if not df_avail.empty:
        # حساب التوزيع السكني
        city_stats = df_avail['city'].value_counts().reset_index()
        city_stats.columns = ['المنطقة السكنية', 'عدد المعلمين المتاحين']
        
        # عرض الإحصائيات في أعمدة لتوفير المساحة
        cols = st.columns(4)
        for i, row in city_stats.iterrows():
            cols[i % 4].metric(row['المنطقة السكنية'], row['عدد المعلمين المتاحين'])
    else:
        st.warning("⚠️ لا يوجد معلمون متاحون حالياً تنطبق عليهم الشروط (معلم + يرغب + يصلح).")

    st.divider()
    
    # 3. أدوات التوزيع
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_hall = st.selectbox("اختر القاعة المستهدفة:", [""] + list(hall_map.keys()))
        selected_cities = st.multiselect("اختر مناطق السكن المطلوب السحب منها:", options=sorted(df_avail['city'].unique().tolist()) if not df_avail.empty else [])
    
    with col_a2:
        req_proctors = st.number_input("العدد المطلوب:", min_value=1, value=10)
        if st.button("🚀 تنفيذ التوزيع العشوائي"):
            if target_hall and selected_cities:
                pool = df_avail[df_avail['city'].isin(selected_cities)].sample(frac=1).reset_index(drop=True).head(req_proctors)
                for _, t in pool.iterrows():
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                              (target_hall, "مراقب", hall_map.get(target_hall, ""), st.session_state.username, t['id']))
                conn.commit()
                st.session_state.last_assigned_proctors = pool[['name', 'id', 'city', 'school']]
                st.success(f"✅ تم توزيع {len(pool)} مراقب بنجاح")
                st.rerun()

    if 'last_assigned_proctors' in st.session_state and st.session_state.last_assigned_proctors is not None:
        st.divider()
        st.markdown("### 📋 كشف الموزعين في القاعة المختارة:")
        st.dataframe(st.session_state.last_assigned_proctors, use_container_width=True)

# --- تبويب الإدارة (مع زر الحذف للمراقبين) ---
with tab_manage:
    st.subheader("📊 إدارة القاعات")
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != '' AND hall IS NOT NULL", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة لإدارتها:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            c1, c2 = st.columns(2)
            with c1:
                proctors_only = df_hall_details[df_hall_details['role'] == 'مراقب']
                if st.button(f"🗑️ حذف مراقبي قاعة {h_choice} فقط", type="secondary"):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE hall=? AND role='مراقب'", (st.session_state.username, h_choice))
                    conn.commit()
                    st.success("تم حذف المراقبين مع الإبقاء على الإدارة")
                    st.rerun()
            with c2:
                bulk_f = generate_bulk_word(df_hall_details, h_choice)
                if bulk_f: st.download_button("📥 تحميل كتب القاعة", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")
            
            st.dataframe(df_hall_details[['name', 'role', 'school', 'city']], use_container_width=True)

# (تكملة بقية التابات الرفع والسجل كما هي)
