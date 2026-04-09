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

# إنشاء الجداول إذا لم تكن موجودة
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
st.sidebar.markdown(f"### 👤 الموظف: **{st.session_state.username}**")
if st.sidebar.button("🚪 خروج"):
    st.session_state.logged_in = False
    st.rerun()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs([
    "🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"
])

# --- تبويب البحث ---
with tab_search:
    st.subheader("إدارة الموظفين والتعيين اليدوي")
    # (كود البحث والتعيين اليدوي يبقى كما هو في النسخ المستقرة)

# --- تبويب التوزيع التلقائي (المعدل: الإحصائيات في الأسفل داخل Expander) ---
with tab_auto:
    st.subheader("🤖 التوزيع التلقائي للمراقبين")
    
    # تحضير البيانات
    df_h_data_auto = pd.read_sql("SELECT * FROM halls", conn)
    hall_map_auto = {r['hall_name']: r['city'] for _, r in df_h_data_auto.iterrows()}
    
    query_avail = """
        SELECT * FROM teachers 
        WHERE current_job = 'معلم' 
        AND preference = 'يرغب' 
        AND ability = 'يصلح' 
        AND (hall = '' OR hall IS NULL OR hall = 'nan')
    """
    df_avail = pd.read_sql(query_avail, conn)

    # 1. قسم أدوات التوزيع (في الأعلى)
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_hall = st.selectbox("اختر القاعة المستهدفة للتوزيع:", [""] + list(hall_map_auto.keys()), key="auto_h")
        selected_cities = st.multiselect("اختر مناطق سكن المعلمين:", options=sorted(df_avail['city'].unique().tolist()) if not df_avail.empty else [])
    
    with col_a2:
        req_proctors = st.number_input("العدد المطلوب من المراقبين:", min_value=1, value=10)
        if st.button("🚀 تنفيذ التوزيع العشوائي", use_container_width=True):
            if target_hall and selected_cities:
                pool = df_avail[df_avail['city'].isin(selected_cities)].sample(frac=1).reset_index(drop=True).head(req_proctors)
                if not pool.empty:
                    for _, t in pool.iterrows():
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                                  (target_hall, "مراقب", hall_map_auto.get(target_hall, ""), st.session_state.username, t['id']))
                    conn.commit()
                    st.session_state.last_assigned_proctors = pool[['name', 'id', 'city', 'school']]
                    add_log("توزيع تلقائي", f"توزيع {len(pool)} مراقب على قاعة {target_hall}")
                    st.success(f"✅ تم توزيع {len(pool)} مراقب بنجاح")
                    st.rerun()
                else:
                    st.error("❌ لا يوجد موظفون متاحون في المناطق المختارة")

    # 2. عرض الموزعين حالياً (بعد الضغط على الزر)
    if 'last_assigned_proctors' in st.session_state and st.session_state.last_assigned_proctors is not None:
        st.divider()
        st.markdown(f"### 📋 الموزعون حالياً في قاعة: {target_hall}")
        st.dataframe(st.session_state.last_assigned_proctors, use_container_width=True, hide_index=True)
        if st.button("🧹 مسح هذا الكشف"):
            st.session_state.last_assigned_proctors = None
            st.rerun()

    st.divider()

    # 3. إحصائيات المتاحين حسب السكن (في الأسفل داخل Expander كما طلبت)
    with st.expander("📊 عرض أعداد المعلمين المتاحين في كل منطقة سكنية"):
        if not df_avail.empty:
            city_stats = df_avail['city'].value_counts().reset_index()
            city_stats.columns = ['المنطقة السكنية', 'عدد المتاحين']
            
            # عرضها بشكل مربعات صغيرة (Metrics)
            cols = st.columns(4)
            for i, row in city_stats.iterrows():
                cols[i % 4].metric(row['المنطقة السكنية'], f"{row['عدد المتاحين']} معلم")
        else:
            st.write("لا يوجد معلمون متاحون حالياً وفق الشروط.")

# --- تبويب الإدارة (مع ميزة حذف المراقبين فقط) ---
with tab_manage:
    st.subheader("📊 إدارة القاعات المنجزة")
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != '' AND hall IS NOT NULL", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض والإدارة:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            c_man1, c_man2 = st.columns(2)
            with c_man1:
                # زر حذف المراقبين فقط (يستثني الإدارة والآذنة)
                proctors_count = len(df_hall_details[df_hall_details['role'] == 'مراقب'])
                if st.button(f"🗑️ حذف تكليف ({proctors_count}) مراقب من {h_choice}", type="secondary", use_container_width=True):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE hall=? AND role='مراقب'", (st.session_state.username, h_choice))
                    conn.commit()
                    add_log("حذف مراقبين", f"إزالة مراقبي قاعة {h_choice}")
                    st.success("تم الحذف بنجاح")
                    st.rerun()
            with c_man2:
                bulk_f = generate_bulk_word(df_hall_details, h_choice)
                if bulk_f: st.download_button(f"📄 تحميل كتب قاعة {h_choice}", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx", use_container_width=True)
            
            st.dataframe(df_hall_details[['name', 'role', 'school', 'city']], use_container_width=True)

# (باقي التابات: الرفع والسجلات تبقى كما هي)
