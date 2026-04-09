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

if 'popover_counter' not in st.session_state:
    st.session_state['popover_counter'] = 0

# =====================================
# 2. إعدادات الواجهة وقاعدة البيانات
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .main, .stApp { 
        direction: rtl; 
        text-align: right; 
        background-color: #0e1117; 
    }
    div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect {
        direction: rtl !important;
        text-align: right !important;
    }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    .editor-info { color: #ffc107 !important; font-size: 0.9rem; font-weight: bold; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; color: #00ffcc !important; }
    .stat-card {
        flex: 1;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        min-width: 150px;
        border: 1px solid #333;
    }
    .stat-wants { border-top: 5px solid #28a745; background-color: #1a2e1f; }
    .stat-no-wants { border-top: 5px solid #dc3545; background-color: #2e1a1a; }
    
    .move-to-right {
        text-align: right !important;
        direction: rtl !important;
        display: block;
        width: 100%;
        color: white;
    }
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

@st.cache_data(ttl=60)
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

TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 3. وظائف معالجة الملفات
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
        'ZJOB': str(row.get('role', '')),
        'ZHALL': h_name_final,
        'ZLOC': h_city_final,
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
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

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
# 4. الواجهة الرئيسية (التعديلات المطلوبة)
# =====================================

# التعديل: الزر بقي في مكانه (العمود الثاني) والنص تم محاذاته لليسار
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    st.markdown(f"""
        <div style="background-color: #1a1c23; padding: 10px 15px; border-radius: 8px; border-left: 5px solid #00ffcc; text-align: left;">
            <strong style="color: white; font-size: 1.1rem;">{st.session_state.username}</strong>
            <span style="color: #bbb;"> :👤 الموظف الحالي</span> 
        </div>
    """, unsafe_allow_html=True)

with header_col2:
    if st.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

with tab_search:
    st.markdown('<h2 class="move-to-right">إدارة الموظفين</h2>', unsafe_allow_html=True)
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            display_phone = str(row['phone'])
            if display_phone.startswith('5') and len(display_phone) == 9:
                display_phone = '0' + display_phone

            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                # ... كود عرض البيانات (نفس كودك الأصلي) ...
                st.markdown(f"🆔 الهوية: {row.get('id', '---')} | 📱 الجوال: {display_phone}")
                
                with st.popover("📝 تعديل البيانات الأساسية", key=f"pop_{row['id']}_{st.session_state.popover_counter}"):
                    u_name = st.text_input("الاسم", value=row['name'], key=f"un_{row['id']}")
                    # (بقية حقول التعديل...)
                    if st.button("💾 تحديث وحفظ", key=f"save_base_{row['id']}"):
                        c.execute("UPDATE teachers SET name=?, phone=?, updated_by=? WHERE id=?", (u_name, display_phone, st.session_state.username, row['id']))
                        conn.commit(); add_log("تعديل بيانات", u_name); st.rerun()

                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0), key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        conn.commit(); add_log("حفظ تكليف", row['name']); st.success("✅ تم الحفظ"); time.sleep(0.5); st.rerun()
                    if row['hall']:
                        if st.button("📥 إنشاء الكتاب", key=f"gen_s_{row['id']}"):
                            f_word = generate_single_doc(row)
                            if f_word: st.download_button("📥 تحميل", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_s_{row['id']}")

with tab_auto:
    st.markdown('<h2 class="move-to-right">🤖 نظام التوزيع التلقائي</h2>', unsafe_allow_html=True)
    # (كود التوزيع التلقائي كما هو في ملفك)

with tab_upload:
    st.markdown('<h2 class="move-to-right">تحديث القالب والبيانات</h2>', unsafe_allow_html=True)
    # (كود رفع البيانات كما هو في ملفك)

with tab_manage:
    df_all_teachers = get_cached_teachers()
    st.divider()
    assigned_halls = sorted(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]['hall'].unique().tolist())
    
    if assigned_halls:
        h_choice = st.selectbox("اختر قاعة لعرض الكادر:", [""] + assigned_halls)
        if h_choice:
            df_hall_details = df_all_teachers[df_all_teachers['hall'] == h_choice].copy()
            st.markdown(f'<h4 class="move-to-right">📊 توزيع الكادر في قاعة: {h_choice}</h4>', unsafe_allow_html=True)
            
            # التعديل: إعادة المعداد (Counter)
            col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
            with col_stat1: st.metric("رئيس قاعة", len(df_hall_details[df_hall_details['role'] == "رئيس قاعة"]))
            with col_stat2: st.metric("مساعد رئيس", len(df_hall_details[df_hall_details['role'] == "مساعد رئيس قاعة"]))
            with col_stat3: st.metric("مراقب", len(df_hall_details[df_hall_details['role'] == "مراقب"]))
            with col_stat4: st.metric("آذن", len(df_hall_details[df_hall_details['role'] == "آذن"]))
            
            # (بقية كود عرض الجدول وتفريغ القاعة كما هو في ملفك)
            st.dataframe(df_hall_details[['name', 'role', 'school', 'phone']], use_container_width=True)

with tab_logs:
    st.markdown('<h2 class="move-to-right">📜 سجل العمليات</h2>', unsafe_allow_html=True)
    
    # التعديل: إعادة زر حذف السجل
    if st.button("🗑️ حذف السجل بالكامل", key="clear_all_logs_btn"):
        c.execute("DELETE FROM logs")
        conn.commit()
        st.success("تم حذف سجل العمليات بنجاح")
        st.rerun()

    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
