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
    .user-box {
        background-color: #1a1c23;
        padding: 5px 15px;
        border-radius: 8px;
        border-right: 5px solid #00ffcc;
        display: inline-block;
        float: right;
    }
    .stat-box {
        background-color: #1e2130;
        padding: 10px;
        border-radius: 10px;
        text-align: center;
        border: 1px solid #333;
        color: white;
    }
    .stat-number {
        font-size: 1.5rem;
        font-weight: bold;
        color: #00ffcc;
    }
    div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect {
        direction: rtl !important;
        text-align: right !important;
    }
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
# 4. الواجهة الرئيسية
# =====================================

header_col1, header_col2 = st.columns([4, 1])
with header_col1:
    st.markdown(f'<div class="user-box"><span style="color: #bbb;">👤 الموظف:</span> <strong style="color: white;">{st.session_state.username}</strong></div>', unsafe_allow_html=True)
with header_col2:
    if st.button("🚪 خروج", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

# (تبويب البحث والتوزيع ورفع البيانات تبقى كما هي في كودك الأصلي)
with tab_search:
    st.markdown('<h2 class="move-to-right">إدارة الموظفين</h2>', unsafe_allow_html=True)
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | {row['hall'] or 'غير مكلف'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"h_{row['id']}")
                with c2:
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0), key=f"r_{row['id']}")
                if st.button("💾 حفظ", key=f"sv_{row['id']}"):
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                    conn.commit()
                    st.success("تم الحفظ")
                    st.rerun()

with tab_auto:
    st.markdown('<h2 class="move-to-right">🤖 التوزيع التلقائي</h2>', unsafe_allow_html=True)
    # كود التوزيع التلقائي الأصلي...

with tab_upload:
    st.markdown('<h2 class="move-to-right">تحديث البيانات</h2>', unsafe_allow_html=True)
    # كود الرفع الأصلي...

# =====================================
# الجزء المعاد (المعداد في تبويب الإدارة)
# =====================================
with tab_manage:
    df_all_teachers = get_cached_teachers()
    st.markdown('<h3 class="move-to-right">📊 إدارة القاعات والتقارير</h3>', unsafe_allow_html=True)
    
    assigned_halls = sorted(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]['hall'].unique().tolist())
    
    if assigned_halls:
        h_choice = st.selectbox("اختر قاعة لعرض الكادر:", [""] + assigned_halls)
        if h_choice:
            df_hall_details = df_all_teachers[df_all_teachers['hall'] == h_choice].copy()
            
            # --- المعداد (استعادة الجزء المحذوف) ---
            st.markdown(f'<h4 class="move-to-right">🔢 معداد قاعة: {h_choice}</h4>', unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                count_boss = len(df_hall_details[df_hall_details['role'] == "رئيس قاعة"])
                st.markdown(f'<div class="stat-box">رئيس قاعة<br><span class="stat-number">{count_boss}</span></div>', unsafe_allow_html=True)
            with m2:
                count_asst = len(df_hall_details[df_hall_details['role'] == "مساعد رئيس قاعة"])
                st.markdown(f'<div class="stat-box">مساعد رئيس<br><span class="stat-number">{count_asst}</span></div>', unsafe_allow_html=True)
            with m3:
                count_obs = len(df_hall_details[df_hall_details['role'] == "مراقب"])
                st.markdown(f'<div class="stat-box">مراقب<br><span class="stat-number">{count_obs}</span></div>', unsafe_allow_html=True)
            with m4:
                count_serv = len(df_hall_details[df_hall_details['role'] == "آذن"])
                st.markdown(f'<div class="stat-box">آذن<br><span class="stat-number">{count_serv}</span></div>', unsafe_allow_html=True)
            
            st.divider()
            # --- عرض الجدول ---
            df_to_show = df_hall_details[['name', 'role', 'school', 'city', 'phone']].copy()
            df_to_show.insert(0, 'م', range(1, 1 + len(df_to_show)))
            df_to_show.columns = ['الرقم', 'الاسم', 'المهمة', 'المدرسة', 'السكن', 'الجوال']
            st.markdown(df_to_show.style.set_properties(**{'text-align': 'right'}).hide(axis="index").to_html(), unsafe_allow_html=True)

            # --- أزرار الإجراءات ---
            st.markdown("<br>", unsafe_allow_html=True)
            cb1, cb2, cb3 = st.columns(3)
            with cb1:
                if st.button(f"🗑️ تفريغ القاعة", key="clr_h"):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE hall=?", (h_choice,))
                    conn.commit()
                    st.rerun()
            with cb2:
                if st.button(f"📄 إنشاء وورد", key="wd_h"):
                    bulk = generate_bulk_word(df_hall_details, h_choice)
                    st.download_button("تحميل الوورد", bulk, f"تكليفات_{h_choice}.docx")
            with cb3:
                # تصدير إكسل للقاعة
                output_h = io.BytesIO()
                df_hall_details.to_excel(output_h, index=False)
                st.download_button("📊 تحميل إكسل", output_h.getvalue(), f"كشف_{h_choice}.xlsx")

with tab_logs:
    st.markdown('<h2 class="move-to-right">📜 سجل العمليات</h2>', unsafe_allow_html=True)
    df_l = pd.read_sql("SELECT user, action, details, timestamp FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
