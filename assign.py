import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import time
from datetime import datetime

# =====================================
# 1. نظام تسجيل الدخول (ثابت كما هو)
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
# 2. الاختيار الرئيسي (التوجيهي vs التوظيف)
# =====================================
st.sidebar.title("🗂️ القائمة الرئيسية")
system_choice = st.sidebar.radio("اختر النظام المطلوب العمل عليه:", 
                                 ["مراقبة الثانوية العامة", "مراقبة امتحان التوظيف"])

# إعداد المتغيرات بناءً على النظام المختار
if system_choice == "مراقبة الثانوية العامة":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    SYS_TITLE = "نظام التوجيهي 2026"
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    SYS_TITLE = "نظام امتحان التوظيف 2026"

# =====================================
# 3. إعدادات الواجهة والاتصال
# =====================================
st.set_page_config(page_title=SYS_TITLE, layout="wide", initial_sidebar_state="expanded")

# CSS (نفس التنسيق السابق تماماً)
st.markdown("""
    <style>
    .main, .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    .user-box { background-color: #1a1c23; padding: 5px 15px; border-radius: 8px; border-right: 5px solid #00ffcc; display: inline-block; float: right; }
    .counter-card { background-color: #1a1c23; padding: 10px; border-radius: 10px; text-align: center; border: 1px solid #333; margin-bottom: 5px; }
    .counter-label { color: #bbb; font-size: 0.85rem; }
    .counter-value { color: #00ffcc; font-size: 1.5rem; font-weight: bold; }
    div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect { direction: rtl !important; text-align: right !important; }
    .move-to-right { text-align: right !important; direction: rtl !important; display: block; width: 100%; color: white; }
    .stat-card { flex: 1; padding: 15px; border-radius: 10px; text-align: center; min-width: 150px; border: 1px solid #333; }
    .stat-wants { border-top: 5px solid #28a745; background-color: #1a2e1f; }
    .stat-no-wants { border-top: 5px solid #dc3545; background-color: #2e1a1a; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect(DB_NAME, check_same_thread=False)
c = conn.cursor()

# إنشاء الجداول (محددة لتشمل الأعمدة الجديدة في التوظيف تلقائياً)
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT, 
             relative TEXT, relative_exam TEXT)''')

c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS logs 
             (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TEXT)''')
conn.commit()

# الدوال الأساسية (نفس المنطق السابق)
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

# =====================================
# 4. وظائف معالجة الوورد (المطورة للتوظيف)
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9: phone_val = '0' + phone_val
    
    repls = {
        'ZNAME': str(row.get('name', '')),
        'ZID': str(row.get('id', '')),
        'ZPHONE': phone_val,
        'ZJOB': str(row.get('role', '') or '---'),
        'ZHALL': str(h_name) if h_name and str(h_name).lower() != 'nan' else "---",
        'ZLOC': str(h_city) if h_city and str(h_city).lower() != 'nan' else "---",
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
    if not os.path.exists(TEMPLATE_NAME): return None
    doc = Document(TEMPLATE_NAME)
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

def generate_bulk_word(df, h_name):
    if not os.path.exists(TEMPLATE_NAME): return None
    final_doc = Document(TEMPLATE_NAME); final_doc._body.clear_content()
    for idx, row in df.iterrows():
        temp_doc = Document(TEMPLATE_NAME)
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        if idx > 0: final_doc.add_page_break()
        for element in temp_doc.element.body:
            if not element.tag.endswith('sectPr'): final_doc.element.body.append(element)
    out = io.BytesIO(); final_doc.save(out); out.seek(0)
    return out

# =====================================
# 5. التبويبات (نفس الوظائف السابقة)
# =====================================
st.title(f"🏢 {system_choice}")

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

with tab_search:
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال", key="search_input")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                st.write(f"المدرسة: {row['school']} | السكن: {row['city']}")
                if system_choice == "مراقبة امتحان التوظيف":
                    st.warning(f"قريب مباشر: {row['relative']} | امتحان القريب: {row['relative_exam']}")
                
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), key=f"sel_h_{row['id']}", index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0))
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], key=f"sel_r_{row['id']}", index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0))
                with c2:
                    if st.button("💾 حفظ", key=f"save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        conn.commit()
                        add_log("تكليف يدوي", f"تكليف {row['name']} في {sel_h}")
                        st.rerun()
                    if row['hall']:
                        f_word = generate_single_doc(row)
                        if f_word: st.download_button("📥 تحميل الكتاب", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

with tab_auto:
    df_all = get_cached_teachers()
    df_qualified = df_all[(df_all['ability'] == 'يصلح') & (df_all['preference'] == 'يرغب') & ((df_all['hall'] == '') | (df_all['hall'].isna()))]
    
    st.info(f"عدد المتاحين للتوزيع: {len(df_qualified)}")
    target_h = st.selectbox("القاعة المستهدفة", [""] + list(hall_map.keys()), key="auto_h")
    num = st.number_input("العدد المطلوب", min_value=0, max_value=len(df_qualified), value=0)
    
    if st.button("🚀 بدء التوزيع التلقائي"):
        if target_h and num > 0:
            sample = df_qualified.sample(n=int(num))
            for _, r in sample.iterrows():
                c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=?, updated_by='توزيع تلقائي' WHERE id=?", (target_h, hall_map[target_h], r['id']))
            conn.commit()
            add_log("توزيع تلقائي", f"توزيع {num} موظف على {target_h}")
            st.success("تم التوزيع بنجاح"); st.rerun()

with tab_upload:
    up_tpl = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
    if up_tpl:
        with open(TEMPLATE_NAME, "wb") as f: f.write(up_tpl.getbuffer())
        st.success("تم تحديث القالب")
    
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
            dft.columns = dft.columns.str.strip().str.lower()
            # التأكد من وجود الأعمدة الجديدة
            for col in ['relative', 'relative_exam', 'role', 'hall', 'hall_city', 'updated_by', 'preference', 'current_job', 'ability']:
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            pd.read_csv(HALLS_URL).to_sql('halls', conn, if_exists='replace', index=False)
            add_log("تحديث بيانات", "سحب من جوجل شيت")
            st.success("تم التحديث"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    df_all_teachers = get_cached_teachers()
    st.metric("إجمالي الموظفين", len(df_all_teachers))
    # عرض القاعات
    assigned_halls = sorted(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]['hall'].unique().tolist())
    h_choice = st.selectbox("عرض كادر قاعة:", [""] + assigned_halls)
    if h_choice:
        df_hall = df_all_teachers[df_all_teachers['hall'] == h_choice]
        st.table(df_hall[['name', 'role', 'school', 'phone']])
        if st.button("📄 تحميل كافة الكتب لهذه القاعة"):
            bulk_f = generate_bulk_word(df_hall, h_choice)
            if bulk_f: st.download_button("📥 اضغط للتحميل", data=bulk_f, file_name=f"كتب_{h_choice}.docx")

with tab_logs:
    df_l = pd.read_sql("SELECT * FROM logs ORDER BY id DESC LIMIT 50", conn)
    st.dataframe(df_l, use_container_width=True)
