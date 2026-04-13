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
import copy

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

# 🔧 التعديل 1: إضافة الوضع الثالث (نسخة معدلة بسيطة فقط)
if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان الثانوية العامة "
elif st.session_state['system_mode'] == "tasheeh":  # ← إضافة جديدة فقط
    DB_NAME = "data_tasheeh.db"                      # ← إضافة جديدة فقط
    TEMPLATE_NAME = "template.docx"                  # ← نستخدم نفس القالب
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=0&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=1885970999&single=true&output=csv"
    PAGE_TITLE = "نظام تصحيح الثانوية العامة"        # ← إضافة جديدة فقط
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان التوظيف"

st.set_page_config(page_title=PAGE_TITLE, layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
    <style>
        .custom-header {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            background-color: #1a1c23;
            color: white;
            text-align: center;
            padding: 15px 0;
            z-index: 999999;
            border-bottom: 2px solid #00ffcc;
            line-height: 1.5;
            direction: rtl;
            box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
        }
        .stApp {
            margin-top: 80px;
        }
        header {visibility: hidden;}
    </style>
    
    <div class="custom-header">
        <div style="font-weight: bold; font-size: 1.2rem;">إعداد وتصميم : عوض نعمان ريده</div>
        <div style="font-size: 1rem; color: #00ffcc;">قسم الامتحانات - مديرية التربية والتعليم جنوب نابلس</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
    <style>
    .main, .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    .user-box { background-color: #1a1c23; padding: 5px 15px; border-radius: 8px; border-right: 5px solid #00ffcc; display: inline-block; float: right; }
    .counter-card { background-color: #1a1c23; padding: 10px; border-radius: 10px; text-align: center; border: 1px solid #333; margin-bottom: 5px; }
    .counter-label { color: #bbb; font-size: 0.85rem; }
    .counter-value { color: #00ffcc; font-size: 1.5rem; font-weight: bold; }
    div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect { direction: rtl !important; text-align: right !important; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    .editor-info { color: #ffc107 !important; font-size: 0.9rem; font-weight: bold; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; color: #00ffcc !important; }
    .stat-card { flex: 1; padding: 15px; border-radius: 10px; text-align: center; min-width: 150px; border: 1px solid #333; }
    .stat-wants { border-top: 5px solid #28a745; background-color: #1a2e1f; }
    .stat-no-wants { border-top: 5px solid #dc3545; background-color: #2e1a1a; }
    .move-to-right { text-align: right !important; direction: rtl !important; display: block; width: 100%; color: white; }
    [data-testid="stSidebar"] { display: none; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect(DB_NAME, check_same_thread=False, timeout=30)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT,
             relative TEXT, relative_exam TEXT)''')

# 🔧 التعديل 2: إضافة عمود subject فقط
for col in ['relative', 'relative_exam', 'subject']:  # ← أضفنا subject فقط
    try:
        c.execute(f"ALTER TABLE teachers ADD COLUMN {col} TEXT DEFAULT ''")
        conn.commit()
    except:
        pass

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
        st.error(f"❌ ملف القالب '{TEMPLATE_NAME}' غير موجود، يرجى رفعه من تبويب 'رفع البيانات'")
        return None
    doc = Document(TEMPLATE_NAME)
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

def generate_bulk_word(df, h_name):
    if not os.path.exists(TEMPLATE_NAME):
        st.error(f"❌ ملف القالب '{TEMPLATE_NAME}' غير موجود، يرجى رفعه من تبويب 'رفع البيانات'")
        return None
        
    final_doc = Document(TEMPLATE_NAME)
    final_doc._body.clear_content()
    rows_list = list(df.iterrows())
    
    for i, (idx, row) in enumerate(rows_list):
        temp_doc = Document(TEMPLATE_NAME)
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        
        elements = [el for el in temp_doc.element.body if not el.tag.endswith('sectPr')]
        
        while elements:
            last = elements[-1]
            if last.tag.endswith('}p'):
                text = ''.join(t.text or '' for t in last.iter(qn('w:t')))
                if not text.strip():
                    elements.pop()
                    continue
            break
        
        for element in elements:
            final_doc.element.body.append(copy.deepcopy(element))
        
        if i < len(rows_list) - 1:
            p = OxmlElement('w:p')
            r = OxmlElement('w:r')
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            r.append(br)
            p.append(r)
            final_doc.element.body.append(p)
    
    out = io.BytesIO()
    final_doc.save(out)
    out.seek(0)
    return out

# =====================================
# 4. الواجهة الرئيسية
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
    # 🔧 التعديل 3: إضافة عمود ثالث للزر الجديد
    btn_col1, btn_col2, btn_col3, btn_spacer = st.columns([1, 1, 1, 1])  # ← غيرنا من [1,1,2] لـ [1,1,1,1]
    with btn_col1:
        if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
            switch_system("tawjihi")
    with btn_col2:
        if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
            switch_system("tawzif")
    # 🔧 إضافة الزر الثالث الجديد فقط
    with btn_col3:  # ← كتلة جديدة فقط
        if st.button("✅ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="tasheeh" else "secondary"):
            switch_system("tasheeh")

with header_col2:
    if st.button("🚪 تسجيل الخروج", key="logout_btn", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

# ... (كل كود التبويبات الأصلي كما هو تماماً - لم نعدله) ...
# للتوفير في المساحة، الكود الأصلي من tab_search إلى tab_logs يبقى كما أرسلته تماماً
# ✅ لم نغير أي سطر فيه

# ============================================================================
# ✨✨✨ بدايَة نظام تصحيح الثانوية العامة - وحدة مستقلة تماماً ✨✨✨
# يعمل فقط عند اختيار وضع "tasheeh" - لا يلمس أي جزء من الكود الأصلي
# ============================================================================

if st.session_state.get('system_mode') == "tasheeh":
    
    # === تهيئة الجداول الخاصة بالتصحيح ===
    c.execute('''CREATE TABLE IF NOT EXISTS tasheeh_assignments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  teacher_id TEXT, teacher_name TEXT, subject TEXT,
                  hall_name TEXT, hall_city TEXT, exam_name TEXT,
                  assignment_data TEXT, created_at TEXT, created_by TEXT)''')
    conn.commit()
    
    # === دوال مساعدة لنظام التصحيح فقط ===
    def load_tasheeh_teachers():
        try:
            df = pd.read_csv(TEACHERS_URL, dtype=str)
            df.columns = df.columns.str.strip().str.lower()
            rename_map = {
                'رقم الهوية': 'id', 'الاسم': 'name', 'المبحث': 'subject',
                'مكان سكن المعلم': 'city', 'اسم المدرسة': 'school',
                'رقم جواله': 'phone', 'هل له قريب مباشر او لا': 'relative'
            }
            df = df.rename(columns={k:v for k,v in rename_map.items() if k in df.columns})
            return df
        except Exception as e:
            st.error(f"❌ خطأ في تحميل بيانات المعلمين: {e}")
            return None
    
    def load_tasheeh_halls():
        try:
            df = pd.read_csv(HALLS_URL, dtype=str)
            df.columns = df.columns.str.strip().str.upper()
            return df if 'ZHALL' in df.columns and 'ZLOC' in df.columns else None
        except Exception as e:
            st.error(f"❌ خطأ في تحميل بيانات القاعات: {e}")
            return None
    
    def generate_tasheeh_letter(teacher_data, exam_name="امتحان الثانوية العامة"):
        """إنشاء كتاب تكليف بتصحيح باستخدام المتغيرات المطلوبة: ZNAME, ZID, ZTEST, ZHALL, ZLOC, ZWORK, ZCITY"""
        if not os.path.exists(TEMPLATE_NAME):
            return None
        
        doc = Document(TEMPLATE_NAME)
        replacements = {
            'ZNAME': teacher_data.get('name', '---'),
            'ZID': teacher_data.get('id', '---'),
            'ZTEST': exam_name,
            'ZHALL': teacher_data.get('hall_name', '---'),
            'ZLOC': teacher_data.get('hall_city', '---'),
            'ZWORK': teacher_data.get('school', '---'),
            'ZCITY': teacher_data.get('city', '---')
        }
        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    for run in para.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, str(val))
                            run.bold = True
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for key, val in replacements.items():
                            if key in para.text:
                                for run in para.runs:
                                    if key in run.text:
                                        run.text = run.text.replace(key, str(val))
                                        run.bold = True
        return doc
    
    # === واجهة نظام التصحيح ===
    st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a1c23 0%, #2d3748 100%); 
                    padding: 20px; border-radius: 15px; border: 2px solid #00ffcc;
                    margin: 20px 0; text-align: center;">
            <h2 style="color: #00ffcc; margin: 0;">✨ نظام تصحيح الثانوية العامة ✨</h2>
            <p style="color: #bbb; margin: 10px 0 0 0;">توزيع المصححين حسب المبحث والقاعة - تصدير كتب التكليف</p>
        </div>
    """, unsafe_allow_html=True)
    
    # تبويبات نظام التصحيح الداخلية
    corr_tab1, corr_tab2, corr_tab3, corr_tab4 = st.tabs([
        "📥 رفع البيانات", "🔄 التوزيع التلقائي", "📄 كتب التكليف", "📜 سجل العمليات"
    ])
    
    # === تبويب رفع البيانات ===
    with corr_tab1:
        st.markdown("### 📋 متطلبات ملف المعلمين (إكسل/CSV)")
        st.info("""
        **الأعمدة المطلوبة (بالعربي أو الإنجليزي):**
        - رقم الهوية / id
        - الاسم / name  
        - المبحث / subject *(مهم جداً)*
        - مكان سكن المعلم / city
        - اسم المدرسة / school *(اختياري)*
        - رقم الجوال / phone *(اختياري)*
        """)
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            if st.button("🔄 تحميل بيانات المعلمين من Google Sheets", use_container_width=True):
                with st.spinner("جاري التحميل..."):
                    df_t = load_tasheeh_teachers()
                    if df_t is not None:
                        st.session_state['tasheeh_teachers'] = df_t
                        st.success(f"✅ تم تحميل {len(df_t)} معلم/معلمة")
                        st.dataframe(df_t.head(3), use_container_width=True)
        with col_up2:
            if st.button("🏛️ تحميل بيانات القاعات من Google Sheets", use_container_width=True):
                with st.spinner("جاري التحميل..."):
                    df_h = load_tasheeh_halls()
                    if df_h is not None:
                        st.session_state['tasheeh_halls'] = df_h
                        st.success(f"✅ تم تحميل {len(df_h)} قاعة")
                        st.dataframe(df_h.head(3), use_container_width=True)
        
        st.divider()
        with st.expander("📤 رفع ملفات يدوياً (بدلاً من جوجل شيت)"):
            c1, c2 = st.columns(2)
            with c1:
                manual_teachers = st.file_uploader("ملف المعلمين", type=['csv', 'xlsx'], key="mt_up")
                if manual_teachers:
                    try:
                        df = pd.read_csv(manual_teachers) if manual_teachers.name.endswith('.csv') else pd.read_excel(manual_teachers)
                        st.session_state['tasheeh_teachers'] = df
                        st.success("✅ تم رفع الملف")
                    except Exception as e:
                        st.error(f"خطأ: {e}")
            with c2:
                manual_halls = st.file_uploader("ملف القاعات", type=['csv', 'xlsx'], key="mh_up")
                if manual_halls:
                    try:
                        df = pd.read_csv(manual_halls) if manual_halls.name.endswith('.csv') else pd.read_excel(manual_halls)
                        st.session_state['tasheeh_halls'] = df
                        st.success("✅ تم رفع الملف")
                    except Exception as e:
                        st.error(f"خطأ: {e}")
    
    # === تبويب التوزيع التلقائي ===
    with corr_tab2:
        if 'tasheeh_teachers' not in st.session_state or 'tasheeh_halls' not in st.session_state:
            st.warning("⚠️ يرجى تحميل بيانات المعلمين والقاعات أولاً من تبويب 'رفع البيانات'")
        else:
            teachers_df = st.session_state['tasheeh_teachers']
            halls_df = st.session_state['tasheeh_halls']
            
            st.markdown("### 🎯 إعدادات التوزيع التلقائي")
            col_dist1, col_dist2 = st.columns(2)
            with col_dist1:
                hall_options = [""] + list(halls_df['ZHALL'].unique())
                selected_hall = st.selectbox("🏛️ اختر القاعة (اختياري):", hall_options)
            with col_dist2:
                if 'subject' in teachers_df.columns:
                    subject_options = [""] + sorted(teachers_df['subject'].dropna().unique().tolist())
                    selected_subject = st.selectbox("📚 اختر المبحث (اختياري):", subject_options)
                else:
                    st.error("❌ عمود 'المبحث' غير موجود في بيانات المعلمين")
                    selected_subject = ""
            
            exam_name = st.text_input("📝 اسم الامتحان:", value="امتحان الثانوية العامة 2026")
            st.divider()
            
            if selected_subject:
                preview_df = teachers_df[teachers_df['subject'] == selected_subject]
            else:
                preview_df = teachers_df
            
            st.markdown(f"👥 عدد المعلمين المرشحين للتوزيع: **{len(preview_df)}**")
            if len(preview_df) > 0:
                st.dataframe(preview_df[['name', 'subject', 'city']].head(5), use_container_width=True)
            
            if st.button("🚀 تنفيذ التوزيع التلقائي", type="primary", use_container_width=True):
                assignments = []
                for _, teacher in preview_df.iterrows():
                    if selected_hall:
                        candidate_halls = halls_df[halls_df['ZHALL'] == selected_hall]
                    else:
                        city_halls = halls_df[halls_df['ZLOC'] == teacher.get('city', '')]
                        candidate_halls = city_halls if not city_halls.empty else halls_df
                    
                    if not candidate_halls.empty:
                        hall_info = candidate_halls.sample(n=1).iloc[0]
                        assignment = {
                            'id': teacher.get('id', ''), 'name': teacher.get('name', ''),
                            'subject': teacher.get('subject', ''), 'hall_name': hall_info['ZHALL'],
                            'hall_city': hall_info['ZLOC'], 'exam_name': exam_name,
                            'school': teacher.get('school', ''), 'city': teacher.get('city', ''),
                            'phone': teacher.get('phone', ''),
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M")
                        }
                        assignments.append(assignment)
                
                if assignments:
                    st.session_state['tasheeh_assignments'] = assignments
                    add_log("توزيع تصحيح", f"تم توزيع {len(assignments)} معلم لمبحث {selected_subject or 'جميع المباحث'}")
                    st.success(f"✅ تم توزيع {len(assignments)} معلم/معلمة بنجاح!")
                    try:
                        for a in assignments:
                            c.execute("""INSERT INTO tasheeh_assignments 
                                         (teacher_id, teacher_name, subject, hall_name, hall_city, exam_name, created_at, created_by)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                                     (a['id'], a['name'], a['subject'], a['hall_name'], a['hall_city'], 
                                      a['exam_name'], a['timestamp'], st.session_state.username))
                        conn.commit()
                    except: pass
                else:
                    st.warning("⚠️ لم يتم توزيع أي معلم - تأكد من تطابق البيانات")
    
    # === تبويب كتب التكليف ===
    with corr_tab3:
        if 'tasheeh_assignments' not in st.session_state or not st.session_state['tasheeh_assignments']:
            st.info("📌 قم بالتوزيع التلقائي أولاً لإنشاء كتب التكليف")
        else:
            assignments = st.session_state['tasheeh_assignments']
            st.markdown(f"### 📄 كتب التكليف الجاهزة: **{len(assignments)}**")
            
            export_col1, export_col2 = st.columns(2)
            with export_col1:
                if st.button("📦 إنشاء ملفات وورد فردية", use_container_width=True):
                    if not os.path.exists(TEMPLATE_NAME):
                        st.error(f"❌ ملف القالب '{TEMPLATE_NAME}' غير موجود")
                    else:
                        docs = []
                        for a in assignments:
                            doc = generate_tasheeh_letter(a, a['exam_name'])
                            if doc:
                                bio = io.BytesIO()
                                doc.save(bio)
                                bio.seek(0)
                                fname = f"تكليف_{a['name']}_{a['id']}.docx"
                                docs.append((bio, fname, a['name']))
                        if docs:
                            st.success(f"✅ تم إنشاء {len(docs)} ملف وورد")
                            for bio, fname, tname in docs[:10]:
                                st.download_button(label=f"📥 {tname}", data=bio.getvalue(), file_name=fname,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_corr_{fname}")
                            if len(docs) > 10:
                                st.info(f"⚠️ تم عرض أول 10 ملفات فقط. إجمالي الملفات: {len(docs)}")
            
            with export_col2:
                if st.button("📊 تصدير إكسل شامل", use_container_width=True):
                    df_assign = pd.DataFrame(assignments)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_assign.to_excel(writer, index=False, sheet_name='التكليفات')
                    output.seek(0)
                    st.download_button(label="📥 تحميل ملف الإكسل", data=output.getvalue(),
                        file_name=f"تكاليف_التصحيح_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_corr_excel")
                    add_log("تصدير تصحيح", "تم تصدير ملف الإكسل الشامل")
    
    # === تبويب سجل العمليات ===
    with corr_tab4:
        st.markdown("### 📜 سجل عمليات نظام التصحيح")
        df_l = pd.read_sql("""
            SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' 
            FROM logs WHERE action LIKE '%تصحيح%' OR action LIKE '%توزيع تصحيح%'
            ORDER BY id DESC LIMIT 100
        """, conn)
        if not df_l.empty:
            st.dataframe(df_l, use_container_width=True)
        else:
            st.info("لا توجد سجلات لنظام التصحيح بعد")
        
        if st.button("🗑️ مسح سجل التصحيح فقط", key="clear_corr_logs"):
            try:
                # لا نحذف كل السجلات، فقط نعرض رسالة تأكيد
                st.warning("⚠️ سجلات النظام مشتركة - لا يمكن حذف سجلات التصحيح فقط دون التأثير على الأنظمة الأخرى")
            except Exception as e:
                st.error(f"خطأ: {e}")
    
    # إنهاء العرض هنا لعدم عرض التبويبات الأخرى عند اختيار وضع التصحيح
    st.stop()
