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
    st.session_state['popover_counter'] = 0

if 'system_mode' not in st.session_state:
    st.session_state['system_mode'] = "tawjihi"

# دالة التبديل لضمان تنظيف الذاكرة ومنع تداخل البيانات
def switch_system(mode):
    st.session_state['system_mode'] = mode
    st.cache_data.clear()  # مسح الكاش لضمان تحديث الأرقام والبيانات
    st.rerun()

# تحديد المتغيرات بناءً على النظام النشط
if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التوجيهي 2026"
else:
    DB_NAME = "data_tawzif.db"
    TEMPLATE_NAME = "template_tawzif.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=821672282&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTIka1g67VWzR7UKmdR6eb79WuCFaC-qTNTeNMYbjzkz_HmBR_Qwe6o5RGbPyPqiaY_y_z3k2YdbibO/pub?gid=932943855&single=true&output=csv"
    PAGE_TITLE = "نظام امتحان التوظيف 2026"

st.set_page_config(page_title=PAGE_TITLE, layout="wide", initial_sidebar_state="collapsed")
# --- إضافة الترويسة الثابتة في أعلى الصفحة ---
# --- إضافة الترويسة الثابتة في أعلى الصفحة بصيغة مطورة ---
st.markdown("""
    <style>
        /* تثبيت الترويسة ومنعها من الاختفاء */
        .custom-header {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            background-color: #1a1c23;
            color: white;
            text-align: center;
            padding: 15px 0;
            z-index: 999999; /* رقم عالي جداً لضمان الظهور فوق كل شيء */
            border-bottom: 2px solid #00ffcc;
            line-height: 1.5;
            direction: rtl;
            box-shadow: 0px 4px 10px rgba(0,0,0,0.5);
        }
        
        /* إزاحة محتوى التطبيق لأسفل لكي لا تغطيه الترويسة */
        .stApp {
            margin-top: 80px;
        }

        /* إخفاء الهيدر الافتراضي لستريمليت لزيادة المساحة (اختياري) */
        header {visibility: hidden;}
    </style>
    
    <div class="custom-header">
        <div style="font-weight: bold; font-size: 1.2rem;">إعداد وتصميم : عوض نعمان ريده</div>
        <div style="font-size: 1rem; color: #00ffcc;">قسم الامتحانات - مديرية التربية والتعليم جنوب نابلس</div>
    </div>
    """, unsafe_allow_html=True)
# (تنسيقات CSS)
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

conn = sqlite3.connect(DB_NAME, check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT,
             relative TEXT, relative_exam TEXT)''')

try:
    c.execute("ALTER TABLE teachers ADD COLUMN relative TEXT DEFAULT ''")
    c.execute("ALTER TABLE teachers ADD COLUMN relative_exam TEXT DEFAULT ''")
except:
    pass 

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
    if not os.path.exists(TEMPLATE_NAME): return None
    doc = Document(TEMPLATE_NAME)
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

def generate_bulk_word(df, h_name):
    if not os.path.exists(TEMPLATE_NAME):
        return None
        
    final_doc = Document(TEMPLATE_NAME)
    final_doc._body.clear_content()
    rows_list = list(df.iterrows())
    
    for i, (idx, row) in enumerate(rows_list):
        temp_doc = Document(TEMPLATE_NAME)
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        
        elements = [el for el in temp_doc.element.body if not el.tag.endswith('sectPr')]
        
        # حذف الفقرات الفارغة من نهاية كل كتاب
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
        
        # page break بعد كل كتاب إلا الأخير
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
    btn_col1, btn_col2, btn_spacer = st.columns([1, 1, 2])
    with btn_col1:
        if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
            switch_system("tawjihi")
    with btn_col2:
        if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
            switch_system("tawzif")

with header_col2:
    if st.button("🚪 تسجيل الخروج", key="logout_btn", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

with tab_search:
    st.markdown(f'<h2 class="move-to-right">إدارة الموظفين - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            display_phone = str(row.get('phone', '---'))
            if display_phone.startswith('5') and len(display_phone) == 9:
                display_phone = '0' + display_phone

            with st.expander(f"👤 {row.get('name', 'اسم غير معروف')} | القاعة: {row.get('hall') or 'غير مكلف'}"):
                def safe_get(key):
                    v = str(row.get(key, '---')).strip()
                    return '---' if v.lower() in ['nan', 'none', ''] else v

                v_id = safe_get('id')
                v_city = safe_get('city')
                v_school = safe_get('school')
                v_job = safe_get('current_job')
                v_abil = safe_get('ability')
                v_pref = safe_get('preference')

                rel_html = ''
                if st.session_state.system_mode == 'tawzif':
                    v_rel = safe_get('relative')
                    v_relex = safe_get('relative_exam')
                    rel_html = f'<tr><td style="padding: 5px; color: #ffc107;"><b>🔗 قريب:</b> {v_rel}</td><td style="padding: 5px; color: #ffc107;"><b>📝 الامتحان:</b> {v_relex}</td></tr>'

                full_table = f'<div style="background-color: #1a1c23; padding: 15px; border-radius: 10px; border: 1px solid #444; border-right: 5px solid #00ffcc; margin-bottom: 15px; text-align: right; direction: rtl;"><table style="width:100%; color: white; border: none;"><tr><td style="padding: 5px;"><b>🆔 الهوية:</b> {v_id}</td><td style="padding: 5px;"><b>📱 الجوال:</b> {display_phone}</td></tr><tr><td style="padding: 5px;"><b>🏡 السكن:</b> {v_city}</td><td style="padding: 5px;"><b>🏫 المدرسة:</b> {v_school}</td></tr><tr><td style="padding: 5px;"><b>📝 الرغبة:</b> {v_pref}</td><td style="padding: 5px;"><b>💼 الوظيفة:</b> {v_job}</td></tr>{rel_html}<tr><td colspan="2" style="padding: 5px; border-top: 1px solid #444; color: #ffc107;"><b>⚠️ صلاحية المراقبة:</b> {v_abil}</td></tr></table></div>'
                st.markdown(full_table, unsafe_allow_html=True)
                st.markdown(f"<span class='editor-info'>آخر تعديل: {row['updated_by'] or 'لا يوجد'}</span>", unsafe_allow_html=True)
                
                # --- التعديل لحل مشكلة التكرار ---
                with st.popover("📝 تعديل البيانات الأساسية", key=f"pop_{st.session_state.system_mode}_{row['id']}_{st.session_state.popover_counter}"):
                    u_name = st.text_input("الاسم", value=row['name'], key=f"un_{st.session_state.system_mode}_{row['id']}")
                    u_phone = st.text_input("رقم الجوال", value=display_phone, key=f"up_{st.session_state.system_mode}_{row['id']}")
                    u_school = st.text_input("المدرسة", value=row['school'], key=f"us_{st.session_state.system_mode}_{row['id']}")
                    u_city = st.text_input("السكن", value=row['city'], key=f"uc_{st.session_state.system_mode}_{row['id']}")
                    u_job = st.text_input("الوظيفة الأساسية", value=row['current_job'], key=f"uj_{st.session_state.system_mode}_{row['id']}")
                    u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2), key=f"upr_{st.session_state.system_mode}_{row['id']}")
                    u_abil = st.selectbox("صلاحية المراقبة", ["يصلح", "لا يصلح", "لم تحدد"], index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2), key=f"uab_{st.session_state.system_mode}_{row['id']}")
                    
                    if st.session_state.system_mode == "tawzif":
                        u_rel = st.selectbox("هل له قريب؟", ["نعم", "لا"], index=0 if row.get('relative')=="نعم" else 1, key=f"urel_{st.session_state.system_mode}_{row['id']}")
                        u_relex = st.text_input("اسم امتحان القريب", value=row.get('relative_exam', ''), key=f"urex_{st.session_state.system_mode}_{row['id']}")

                    if st.button("💾 تحديث وحفظ", key=f"save_base_{st.session_state.system_mode}_{row['id']}"):
                        if st.session_state.system_mode == "tawzif":
                            c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, relative=?, relative_exam=?, updated_by=? 
                                         WHERE id=?""", (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, u_rel, u_relex, st.session_state.username, row['id']))
                        else:
                            c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, updated_by=? 
                                         WHERE id=?""", (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                        conn.commit()
                        add_log("تعديل بيانات أساسية", f"تعديل بيانات {u_name}")
                        st.session_state.popover_counter += 1
                        st.success("✅ تم الحفظ")
                        time.sleep(0.5)
                        st.rerun()

                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    current_hall = row['hall'] if row['hall'] and str(row['hall']).lower() != 'nan' else ""
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), 
                                         index=(list(hall_map.keys()).index(current_hall)+1 if current_hall in hall_map else 0), 
                                         key=f"q_h_{st.session_state.system_mode}_{row['id']}")
                    
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], 
                                         index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0),
                                         key=f"q_r_{st.session_state.system_mode}_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{st.session_state.system_mode}_{row['id']}"):
                        h_city_val = hall_map.get(sel_h, "")
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                                  (sel_h, sel_r, h_city_val, st.session_state.username, row['id']))
                        conn.commit()
                        add_log("حفظ تكليف", f"تم تكليف {row['name']} في {sel_h}")
                        st.success("✅ تم الحفظ")
                        time.sleep(0.5)
                        st.rerun()
                    
                    is_assigned = row['hall'] and str(row['hall']).strip() != "" and str(row['hall']).lower() != 'nan'
                    
                    if is_assigned:
                        if st.button("❌ إلغاء التكليف", key=f"del_search_{st.session_state.system_mode}_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", 
                                      (st.session_state.username, row['id']))
                            conn.commit()
                            add_log("إلغاء تكليف", f"تم إلغاء تكليف {row['name']}")
                            st.rerun()
                        
                        if st.button("📥 إنشاء الكتاب", key=f"gen_s_{st.session_state.system_mode}_{row['id']}"):
                            f_word = generate_single_doc(row)
                            if f_word: 
                                st.download_button("📥 تحميل الآن", data=f_word, 
                                               file_name=f"تكليف_{row['name']}.docx", 
                                               key=f"dl_s_{st.session_state.system_mode}_{row['id']}")

with tab_auto:
    st.markdown('<h2 class="move-to-right">🤖 نظام التوزيع التلقائي الذكي</h2>', unsafe_allow_html=True)
    df_all = get_cached_teachers()
    hall_map_auto = {r['hall_name']: r['city'] for _, r in get_cached_halls().iterrows()}
    
    df_qualified = df_all[
        (df_all['ability'] == 'يصلح') & 
        (df_all['preference'] == 'يرغب') & 
        (df_all['current_job'] == 'معلم') &
        ((df_all['hall'] == '') | (df_all['hall'].isna()))
    ]

    can_and_wants = len(df_qualified)
    can_not_wants = len(df_all[(df_all['ability'] == 'يصلح') & (df_all['preference'] == 'لا يرغب') & (df_all['current_job'] == 'معلم') & ((df_all['hall'] == '') | (df_all['hall'].isna()))])
    
    st.markdown(f"""
    <div style="display: flex; gap: 15px; margin-bottom: 20px; direction: rtl;">
        <div class="stat-card stat-wants">
            <span style="color: #bbb; font-size: 0.9rem;">متاح (يصلح ويرغب)</span><br>
            <strong style="font-size: 2rem; color: #28a745;">{can_and_wants}</strong>
        </div>
        <div class="stat-card stat-no-wants">
            <span style="color: #bbb; font-size: 0.9rem;">متاح (يصلح ولا يرغب)</span><br>
            <strong style="font-size: 2rem; color: #dc3545;">{can_not_wants}</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)

    available_cities = sorted(df_qualified['city'].unique().tolist()) if not df_qualified.empty else []
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_h = st.selectbox("اختر القاعة المستهدفة:", [""] + list(hall_map_auto.keys()), key="auto_target_h")
        selected_cities = st.multiselect("السحب من مناطق سكن محددة (اختياري):", available_cities)
        
    with col_a2:
        if selected_cities:
            df_pool = df_qualified[df_qualified['city'].isin(selected_cities)]
        else:
            df_pool = df_qualified
            
        st.info(f"عدد المعلمين المتاحين للسحب الآن: {len(df_pool)}")
        num_to_assign = st.number_input("العدد المطلوب توزيعه:", min_value=0, max_value=len(df_pool) if not df_pool.empty else 0, value=0)

        if st.button("🚀 ابدأ التوزيع التلقائي الآن", use_container_width=True, disabled=(num_to_assign == 0 or not target_h)):
            selected_sample = df_pool.sample(n=int(num_to_assign))
            for _, r in selected_sample.iterrows():
                c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=?, updated_by='توزيع تلقائي' WHERE id=?", 
                          (target_h, hall_map_auto[target_h], r['id']))
            conn.commit()
            add_log("توزيع تلقائي", f"توزيع {num_to_assign} معلم على قاعة {target_h}")
            st.success(f"✅ تم توزيع {num_to_assign} بنجاح!")
            time.sleep(1)
            st.rerun()

with tab_upload:
    st.markdown(f'<h2 class="move-to-right">تحديث القالب والبيانات - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
    up_tpl = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
    if up_tpl:
        with open(TEMPLATE_NAME, "wb") as f:
            f.write(up_tpl.getbuffer())
        add_log("تحديث قالب", f"تم رفع قالب {TEMPLATE_NAME} جديد")
        st.success("تم تحديث قالب الوورد بنجاح")
    
    st.divider()
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str}) 
            dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            for col in ['phone', 'role', 'hall', 'hall_city', 'updated_by', 'preference', 'current_job', 'ability', 'relative', 'relative_exam']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            
            dfh = pd.read_csv(HALLS_URL)
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            
            add_log("تحديث بيانات", "تحديث من جوجل شيت")
            st.success("تم التحديث بنجاح")
            st.cache_data.clear()
            st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    df_all_teachers = get_cached_teachers()
    total_count = len(df_all_teachers)
    assigned_count = len(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0])
    remaining_count = total_count - assigned_count

    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي الموظفين", total_count)
    c_m2.metric("تم إنجازهم", assigned_count)
    c_m3.metric("المتبقي", remaining_count)
    
    st.divider()
    st.markdown('<h3 class="move-to-right">📦 تصدير البيانات المعدلة</h3>', unsafe_allow_html=True)
    df_export = df_all_teachers.copy()
    
    arabic_cols = ['رقم الهوية', 'الاسم كامل', 'رقم الجوال', 'المدرسة', 'السكن', 'المهمة المكلف بها', 'القاعة', 'مدينة القاعة', 'الموظف المعدل', 'الرغبة', 'الوظيفة', 'الصلاحية', 'قريب مباشر', 'امتحان القريب']
    df_export.columns = arabic_cols[:len(df_export.columns)]
    
    output_all = io.BytesIO()
    with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
        # تصدير البيانات للشيت
        df_export.to_excel(writer, index=False, sheet_name='الموظفين')
        workbook = writer.book
        worksheet = writer.sheets['الموظفين']
        
        # 1. تنسيق الرأس (خط 14، عريض، خلفية ملونة، حدود كاملة)
        h_fmt = workbook.add_format({
            'bold': True, 
            'font_size': 14, 
            'border': 1, 
            'align': 'center', 
            'valign': 'vcenter', 
            'bg_color': '#D7E4BC'
        })

        # 2. تنسيق محتوى الخلايا (خط 14، عريض، حدود كاملة)
        c_fmt = workbook.add_format({
            'bold': True, 
            'font_size': 14, 
            'border': 1, 
            'align': 'right', 
            'valign': 'vcenter'
        })
        
        # 3. إعدادات اتجاه الورقة والطباعة
        worksheet.right_to_left()      # من اليمين لليسار
        worksheet.set_landscape()      # طباعة بالعرض (Horizontal)
        worksheet.fit_to_pages(1, 0)   # ضغط كل الأعمدة لتظهر في ورقة واحدة عند الطباعة
        
        # 4. حلقة لتنسيق كل عمود وضبط عرضه تلقائياً حسب المحتوى
        for col_num, col_name in enumerate(df_export.columns):
            # إعادة كتابة الرأس بالتنسيق الجديد
            worksheet.write(0, col_num, col_name, h_fmt)
            
            # حساب طول المحتوى بأمان (لمنع خطأ TypeError)
            column_data = df_export[col_name].astype(str).str.len()
            max_data_len = column_data.max() if not column_data.empty else 0
            
            # حساب العرض المناسب (الأكبر بين طول البيانات وطول اسم العمود)
            # أضفنا 6 كمتسع إضافي لأن الخط 14 Bold يأخذ مساحة أكبر
            calculated_width = max(max_data_len, len(str(col_name))) + 6
            
            # تحديد حد أقصى للعرض 50 كي لا تصبح الأعمدة ضخمة جداً
            final_width = min(calculated_width, 50)
            
            # تطبيق العرض والتنسيق (الحدود والخط 14) على العمود بالكامل
            worksheet.set_column(col_num, col_num, final_width, c_fmt)
    
    st.download_button("📥 تحميل إكسل معدل", data=output_all.getvalue(), file_name=f"كشف_معدل_{st.session_state.system_mode}_{datetime.now().strftime('%Y%m%d')}.xlsx")

    st.divider()
    assigned_halls = sorted(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]['hall'].unique().tolist())
    
    if assigned_halls:
        h_choice = st.selectbox("اختر قاعة لعرض الكادر والإحصائيات:", [""] + assigned_halls)
        if h_choice:
            df_hall_details = df_all_teachers[df_all_teachers['hall'] == h_choice].copy()
            
            st.markdown(f'<h4 class="move-to-right">🔢 معداد قاعة: {h_choice}</h4>', unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            with m1: st.markdown(f'<div class="counter-card"><div class="counter-label">رئيس قاعة</div><div class="counter-value">{len(df_hall_details[df_hall_details["role"] == "رئيس قاعة"])}</div></div>', unsafe_allow_html=True)
            with m2: st.markdown(f'<div class="counter-card"><div class="counter-label">مساعد رئيس</div><div class="counter-value">{len(df_hall_details[df_hall_details["role"] == "مساعد رئيس قاعة"])}</div></div>', unsafe_allow_html=True)
            with m3: st.markdown(f'<div class="counter-card"><div class="counter-label">مراقب</div><div class="counter-value">{len(df_hall_details[df_hall_details["role"] == "مراقب"])}</div></div>', unsafe_allow_html=True)
            with m4: st.markdown(f'<div class="counter-card"><div class="counter-label">آذن</div><div class="counter-value">{len(df_hall_details[df_hall_details["role"] == "آذن"])}</div></div>', unsafe_allow_html=True)
            
            st.divider()
            
            if not df_hall_details.empty:
                df_to_show = df_hall_details[['name', 'role', 'school', 'city', 'phone']].copy()
                df_to_show.insert(0, 'م', range(1, 1 + len(df_to_show)))
                df_to_show.columns = ['الرقم', 'الاسم', 'المهمة', 'المدرسة', 'السكن', 'الجوال']
                
                styled_df = df_to_show.style.set_properties(**{
                    'text-align': 'right',
                    'direction': 'rtl'
                }).hide(axis="index")
                
                st.markdown(styled_df.to_html(), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            col_btns1, col_btns2, col_btns3 = st.columns([1, 1.2, 1.2])
            
            with col_btns1:
                if st.button(f"🗑️ تفريغ قاعة {h_choice}", key=f"del_hall_{h_choice}"):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE hall=?", (st.session_state.username, h_choice))
                    conn.commit()
                    add_log("تفريغ قاعة", f"تم مسح كافة تكليفات قاعة {h_choice}")
                    st.success("تم تفريغ القاعة")
                    time.sleep(0.5)
                    st.rerun()
            
            with col_btns2:
                if st.button(f"📄 إنشاء كتب قاعة {h_choice}", key=f"gen_bulk_{h_choice}"):
                    bulk_f = generate_bulk_word(df_hall_details, h_choice)
                    if bulk_f: 
                        st.download_button("📥 تحميل الوورد", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")
            
            with col_btns3:
                output_hall_excel = io.BytesIO()
                df_hall_excel = df_hall_details.copy()
                df_hall_excel.insert(0, 'الرقم', range(1, 1 + len(df_hall_excel)))
                df_final_export = df_hall_excel[['الرقم', 'name', 'id', 'phone', 'school', 'role', 'city']]
                df_final_export.columns = ['الرقم', 'الاسم الرباعي', 'رقم الهوية', 'رقم الجوال', 'المدرسة', 'المهمة', 'العنوان']

                with pd.ExcelWriter(output_hall_excel, engine='xlsxwriter') as writer:
                    df_final_export.to_excel(writer, index=False, sheet_name='كشف_القاعة')
                    workbook = writer.book
                    worksheet = writer.sheets['كشف_القاعة']
                    
                    # 1. تعريف التنسيقات
                    h_fmt = workbook.add_format({
                        'bold': True, 'font_size': 14, 'border': 1, 
                        'align': 'center', 'valign': 'vcenter', 'bg_color': '#BDD7EE'
                    })
                    c_fmt = workbook.add_format({
                        'bold': True, 'font_size': 14, 'border': 1,
                        'align': 'right', 'valign': 'vcenter'
                    })
                    
                    # 2. إعدادات اتجاه الصفحة والطباعة
                    worksheet.right_to_left()
                    worksheet.set_landscape()
                    worksheet.fit_to_pages(1, 0)
                    
                    # 3. تطبيق التنسيق وضبط عرض الأعمدة تلقائياً
                    for col_num, col_name in enumerate(df_final_export.columns):
                        # كتابة العنوان بتنسيق الرأس
                        worksheet.write(0, col_num, col_name, h_fmt)
                        
                        # حساب أقصى طول في العمود الحالي (بين اسم العمود والبيانات)
                        column_length = max(
                            df_final_export[col_name].astype(str).map(len).max(),
                            len(str(col_name))
                        ) + 4 # زيادة بسيطة للهامش
                        
                        # ضبط عرض العمود بناءً على الطول المحسوب وتطبيق التنسيق
                        worksheet.set_column(col_num, col_num, column_length, c_fmt)
                
                st.download_button(f"📊 كشف إكسل {h_choice}", data=output_hall_excel.getvalue(), file_name=f"كشف_{h_choice}.xlsx")

with tab_logs:
    st.markdown('<h2 class="move-to-right">📜 سجل العمليات</h2>', unsafe_allow_html=True)
    if st.button("🗑️ حذف كافة السجلات نهائياً", key="clear_all_logs"):
        try:
            c.execute("DELETE FROM logs")
            conn.commit()
            st.success("✅ تم مسح سجل العمليات بالكامل")
            time.sleep(0.5)
            st.rerun()
        except Exception as e:
            st.error(f"خطأ أثناء الحذف: {e}")

    st.divider()
    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    if not df_l.empty:
        st.dataframe(df_l, use_container_width=True)
    else:
        st.info("سجل العمليات فارغ حالياً.")
