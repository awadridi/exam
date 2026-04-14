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

# 🔧 التعديل 1: إضافة الوضع الثالث
if st.session_state['system_mode'] == "tawjihi":
    DB_NAME = "data_system_v26.db"
    TEMPLATE_NAME = "template.docx"
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"
    PAGE_TITLE = "نظام التكليفات امتحان الثانوية العامة "
elif st.session_state['system_mode'] == "tasheeh":
    DB_NAME = "data_tasheeh.db"
    TEMPLATE_NAME = "template_tasheeh.docx"  # 👈 اسم جديد خاص بقالب التصحيح
    TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=0&single=true&output=csv"
    HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVP8cQV8GHlaWXETc9rGzteNwDVPg8iyyZ9zCXFq-J1_t0q4sxveFchsN5XbuTiZgJBeTpC3VBMc7k/pub?gid=1885970999&single=true&output=csv"
    PAGE_TITLE = "نظام تصحيح الثانوية العامة"
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

# 🔧 التعديل 2: إضافة عمود subject
for col in ['relative', 'relative_exam', 'subject']:
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
    btn_col1, btn_col2, btn_col3, btn_spacer = st.columns([1, 1, 1, 1])
    with btn_col1:
        if st.button("📝 الثانوية العامة", use_container_width=True, type="primary" if st.session_state.system_mode=="tawjihi" else "secondary"):
            switch_system("tawjihi")
    with btn_col2:
        if st.button("👨‍🏫 امتحان التوظيف", use_container_width=True, type="primary" if st.session_state.system_mode=="tawzif" else "secondary"):
            switch_system("tawzif")
    with btn_col3:
        if st.button("✅ تصحيح الثانوية", use_container_width=True, type="primary" if st.session_state.system_mode=="tasheeh" else "secondary"):
            switch_system("tasheeh")

with header_col2:
    if st.button("🚪 تسجيل الخروج", key="logout_btn", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

st.divider()

# 🔴🔴🔴 التبويبات الأصلية تظهر فقط إذا لم يكن الوضع "تصحيح الثانوية" 🔴🔴🔴
if st.session_state['system_mode'] != "tasheeh":
    
    tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs([
        "🔍 البحث والتعيين", 
        "🤖 التوزيع التلقائي", 
        "📥 رفع البيانات", 
        "📊 الإدارة والإحصائيات", 
        "📜 سجل العمليات"
    ])

    # ==================== تبويب البحث ====================
    with tab_search:
        st.markdown(f'<h2 class="move-to-right">إدارة الموظفين - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
        df_h_data = get_cached_halls()
        hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
        
        q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
        if q:
            df_teachers = get_cached_teachers()
            results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
            
            for idx, row in results.iterrows():
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
                    
                    with st.popover("📝 تعديل البيانات", key=f"pop_{row['id']}_{idx}_{st.session_state.popover_counter}"):
                        u_name = st.text_input("الاسم", value=row['name'], key=f"un_{st.session_state.system_mode}_{row['id']}_{idx}")
                        u_phone = st.text_input("رقم الجوال", value=display_phone, key=f"up_{st.session_state.system_mode}_{row['id']}_{idx}")
                        u_school = st.text_input("المدرسة", value=row['school'], key=f"us_{st.session_state.system_mode}_{row['id']}_{idx}")
                        u_city = st.text_input("السكن", value=row['city'], key=f"uc_{st.session_state.system_mode}_{row['id']}_{idx}")
                        u_job = st.text_input("الوظيفة الأساسية", value=row['current_job'], key=f"uj_{st.session_state.system_mode}_{row['id']}_{idx}")
                        
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], 
                                            index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2), 
                                            key=f"upr_{st.session_state.system_mode}_{row['id']}_{idx}")
                        
                        u_abil = st.selectbox("صلاحية المراقبة", ["يصلح", "لا يصلح", "لم تحدد"], 
                                            index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2), 
                                            key=f"uab_{st.session_state.system_mode}_{row['id']}_{idx}")

                        if st.session_state.system_mode == "tawzif":
                            u_rel = st.selectbox("هل له قريب؟", ["نعم", "لا"], index=0 if row.get('relative')=="نعم" else 1, key=f"urel_{row['id']}_{idx}")
                            u_relex = st.text_input("اسم امتحان القريب", value=row.get('relative_exam', ''), key=f"urex_{row['id']}_{idx}")

                        if st.button("💾 تحديث وحفظ", key=f"save_base_{row['id']}_{idx}_{st.session_state.popover_counter}"):
                            if st.session_state.system_mode == "tawzif":
                                c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, relative=?, relative_exam=?, updated_by=? WHERE id=?""", 
                                         (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, u_rel, u_relex, st.session_state.username, row['id']))
                            else:
                                c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, updated_by=? WHERE id=?""", 
                                         (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                            conn.commit()
                            add_log("تعديل بيانات أساسية", f"تعديل بيانات {u_name}")
                            st.session_state.popover_counter += 1
                            st.cache_data.clear()
                            st.success("✅ تم الحفظ")
                            time.sleep(0.5)
                            st.rerun()

                    st.divider()
                    c1, c2 = st.columns(2)
                    with c1:
                        current_hall = row['hall'] if row['hall'] and str(row['hall']).lower() != 'nan' else ""
                        sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()),
                                            index=(list(hall_map.keys()).index(current_hall)+1 if current_hall in hall_map else 0),
                                            key=f"q_h_{st.session_state.system_mode}_{row['id']}_{idx}")
                        sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"],
                                            index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0),
                                            key=f"q_r_{st.session_state.system_mode}_{row['id']}_{idx}")
                    with c2:
                        if st.button("💾 حفظ التكليف", key=f"btn_save_{st.session_state.system_mode}_{row['id']}_{idx}"):
                            if row['preference'] == 'لا يرغب':
                                st.error("⚠️ هذا المعلم لا يرغب في التكليف، يرجى تغيير حالته أولاً")
                            elif row['ability'] == 'لا يصلح':
                                st.error("⚠️ هذا المعلم لا يصلح للمراقبة، يرجى تغيير حالته أولاً")
                            else:
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
                                    st.download_button("📥 تحميل الآن", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_s_{st.session_state.system_mode}_{row['id']}")

    # ==================== تبويب التوزيع التلقائي ====================
    with tab_auto:
        st.markdown('<h2 class="move-to-right">🤖 نظام التوزيع التلقائي الذكي</h2>', unsafe_allow_html=True)
        df_all = get_cached_teachers()
        hall_map_auto = {r['hall_name']: r['city'] for _, r in get_cached_halls().iterrows()}
        
        df_qualified = df_all[(df_all['ability'] == 'يصلح') & (df_all['preference'] == 'يرغب') & (df_all['current_job'] == 'معلم') & ((df_all['hall'] == '') | (df_all['hall'].isna()))]
        can_and_wants = len(df_qualified)
        can_not_wants = len(df_all[(df_all['ability'] == 'يصلح') & (df_all['preference'] == 'لا يرغب') & (df_all['current_job'] == 'معلم') & ((df_all['hall'] == '') | (df_all['hall'].isna()))])
        
        st.markdown(f"""
        <div style="display: flex; gap: 15px; margin-bottom: 20px; direction: rtl;">
            <div class="stat-card stat-wants"><span style="color: #bbb; font-size: 0.9rem;">متاح (يصلح ويرغب)</span><br><strong style="font-size: 2rem; color: #28a745;">{can_and_wants}</strong></div>
            <div class="stat-card stat-no-wants"><span style="color: #bbb; font-size: 0.9rem;">متاح (يصلح ولا يرغب)</span><br><strong style="font-size: 2rem; color: #dc3545;">{can_not_wants}</strong></div>
        </div>
        """, unsafe_allow_html=True)

        available_cities = sorted(df_qualified['city'].unique().tolist()) if not df_qualified.empty else []
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            target_h = st.selectbox("اختر القاعة المستهدفة:", [""] + list(hall_map_auto.keys()), key="auto_target_h")
            selected_cities = st.multiselect("السحب من مناطق سكن محددة (اختياري):", available_cities)
        with col_a2:
            df_pool = df_qualified[df_qualified['city'].isin(selected_cities)] if selected_cities else df_qualified
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

        st.divider()
        st.markdown('<h3 class="move-to-right">👔 تعيين رئيس القاعة والمساعد والآذن</h3>', unsafe_allow_html=True)
        df_managers = df_all[(df_all['current_job'] == 'مدير مدرسة') & (df_all['preference'] == 'يرغب') & ((df_all['hall'].isna()) | (df_all['hall'].astype(str).str.strip().isin(['', 'nan', 'None', 'NaN'])))]
        df_secretaries = df_all[(df_all['current_job'] == 'سكرتير') & (df_all['preference'] == 'يرغب') & ((df_all['hall'].isna()) | (df_all['hall'].astype(str).str.strip().isin(['', 'nan', 'None', 'NaN'])))]
        df_janitors = df_all[(df_all['current_job'] == 'آذن') & (df_all['preference'] == 'يرغب') & ((df_all['hall'].isna()) | (df_all['hall'].astype(str).str.strip().isin(['', 'nan', 'None', 'NaN'])))]

        target_h2 = ""
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            target_h2 = st.selectbox("اختر القاعة:", [""] + list(hall_map_auto.keys()), key="role_target_h")
        with col_r2:
            st.info(f"مدراء متاحين: {len(df_managers)} | سكرتارية: {len(df_secretaries)} | آذنة: {len(df_janitors)}")

        if target_h2:
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                sel_manager = st.selectbox("👑 رئيس القاعة (مدير مدرسة):", [""] + df_managers['name'].tolist(), key="sel_manager")
            with col_s2:
                sel_secretary = st.selectbox("📋 مساعد الرئيس (سكرتير):", [""] + df_secretaries['name'].tolist(), key="sel_secretary")
            with col_s3:
                sel_janitor = st.selectbox("🔑 الآذن:", [""] + df_janitors['name'].tolist(), key="sel_janitor")
            
            if st.button("💾 حفظ التعيينات", use_container_width=True, key="save_roles"):
                saved = []
                if sel_manager:
                    manager_id = df_managers[df_managers['name'] == sel_manager]['id'].values[0]
                    c.execute("UPDATE teachers SET hall=?, role='رئيس قاعة', hall_city=?, updated_by=? WHERE id=?", (target_h2, hall_map_auto[target_h2], st.session_state.username, manager_id))
                    saved.append(f"رئيس قاعة: {sel_manager}")
                if sel_secretary:
                    secretary_id = df_secretaries[df_secretaries['name'] == sel_secretary]['id'].values[0]
                    c.execute("UPDATE teachers SET hall=?, role='مساعد رئيس قاعة', hall_city=?, updated_by=? WHERE id=?", (target_h2, hall_map_auto[target_h2], st.session_state.username, secretary_id))
                    saved.append(f"مساعد رئيس: {sel_secretary}")
                if sel_janitor:
                    janitor_id = df_janitors[df_janitors['name'] == sel_janitor]['id'].values[0]
                    c.execute("UPDATE teachers SET hall=?, role='آذن', hall_city=?, updated_by=? WHERE id=?", (target_h2, hall_map_auto[target_h2], st.session_state.username, janitor_id))
                    saved.append(f"آذن: {sel_janitor}")
                if saved:
                    conn.commit()
                    add_log("تعيين أدوار", f"قاعة {target_h2}: {' | '.join(saved)}")
                    st.success(f"✅ تم الحفظ: {' | '.join(saved)}")
                    time.sleep(0.5)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.warning("⚠️ لم تختر أي شخص!")

    # ==================== تبويب رفع البيانات ====================
    with tab_upload:
        st.markdown(f'<h2 class="move-to-right">تحديث القالب والبيانات - {PAGE_TITLE}</h2>', unsafe_allow_html=True)
        up_tpl = st.file_uploader(f"ارفع قالب الوورد ({TEMPLATE_NAME})", type="docx")
        if up_tpl:
            with open(TEMPLATE_NAME, "wb") as f:
                f.write(up_tpl.getbuffer())
            add_log("تحديث قالب", f"تم رفع قالب {TEMPLATE_NAME} جديد")
            st.success("تم تحديث قالب الوورد بنجاح")
        
        st.divider()
        if st.button("🗑️ مسح البيانات المكررة"):
            try:
                c.execute("DELETE FROM teachers WHERE rowid NOT IN (SELECT MIN(rowid) FROM teachers GROUP BY id)")
                conn.commit()
                st.cache_data.clear()
                st.success("✅ تم مسح التكرار بنجاح")
                st.rerun()
            except Exception as e:
                st.error(f"خطأ: {e}")

        if st.button("🔄 تحديث من Google Sheets"):
            try:
                conn.execute("PRAGMA journal_mode=WAL")
                conn.commit()
                dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
                dft.columns = dft.columns.str.strip().str.lower()
                if 'id_number' in dft.columns:
                    dft.rename(columns={'id_number': 'id'}, inplace=True)
                dft.to_sql('teachers_temp', conn, if_exists='replace', index=False)
                ids_in_sheet = dft['id'].astype(str).tolist()
                placeholders = ','.join(['?' for _ in ids_in_sheet])
                c.execute(f"DELETE FROM teachers WHERE id NOT IN ({placeholders})", ids_in_sheet)
                conn.commit()
                
                if st.session_state['system_mode'] == 'tawjihi':
                    c.execute("UPDATE teachers SET name = t.name, phone = t.phone, school = t.school, city = t.city, current_job = t.current_job, preference = t.preference, ability = t.ability FROM teachers_temp t WHERE teachers.id = t.id")
                    c.execute("INSERT OR IGNORE INTO teachers (id, name, phone, school, city, current_job, preference, ability) SELECT id, name, phone, school, city, current_job, preference, ability FROM teachers_temp")
                else:
                    c.execute("UPDATE teachers SET name = t.name, phone = t.phone, school = t.school, city = t.city, current_job = t.current_job, preference = t.preference, ability = t.ability, relative = t.relative, relative_exam = t.relative_exam FROM teachers_temp t WHERE teachers.id = t.id")
                    c.execute("INSERT OR IGNORE INTO teachers (id, name, phone, school, city, current_job, preference, ability, relative, relative_exam) SELECT id, name, phone, school, city, current_job, preference, ability, relative, relative_exam FROM teachers_temp")
                conn.commit()
                dfh = pd.read_csv(HALLS_URL)
                dfh.to_sql('halls', conn, if_exists='replace', index=False)
                conn.commit()
                add_log("تحديث بيانات", "تحديث ذكي من جوجل شيت (حفظ التكليفات)")
                st.success("✅ تم التحديث بنجاح مع الحفاظ على التكليفات الحالية")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"خطأ أثناء التحديث: {e}")

    # ==================== تبويب الإدارة ====================
    with tab_manage:
        df_all_teachers = get_cached_teachers()
        total_count = len(df_all_teachers[(df_all_teachers['ability'] == 'يصلح') & (df_all_teachers['preference'] == 'يرغب') & (df_all_teachers['current_job'] == 'معلم')])
        assigned_count = len(df_all_teachers[(df_all_teachers['ability'] == 'يصلح') & (df_all_teachers['preference'] == 'يرغب') & (df_all_teachers['current_job'] == 'معلم') & (df_all_teachers['hall'].astype(str).str.len() > 0)])
        remaining_count = total_count - assigned_count
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("إجمالي الموظفين المتاحين للمراقبة", total_count)
        c_m2.metric("تم إنجازهم", assigned_count)
        c_m3.metric("المتبقي", remaining_count)
        
        st.divider()
        st.markdown('<h3 class="move-to-right">📦 تصدير البيانات المعدلة</h3>', unsafe_allow_html=True)
        df_export = df_all_teachers.copy()
        original_order = ['id', 'name', 'phone', 'school', 'city', 'role', 'hall', 'hall_city', 'preference', 'modified_by', 'job_title', 'permissions', 'relative', 'relative_exam']
        existing_cols = [c for c in original_order if c in df_export.columns]
        df_final = df_export[existing_cols].copy()
        column_mapping = {'id': 'رقم الهوية', 'name': 'الاسم كامل', 'phone': 'رقم الجوال', 'school': 'المدرسة', 'city': 'السكن', 'role': 'المهمة المكلف بها', 'hall': 'القاعة', 'hall_city': 'مدينة القاعة', 'preference': 'الرغبة', 'modified_by': 'الموظف المعدل', 'job_title': 'الوظيفة', 'permissions': 'الصلاحية', 'relative': 'قريب مباشر', 'relative_exam': 'امتحان القريب'}
        df_final.rename(columns=column_mapping, inplace=True)

        output_all = io.BytesIO()
        with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='الموظفين')
            workbook = writer.book
            worksheet = writer.sheets['الموظفين']
            h_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'align': 'center', 'bg_color': '#D7E4BC'})
            c_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'align': 'right'})
            worksheet.right_to_left()
            worksheet.set_landscape()
            worksheet.fit_to_pages(1, 0)
            for col_num, col_name in enumerate(df_final.columns):
                worksheet.write(0, col_num, col_name, h_fmt)
                column_data = df_final[col_name].astype(str).str.len()
                max_len = max(column_data.max() if not column_data.empty else 0, len(str(col_name))) + 7
                worksheet.set_column(col_num, col_num, min(max_len, 50), c_fmt)

        st.download_button(label="📥 تحميل إكسل معدل", data=output_all.getvalue(), file_name=f"كشف_عام_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                    styled_df = df_to_show.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).hide(axis="index")
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
                        df_final_export.to_excel(writer, index=False, sheet_name='كشف_القاعة', startrow=1)
                        workbook = writer.book
                        worksheet = writer.sheets['كشف_القاعة']
                        title_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#BDD7EE'})
                        h_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#BDD7EE'})
                        c_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'align': 'right', 'valign': 'vcenter'})
                        worksheet.right_to_left()
                        worksheet.set_landscape()
                        worksheet.fit_to_pages(1, 0)
                        header_text = f"بيانات قاعة: {h_choice}"
                        worksheet.merge_range(0, 0, 0, 6, header_text, title_fmt)
                        worksheet.set_row(0, 35)
                        for col_num, col_name in enumerate(df_final_export.columns):
                            worksheet.write(1, col_num, col_name, h_fmt)
                            column_length = max(df_final_export[col_name].astype(str).map(len).max(), len(str(col_name))) + 4
                            worksheet.set_column(col_num, col_num, column_length, c_fmt)
                    add_log("تصدير إكسل", f"تحميل كشف قاعة: {h_choice}")
                    st.download_button(label=f"📊 كشف إكسل {h_choice}", data=output_hall_excel.getvalue(), file_name=f"كشف_{h_choice}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_xl_{h_choice}_export")

    # ==================== تبويب السجلات ====================
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

# ============================================================================
# ✨✨✨ نظام تصحيح الثانوية العامة - وحدة مستقلة تماماً ✨✨✨
# ============================================================================
# 🔴 هذا القسم يظهر فقط إذا كان الوضع == "tasheeh"

if st.session_state.get('system_mode') == "tasheeh":
    
    # 1️⃣ إنشاء جداول التخزين الدائم في قاعدة البيانات
    c.execute('''CREATE TABLE IF NOT EXISTS tasheeh_teachers (
        id TEXT PRIMARY KEY, name TEXT, subject TEXT, city TEXT, 
        school TEXT, phone TEXT, relative TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS tasheeh_halls (
        hall_name TEXT PRIMARY KEY, city TEXT
    )''')
    conn.commit()
    
    # 2️⃣ تحميل البيانات تلقائياً من القاعدة عند فتح الموقع
    if 'tasheeh_teachers' not in st.session_state:
        try:
            st.session_state['tasheeh_teachers'] = pd.read_sql("SELECT * FROM tasheeh_teachers", conn)
        except: st.session_state['tasheeh_teachers'] = pd.DataFrame()
    if 'tasheeh_halls' not in st.session_state:
        try:
            st.session_state['tasheeh_halls'] = pd.read_sql("SELECT * FROM tasheeh_halls", conn)
        except: st.session_state['tasheeh_halls'] = pd.DataFrame()

    # 3️⃣ دالة المزامنة الذكية (تحديث + إضافة بدون تكرار)
    def sync_tasheeh_data():
        try:
            with st.spinner("🔄 جاري المزامنة الذكية من Google Sheets..."):
                # مزامنة المعلمين
                df_t = pd.read_csv(TEACHERS_URL, dtype=str)
                df_t.columns = df_t.columns.str.strip().str.lower()
                rename_map = {
                    'رقم الهوية': 'id', 'الاسم': 'name', 'المبحث': 'subject',
                    'مكان سكن المعلم': 'city', 'اسم المدرسة': 'school', 
                    'رقم جواله': 'phone', 'هل له قريب مباشر او لا': 'relative'
                }
                df_t = df_t.rename(columns={k:v for k,v in rename_map.items() if k in df_t.columns})
                
                for _, r in df_t.iterrows():
                    tid = str(r.get('id','')).strip()
                    if not tid: continue # تخطي الصفوف الفارغة
                    c.execute("""INSERT OR REPLACE INTO tasheeh_teachers 
                                 (id, name, subject, city, school, phone, relative) 
                                 VALUES (?,?,?,?,?,?,?)""",
                              (tid, str(r.get('name','')), str(r.get('subject','')),
                               str(r.get('city','')), str(r.get('school','')), 
                               str(r.get('phone','')), str(r.get('relative',''))))
                conn.commit()
                
                # مزامنة القاعات
                df_h = pd.read_csv(HALLS_URL, dtype=str)
                df_h.columns = df_h.columns.str.strip().str.upper()
                for _, r in df_h.iterrows():
                    hname = str(r.get('ZHALL','')).strip()
                    if not hname: continue
                    c.execute("INSERT OR REPLACE INTO tasheeh_halls (hall_name, city) VALUES (?,?)",
                              (hname, str(r.get('ZLOC',''))))
                conn.commit()
                
                # تحديث الجلسة والواجهة
                st.session_state['tasheeh_teachers'] = pd.read_sql("SELECT * FROM tasheeh_teachers", conn)
                st.session_state['tasheeh_halls'] = pd.read_sql("SELECT * FROM tasheeh_halls", conn)
            st.success("✅ تم التحديث الذكي بنجاح! (تم حفظ/تحديث البيانات بدون تكرار)")
            st.rerun()
        except Exception as e:
            st.error(f"❌ خطأ أثناء المزامنة: {e}")

    def generate_tasheeh_letter(data, exam_name):
        if not os.path.exists(TEMPLATE_NAME):
            return None
        doc = Document(TEMPLATE_NAME)
        repls = {
            'ZNAME': data.get('name', '---'), 'ZID': data.get('id', '---'),
            'ZTEST': exam_name, 'ZHALL': data.get('hall_name', '---'),
            'ZLOC': data.get('hall_city', '---'), 'ZWORK': data.get('school', '---'),
            'ZCITY': data.get('city', '---')
        }
        for p in doc.paragraphs:
            for k, v in repls.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text:
                            run.text = run.text.replace(k, str(v))
                            run.bold = True
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for k, v in repls.items():
                            if k in p.text:
                                for run in p.runs:
                                    if k in run.text:
                                        run.text = run.text.replace(k, str(v))
                                        run.bold = True
        return doc
    
    st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a1c23 0%, #2d3748 100%); 
                    padding: 20px; border-radius: 15px; border: 2px solid #00ffcc;
                    margin: 20px 0; text-align: center;">
            <h2 style="color: #00ffcc; margin: 0;">✨ نظام تصحيح الثانوية العامة ✨</h2>
            <p style="color: #bbb; margin: 10px 0 0 0;">توزيع المصححين حسب المبحث والقاعة</p>
        </div>
    """, unsafe_allow_html=True)
    
    corr_tab1, corr_tab2, corr_tab3, corr_tab4 = st.tabs([
        "📥 رفع البيانات", "🔄 التوزيع التلقائي", "📄 كتب التكليف", "📜 سجل العمليات"
    ])
    
    # ==================== تبويب 1: رفع البيانات والمزامنة ====================
    with corr_tab1:
        st.markdown("### 📥 إدارة البيانات وقالب التكليف")
        
        st.markdown("**1️⃣ رفع قالب وورد التصحيح**")
        st.caption("يجب أن يحتوي القالب على الرموز: ZNAME, ZID, ZHALL, ZLOC, ZTEST, ZWORK, ZCITY")
        
        uploaded_tasheeh_tpl = st.file_uploader(
            "📄 اختر ملف القالب (template_tasheeh.docx)", 
            type="docx", 
            key="tasheeh_tpl_uploader_unique"
        )
        if uploaded_tasheeh_tpl is not None:
            try:
                with open(TEMPLATE_NAME, "wb") as f:
                    f.write(uploaded_tasheeh_tpl.getbuffer())
                st.success("✅ تم حفظ القالب بنجاح!")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"❌ خطأ: {e}")
        
        st.divider()
        st.markdown("**2️⃣ المزامنة مع Google Sheets**")
        st.info("💡 البيانات محفوظة تلقائياً في النظام. اضغط هنا فقط إذا أضفت/عدلت بيانات في ملف الإكسل الخارجي.")
        
        if st.button("🔄 مزامنة وتحديث البيانات من Google Sheets", type="primary", use_container_width=True):
            sync_tasheeh_data()
            
        if not st.session_state['tasheeh_teachers'].empty:
            st.markdown(f"📊 **عدد المعلمين المخزنين حالياً:** `{len(st.session_state['tasheeh_teachers'])}`")
            st.dataframe(st.session_state['tasheeh_teachers'].head(), use_container_width=True)
        st.divider()
        st.markdown("### 🧹 تنظيف البيانات من التكرار")
        if st.button("🗑️ حذف المكررات (نفس الاسم والهوية)", type="secondary", use_container_width=True):
            try:
                c.execute("""
                    DELETE FROM tasheeh_teachers 
                    WHERE rowid NOT IN (
                        SELECT MIN(rowid) 
                        FROM tasheeh_teachers 
                        GROUP BY id
                    )
                """)
                conn.commit()
                st.cache_data.clear()
                # تحديث البيانات في الذاكرة فوراً
                st.session_state['tasheeh_teachers'] = pd.read_sql("SELECT * FROM tasheeh_teachers", conn)
                st.success("✅ تم حذف التكرارات بنجاح!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ خطأ: {e}")

    # ==================== تبويب 2: التوزيع التلقائي ====================
        # ==================== تبويب 2: التوزيع التلقائي ====================
        # ==================== تبويب 2: التوزيع التلقائي ====================
        # ==================== تبويب 2: التوزيع التلقائي (المعدل بالتحكم اليدوي) ====================
    with corr_tab2:
        if st.session_state['tasheeh_teachers'].empty or st.session_state['tasheeh_halls'].empty:
            st.warning("⚠️ يرجى مزامنة البيانات أولاً من تبويب 'رفع البيانات'")
        else:
            teachers = st.session_state['tasheeh_teachers']
            halls = st.session_state['tasheeh_halls']
            
            st.markdown("### ⚙️ إعدادات التوزيع")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                # 1️⃣ اختيار القاعة من بيانات الإكسل
                hall_opts = sorted(halls['hall_name'].unique().tolist())
                selected_hall = st.selectbox("🏫 اختر القاعة:", hall_opts, index=0)
                
            with col2:
                # 2️⃣ اختيار المادة من بيانات المعلمين
                subj_opts = sorted(teachers['subject'].dropna().unique().tolist())
                selected_subj = st.selectbox("📚 اختر المادة:", subj_opts, index=0)
                
            with col3:
                # 3️⃣ اختيار التاريخ يدوياً
                default_date = datetime(2026, 6, 20)
                selected_date = st.date_input("📅 تاريخ بداية التصحيح:", value=default_date, min_value=datetime(2026, 1, 1))
            
            # تنسيق التاريخ واليوم
            date_str = selected_date.strftime("%Y/%m/%d")
            days_ar = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
            day_name = days_ar[selected_date.weekday()]
            
            # عرض ملخص سريع قبل التوزيع
            hall_row = halls[halls['hall_name'] == selected_hall]
            hall_city = hall_row['city'].values[0] if not hall_row.empty else "غير محدد"
            st.info(f"📍 القاعة: `{selected_hall}` | 🌆 المدينة: `{hall_city}` | 📅 التاريخ: `{date_str} ({day_name})`")

            if st.button("🚀 توزيع المصححين", type="primary", use_container_width=True):
                # تصفية المعلمين (استبعاد من له قريب + تصفية حسب المادة)
                teachers_filtered = teachers[teachers['relative'].astype(str).str.lower() != 'true']
                pool = teachers_filtered[teachers_filtered['subject'] == selected_subj]
                
                if pool.empty:
                    st.warning(f"⚠️ لا يوجد معلمين متاحين للمادة: `{selected_subj}`")
                else:
                    assignments = []
                    for _, t in pool.iterrows():
                        assignments.append({
                            'id': t.get('id',''), 
                            'name': t.get('name',''), 
                            'subject': t.get('subject',''),
                            'hall_name': selected_hall,      # القاعة المختارة
                            'hall_city': hall_city,          # المدينة المجلوبة تلقائياً من الإكسل
                            'exam_name': selected_subj,      # اسم المادة هو اسم الامتحان
                            'exam_date': date_str,           # التاريخ المختار
                            'exam_day': day_name,            # اليوم المحسوب تلقائياً
                            'school': t.get('school',''), 
                            'city': t.get('city',''),
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M")
                        })
                    
                    if assignments:
                        st.session_state['tasheeh_assignments'] = assignments
                        add_log("توزيع تصحيح", f"توزيع {len(assignments)} معلم - {selected_subj} في {selected_hall} بتاريخ {date_str}")
                        st.success(f"✅ تم توزيع {len(assignments)} معلم بنجاح!")
                        
                        for a in assignments:
                            try:
                                c.execute("""INSERT INTO tasheeh_assignments 
                                             (teacher_id,teacher_name,subject,hall_name,hall_city,exam_name,
                                              exam_date,exam_day,created_at,created_by) 
                                             VALUES (?,?,?,?,?,?,?,?,?,?)""",
                                         (a['id'],a['name'],a['subject'],a['hall_name'],a['hall_city'],a['exam_name'],
                                          a['exam_date'],a['exam_day'],a['timestamp'],st.session_state.username))
                            except: pass
                        conn.commit()
                    else:
                        st.error("❌ لم يتم التوزيع.")
    
    # ==================== تبويب 3: كتب التكليف ====================
        # ==================== تبويب 3: كتب التكليف وإدارتها ====================
        # ==================== تبويب 3: إحصائيات التكليفات والكتب ====================
    with corr_tab3:
        
        # التحقق من وجود البيانات
        if 'tasheeh_teachers' not in st.session_state or st.session_state['tasheeh_teachers'].empty:
            st.warning("⚠️ يرجى تحميل البيانات أولاً من تبويب 'رفع البيانات' لعرض الإحصائيات.")
        else:
            st.markdown("### 📊 إحصائيات التكليفات")
            
            # 1. قائمة المواد المتاحة من بيانات المعلمين
            teachers_df = st.session_state['tasheeh_teachers']
            subjects_list = sorted(teachers_df['subject'].dropna().unique().tolist()) if not teachers_df.empty else []
            
            col_f1, col_f2 = st.columns([1, 2])
            with col_f1:
                filter_subj = st.selectbox("🔍 عرض إحصائيات مادة محددة:", ["الكل"] + subjects_list, index=0)
            
            # 2. فلترة البيانات للحسابات
            assignments_list = st.session_state.get('tasheeh_assignments', [])
            df_assigns = pd.DataFrame(assignments_list) if assignments_list else pd.DataFrame()
            
            if filter_subj != "الكل":
                # فلترة حسب المادة المختارة
                pool_count = len(teachers_df[teachers_df['subject'] == filter_subj])
                assigned_df = df_assigns[df_assigns['subject'] == filter_subj] if not df_assigns.empty else pd.DataFrame()
                assigned_count = len(assigned_df)
                remaining_count = pool_count - assigned_count
            else:
                # عرض الكل
                pool_count = len(teachers_df)
                assigned_count = len(df_assigns)
                assigned_df = df_assigns
                remaining_count = pool_count - assigned_count

            # 3. عرض المقاييس (Metrics)
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                st.metric("📚 إجمالي المعلمين (المادة)", pool_count)
            with c_m2:
                st.metric("✅ تم تكليفهم", assigned_count)
            with c_m3:
                st.metric("⏳ المتبقي للتكليف", remaining_count)

            st.divider()

            # 4. عرض الجدول
            st.markdown(f"### 📋 القائمة الحالية: {filter_subj}")
            if not assigned_df.empty:
                display_cols = ['name', 'subject', 'hall_name', 'hall_city']
                # التأكد من وجود الأعمدة
                safe_cols = [c for c in display_cols if c in assigned_df.columns]
                st.dataframe(assigned_df[safe_cols], use_container_width=True)
            else:
                st.info(f"لا يوجد تكليفات لـ {filter_subj}.")

            st.divider()

            # 5. أزرار التحكم (تحميل وحذف)
            st.markdown("### ⚙️ إدارة وتصدير")
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if st.button("📥 تحميل وورد للمادة الحالية", type="primary", use_container_width=True, disabled=assigned_df.empty):
                    if not os.path.exists(TEMPLATE_NAME):
                        st.error("❌ القالب غير موجود")
                    else:
                        with st.spinner("جاري الإنشاء..."):
                            try:
                                final_doc = Document(TEMPLATE_NAME)
                                final_doc._body.clear_content()
                                current_list = assigned_df.to_dict('records') 
                                
                                for i, a in enumerate(current_list):
                                    temp_doc = Document(TEMPLATE_NAME)
                                    repls = {
                                        'ZNAME': str(a.get('name', '---')), 
                                        'ZID': str(a.get('id', '---')),
                                        'ZTEST': str(a.get('exam_name', '---')), 
                                        'ZHALL': str(a.get('hall_name', '---')),
                                        'ZLOC': str(a.get('hall_city', '---')), 
                                        'ZWORK': str(a.get('subject', '---')), 
                                        'ZCITY': str(a.get('city', '---')),
                                        'ZSUBJECT': str(a.get('subject', '---'))
                                    }
                                    for p in temp_doc.paragraphs:
                                        for k, v in repls.items():
                                            if k in p.text:
                                                for run in p.runs:
                                                    if k in run.text:
                                                        run.text = run.text.replace(k, v)
                                                        run.bold = True
                                    for table in temp_doc.tables:
                                        for row in table.rows:
                                            for cell in row.cells:
                                                for p in cell.paragraphs:
                                                    for k, v in repls.items():
                                                        if k in p.text:
                                                            for run in p.runs:
                                                                if k in run.text:
                                                                    run.text = run.text.replace(k, v)
                                                                    run.bold = True
                                                    
                                    elements = [el for el in temp_doc.element.body if not el.tag.endswith('sectPr')]
                                    for element in elements:
                                        final_doc.element.body.append(copy.deepcopy(element))
                                    
                                    if i < len(current_list) - 1:
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
                                st.success(f"✅ تم إنشاء ملف {filter_subj} بنجاح")
                                st.download_button(
                                    label="📥 تحميل الآن",
                                    data=out.getvalue(),
                                    file_name=f"تكليفات_{filter_subj}_{datetime.now().strftime('%Y%m%d')}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key="dl_word_filtered_tasheeh"
                                )
                            except Exception as e:
                                st.error(f"خطأ: {e}")

            with col_btn2:
                del_btn_label = f"🗑️ حذف تكليفات {filter_subj}"
                if filter_subj == "الكل":
                     del_btn_label = "🗑️ حذف جميع التكليفات"
                
                if st.button(del_btn_label, type="secondary", use_container_width=True, disabled=assigned_df.empty):
                    if filter_subj == "الكل":
                        c.execute("DELETE FROM tasheeh_assignments")
                    else:
                        c.execute("DELETE FROM tasheeh_assignments WHERE subject=?", (filter_subj,))
                    conn.commit()
                    
                    if filter_subj == "الكل":
                        st.session_state['tasheeh_assignments'] = []
                    else:
                        st.session_state['tasheeh_assignments'] = [a for a in st.session_state['tasheeh_assignments'] if a['subject'] != filter_subj]
                    
                    st.success(f"✅ تم حذف تكليفات {filter_subj}")
                    st.rerun()

            with col_btn2:
                if st.button("📊 تصدير إكسل"):
                    df = pd.DataFrame(st.session_state['tasheeh_assignments'])
                    out = io.BytesIO()
                    df.to_excel(out, index=False)
                    out.seek(0)
                    st.download_button(
                        "📥 تحميل إكسل", 
                        out.getvalue(), 
                        f"تصحيح_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_excel_tasheeh_unique"
                    )
            
            st.divider()
            if st.button("📊 تصدير كملف إكسل"):
                df = pd.DataFrame(assigns)
                out = io.BytesIO()
                df.to_excel(out, index=False)
                out.seek(0)
                st.download_button("📥 تحميل إكسل", out.getvalue(), 
                                   f"تصحيح_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_excel_tasheeh_unique")
    
    # ==================== تبويب 4: سجل العمليات ====================
        # ==================== تبويب 4: سجل العمليات ====================
    with corr_tab4:
        st.markdown("### 📜 سجل العمليات الخاص بالتصحيح")
        
        # 🔴🔴 زر حذف السجلات الجديد 🔴🔴
        st.warning("⚠️ هذا الزر سيقوم بحذف سجلات التصحيح نهائياً.")
        if st.button("🗑️ حذف سجلات التصحيح نهائياً", type="primary", key="delete_tasheeh_logs_btn"):
            try:
                # حذف السجلات التي تحتوي على كلمة 'تصحيح' فقط
                c.execute("DELETE FROM logs WHERE action LIKE '%تصحيح%'")
                conn.commit()
                st.success("✅ تم مسح سجلات التصحيح بالكامل")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"❌ خطأ أثناء الحذف: {e}")
        
        st.divider()
        
        # عرض السجلات
        df = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs WHERE action LIKE '%تصحيح%' ORDER BY id DESC", conn)
        if not df.empty:
            st.dataframe(df, use_container_width=True)
        else:
            st.info("لا يوجد سجلات تصحيح حالياً.")
    
    st.stop()
