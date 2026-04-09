import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Mm
import io
import os
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
    [data-testid="stSidebar"] { direction: rtl; text-align: right; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; direction: rtl; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    [data-testid="stMetricValue"] { font-size: 1.6rem !important; color: #00ffcc !important; }
    .city-card {
        background-color: #1a1c23;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #3b3e4a;
        margin-bottom: 10px;
        text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.3);
    }
    .right-align { text-align: right; direction: rtl; width: 100%; margin-bottom: 20px; }
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
# 3. وظائف معالجة الملفات (تم إصلاح مشكلة الصفحات الفارغة في البداية)
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9: phone_val = '0' + phone_val
    
    repls = {
        'ZNAME': str(row.get('name', '')), 
        'ZID': str(row.get('id', '')), 
        'ZJOB': str(row.get('role', '')), 
        'ZHALL': str(h_name), 
        'ZLOC': str(h_city), 
        'ZWORK': str(row.get('school', '')), 
        'ZCITY': str(row.get('city', '')),
        'ZPHONE': phone_val
    }
    
    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            for k, v in repls.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text:
                            run.text = run.text.replace(k, v)

    replace_in_paragraphs(doc_obj.paragraphs)
    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                replace_in_paragraphs(cell.paragraphs)
    
    # حذف الأسطر الفارغة في نهاية المستند
    while len(doc_obj.paragraphs) > 0:
        last_p = doc_obj.paragraphs[-1]
        if not last_p.text.strip():
            p_element = last_p._element
            p_element.getparent().remove(p_element)
        else:
            break
            
    return doc_obj

def generate_bulk_word(df, h_name):
    if not os.path.exists("template.docx"): return None
    
    # إنشاء مستند جديد نظيف تماماً لتجنب الصفحات الفارغة في البداية
    final_doc = Document()
    
    # ضبط هوامش المستند الجديد لتطابق القالب (0.5 سم)
    for section in final_doc.sections:
        section.top_margin = Mm(5)
        section.bottom_margin = Mm(5)
        section.left_margin = Mm(10)
        section.right_margin = Mm(10)

    df_clean = df.reset_index(drop=True)
    
    for idx, row in df_clean.iterrows():
        temp_doc = Document("template.docx")
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        
        # استخراج المحتوى بدون خصائص القسم التي تسبب الفراغات
        elements = [el for el in temp_doc.element.body if not el.tag.endswith('sectPr')]
        
        # تنظيف الفقرات الفارغة في نهاية كل تكليف
        while elements and elements[-1].tag.endswith('p'):
            p_text = "".join(elements[-1].itertext()).strip()
            if not p_text:
                elements.pop()
            else:
                break
                
        for element in elements:
            final_doc.element.body.append(element)
            
        # إضافة فاصل صفحات فقط "بين" التكليفات
        if idx < len(df_clean) - 1:
            final_doc.add_page_break()
            
    out = io.BytesIO()
    final_doc.save(out)
    out.seek(0)
    return out

def generate_single_doc(row):
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# =====================================
# 4. الشريط الجانبي (الموظف وتسجيل الخروج)
# =====================================
st.sidebar.markdown(f"### 👤 الموظف الحالي:")
st.sidebar.info(f"**{st.session_state.username}**")
if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()
st.sidebar.divider()

# =====================================
# 5. التبويبات الرئيسية
# =====================================
tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs([
    "🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"
])

with tab_search:
    st.markdown("<div class='right-align'><h3>🔍 إدارة الموظفين والتعيين اليدوي</h3></div>", unsafe_allow_html=True)
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                st.markdown(f"🆔 الهوية: {row['id']} | 📱 الجوال: {row['phone']} | 🏫 المدرسة: {row['school']}")
                with st.popover("📝 تعديل البيانات الأساسية"):
                    with st.form(key=f"edit_base_{row['id']}"):
                        u_name = st.text_input("الاسم", value=row['name'])
                        u_job = st.text_input("الوظيفة الحالية", value=row['current_job'])
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2))
                        u_abil = st.selectbox("الصلاحية", ["يصلح", "لا يصلح", "لم تحدد"], index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2))
                        if st.form_submit_button("💾 حفظ التعديلات"):
                            c.execute("UPDATE teachers SET name=?, current_job=?, preference=?, ability=?, updated_by=? WHERE id=?", (u_name, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                            conn.commit(); add_log("تعديل بيانات", f"تعديل {u_name}"); st.success("تم التحديث"); st.rerun()
                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), key=f"sel_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], key=f"sel_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"save_btn_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        conn.commit(); add_log("تعيين يدوي", f"تكليف {row['name']} في {sel_h}"); st.success("تم الحفظ"); st.rerun()
                    if row['hall']:
                        if st.button("❌ إلغاء التكليف", key=f"del_btn_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit(); st.rerun()
                        f_word = generate_single_doc(row)
                        if f_word: st.download_button("📥 تحميل الكتاب", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_btn_{row['id']}")

with tab_auto:
    st.markdown("<div class='right-align'><h3>🤖 التوزيع التلقائي للمراقبين</h3></div>", unsafe_allow_html=True)
    df_h_data_auto = pd.read_sql("SELECT * FROM halls", conn)
    hall_map_auto = {r['hall_name']: r['city'] for _, r in df_h_data_auto.iterrows()}
    
    df_avail = pd.read_sql("SELECT * FROM teachers WHERE current_job = 'معلم' AND preference = 'يرغب' AND ability = 'يصلح' AND (hall = '' OR hall IS NULL OR hall = 'nan')", conn)
    df_not_willing = pd.read_sql("SELECT * FROM teachers WHERE current_job = 'معلم' AND preference = 'لا يرغب' AND ability = 'يصلح' AND (hall = '' OR hall IS NULL OR hall = 'nan')", conn)

    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_hall = st.selectbox("اختر القاعة المستهدفة:", [""] + list(hall_map_auto.keys()))
        selected_cities = st.multiselect("اختر مناطق السحب:", options=sorted(df_avail['city'].unique().tolist()) if not df_avail.empty else [])
    with col_a2:
        req_proctors = st.number_input("عدد المراقبين المطلوب:", min_value=1, value=10)
        if st.button("🚀 تنفيذ التوزيع العشوائي", use_container_width=True):
            if target_hall and selected_cities:
                pool = df_avail[df_avail['city'].isin(selected_cities)].sample(frac=1).reset_index(drop=True).head(req_proctors)
                if not pool.empty:
                    for _, t in pool.iterrows():
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (target_hall, "مراقب", hall_map_auto.get(target_hall, ""), st.session_state.username, t['id']))
                    conn.commit(); st.session_state.last_assigned_proctors = pool[['name', 'id', 'city', 'school']]; st.success("✅ تم التوزيع"); st.rerun()

    if 'last_assigned_proctors' in st.session_state and st.session_state.last_assigned_proctors is not None:
        st.divider()
        st.dataframe(st.session_state.last_assigned_proctors, use_container_width=True, hide_index=True)
        if st.button("🧹 إخفاء الجدول"): st.session_state.last_assigned_proctors = None; st.rerun()

    st.divider()
    with st.expander("✅ المعلمون المتاحون (الذين يرغبون) - توزيع سكني"):
        if not df_avail.empty:
            stats_yes = df_avail['city'].value_counts().reset_index()
            cols = st.columns(4)
            for i, r in stats_yes.iterrows():
                with cols[i % 4]:
                    st.markdown(f"<div class='city-card'><b>{r['city']}</b><br><span style='color:#00ffcc; font-size:1.3rem;'>{r['count']} معلم</span></div>", unsafe_allow_html=True)
        else: st.info("لا يوجد متاحون.")

    with st.expander("⚠️ المعلمون الاحتياط (الذين لا يرغبون) - توزيع سكني"):
        if not df_not_willing.empty:
            stats_no = df_not_willing['city'].value_counts().reset_index()
            cols = st.columns(4)
            for i, r in stats_no.iterrows():
                with cols[i % 4]:
                    st.markdown(f"<div class='city-card'><b>{r['city']}</b><br><span style='color:#ff4b4b; font-size:1.3rem;'>{r['count']} معلم</span></div>", unsafe_allow_html=True)
        else: st.info("القائمة فارغة.")

with tab_manage:
    st.markdown("<div class='right-align'><h3>📊 الإدارة والإحصائيات العامة</h3></div>", unsafe_allow_html=True)
    df_all = pd.read_sql("SELECT * FROM teachers", conn)
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي الكادر", len(df_all))
    c_m2.metric("المكلفون حالياً", len(df_all[df_all['hall'].astype(str).str.len() > 1]))
    c_m3.metric("غير المكلفين", len(df_all) - len(df_all[df_all['hall'].astype(str).str.len() > 1]))
    st.divider()
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != '' AND hall IS NOT NULL", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة لإدارتها:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                if st.button(f"🗑️ حذف مراقبي {h_choice}", use_container_width=True):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE hall=? AND role='مراقب'", (h_choice,))
                    conn.commit(); st.rerun()
            with col_m2:
                bulk_f = generate_bulk_word(df_hall_details, h_choice)
                if bulk_f: st.download_button("📥 تحميل كافة الكتب", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx", use_container_width=True)
            st.dataframe(df_hall_details[['name', 'role', 'school', 'city']], use_container_width=True)

with tab_upload:
    up_tpl = st.file_uploader("ارفع قالب وورد جديد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f: f.write(up_tpl.getbuffer())
        st.success("تم تحديث القالب")
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
            dft.columns = dft.columns.str.strip().str.lower()
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.to_sql('halls', conn, if_exists='replace', index=False)
            st.success("تمت المزامنة"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_logs:
    df_l = pd.read_sql("SELECT user, action, details, timestamp FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
