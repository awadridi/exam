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
    if phone_val.startswith('5') and len(phone_val) == 9:
        phone_val = '0' + phone_val
    h_name_final = str(h_name) if h_name and str(h_name).lower() != 'nan' else "---"
    h_city_final = str(h_city) if h_city and str(h_city).lower() != 'nan' else "---"
    repls = {
        '<NAME>': str(row.get('name', '')), 
        '<ID>': str(row.get('id', '')), 
        '<PHONE>': phone_val, 
        '<JOB>': str(row.get('role', '')), 
        '<HALL_NAME>': h_name_final, 
        '<HALL_LOCATION>': h_city_final, 
        '<WORKPLACE>': str(row.get('school', '')), 
        '<CITY>': str(row.get('city', ''))
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
st.sidebar.markdown(f"### 👤 الموظف: **{st.session_state.username}**")
if st.sidebar.button("🚪 خروج"):
    st.session_state.logged_in = False
    st.rerun()

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs([
    "🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"
])

with tab_search:
    st.subheader("إدارة الموظفين")
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            display_phone = str(row['phone'])
            if display_phone.startswith('5') and len(display_phone) == 9: display_phone = '0' + display_phone
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                st.markdown(f"**🆔 الهوية:** {row['id']} | **📱 الجوال:** {display_phone} | **🏡 السكن:** {row['city']}")
                with st.popover("📝 تعديل البيانات الأساسية"):
                    with st.form(key=f"edit_base_{row['id']}"):
                        u_name = st.text_input("الاسم", value=row['name'])
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2))
                        if st.form_submit_button("💾 تحديث وحفظ"):
                            c.execute("UPDATE teachers SET name=?, preference=?, updated_by=? WHERE id=?", (u_name, u_pref, st.session_state.username, row['id']))
                            conn.commit(); add_log("تعديل بيانات", f"تعديل {u_name}"); st.success("تم التحديث"); st.rerun()
                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        conn.commit(); add_log("حفظ تكليف", f"تكليف {row['name']}"); st.success("تم الحفظ"); st.rerun()
                    if row['hall']:
                        if st.button("❌ إلغاء التكليف", key=f"del_search_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", (st.session_state.username, row['id']))
                            conn.commit(); add_log("إلغاء تكليف", f"إلغاء {row['name']}"); st.rerun()

# =====================================
# 5. التوزيع التلقائي (للمراقبين فقط)
# =====================================
with tab_auto:
    st.subheader("🤖 التوزيع التلقائي للمراقبين")
    if 'last_assigned_proctors' not in st.session_state: st.session_state.last_assigned_proctors = None
    df_h_data_auto = pd.read_sql("SELECT * FROM halls", conn)
    hall_map_auto = {r['hall_name']: r['city'] for _, r in df_h_data_auto.iterrows()}
    df_avail = pd.read_sql("SELECT * FROM teachers WHERE current_job = 'معلم' AND preference = 'يرغب' AND ability = 'يصلح' AND (hall = '' OR hall IS NULL OR hall = 'nan')", conn)
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_hall = st.selectbox("اختر القاعة لتوزيع المراقبين عليها:", [""] + list(hall_map_auto.keys()), key="auto_hall_sel")
        selected_cities = st.multiselect("اختر مناطق السكن:", options=sorted(df_avail['city'].unique().tolist()))
    with col_a2:
        req_proctors = st.number_input("عدد المراقبين المطلوبين:", min_value=1, value=10)
        if st.button("🚀 تنفيذ التوزيع العشوائي", use_container_width=True):
            if target_hall and selected_cities:
                pool = df_avail[df_avail['city'].isin(selected_cities)].sample(frac=1).reset_index(drop=True).head(req_proctors)
                for _, t in pool.iterrows():
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (target_hall, "مراقب", hall_map_auto.get(target_hall, ""), st.session_state.username, t['id']))
                conn.commit(); st.session_state.last_assigned_proctors = pool[['name', 'id', 'city', 'school']]; add_log("توزيع تلقائي", f"توزيع {len(pool)} مراقب على {target_hall}"); st.success("تم التوزيع"); st.rerun()
    
    if st.session_state.last_assigned_proctors is not None:
        st.divider(); st.markdown(f"### 📋 كشف الموزعين حالياً على: {target_hall}"); st.dataframe(st.session_state.last_assigned_proctors, use_container_width=True, hide_index=True)

# =====================================
# 6. الإدارة والإحصائيات (إضافة زر الحذف الانتقائي)
# =====================================
with tab_manage:
    df_all_teachers = pd.read_sql("SELECT * FROM teachers", conn)
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي الموظفين", len(df_all_teachers))
    c_m2.metric("تم تكليفهم", len(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]))
    c_m3.metric("المتبقي", len(df_all_teachers) - len(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]))
    
    st.divider()
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != '' AND hall IS NOT NULL", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض والإدارة:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            st.markdown(f"##### 📊 كادر قاعة: {h_choice}")
            
            # --- ميزة حذف المراقبين فقط ---
            proctors_count = len(df_hall_details[df_hall_details['role'] == 'مراقب'])
            st.warning(f"يوجد حالياً {proctors_count} مراقب في هذه القاعة.")
            
            col_manage1, col_manage2 = st.columns(2)
            with col_manage1:
                if st.button(f"🗑️ حذف تكليف مراقبي {h_choice} فقط", type="secondary", use_container_width=True):
                    c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE hall=? AND role='مراقب'", (st.session_state.username, h_choice))
                    conn.commit()
                    add_log("حذف مراقبي قاعة", f"تم حذف جميع المراقبين من قاعة {h_choice}")
                    st.success(f"تم إزالة {proctors_count} مراقب بنجاح. (تم الإبقاء على رئيس القاعة والمساعدين)")
                    time.sleep(1)
                    st.rerun()
            
            with col_manage2:
                bulk_f = generate_bulk_word(df_hall_details, h_choice)
                if bulk_f: st.download_button(f"📄 تحميل كتب قاعة {h_choice}", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")
            
            st.dataframe(df_hall_details[['name', 'role', 'school', 'city']], use_container_width=True)

# =====================================
# 7. الأقسام الأخرى (رفع البيانات، السجل)
# =====================================
with tab_upload:
    up_tpl = st.file_uploader("ارفع قالب الوورد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f: f.write(up_tpl.getbuffer())
        st.success("تم التحديث")
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str})
            dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.to_sql('halls', conn, if_exists='replace', index=False)
            st.success("تم التحديث بنجاح"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_logs:
    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
    if st.button("🗑️ مسح السجل"): c.execute("DELETE FROM logs"); conn.commit(); st.rerun()
