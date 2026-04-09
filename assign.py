import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import time  # استيراد مكتبة الوقت للانتظار
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
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

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
    repls = {'<NAME>': str(row.get('name', '')), '<ID>': str(row.get('id', '')), '<PHONE>': str(row.get('phone', '')), 
             '<JOB>': str(row.get('role', '')), '<HALL_NAME>': str(h_name), '<HALL_LOCATION>': str(h_city), 
             '<WORKPLACE>': str(row.get('school', '')), '<CITY>': str(row.get('city', ''))}
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text: run.text = run.text.replace(k, v); run.bold = True
    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in repls.items():
                        if k in p.text:
                            for run in p.runs:
                                if k in run.text: run.text = run.text.replace(k, v); run.bold = True
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

tab_search, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

with tab_search:
    st.subheader("إدارة الموظفين")
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                
                st.markdown(f"""
                <div style="background-color: #1a1c23; padding: 15px; border-radius: 10px; border: 1px solid #444; border-right: 5px solid #00ffcc; margin-bottom: 15px; text-align: right;">
                    <table style="width:100%; color: white; border: none; direction: rtl;">
                        <tr>
                            <td style="padding: 5px;"><b>🆔 الهوية:</b> {row.get('id', '---')}</td>
                            <td style="padding: 5px;"><b>📱 الجوال:</b> {row.get('phone', '---')}</td>
                        </tr>
                        <tr>
                            <td style="padding: 5px;"><b>🏡 السكن:</b> {row.get('city', '---')}</td>
                            <td style="padding: 5px;"><b>🏫 المدرسة:</b> {row.get('school', '---')}</td>
                        </tr>
                        <tr>
                            <td style="padding: 5px;"><b>📝 الرغبة:</b> {row.get('preference', 'غير محدد')}</td>
                            <td style="padding: 5px;"><b>💼 الوظيفة:</b> {row.get('current_job', 'غير محدد')}</td>
                        </tr>
                        <tr>
                            <td colspan="2" style="padding: 5px; border-top: 1px solid #444; color: #ffc107;">
                                <b>⚠️ صلاحية المراقبة:</b> {row.get('ability', 'لم تحدد')}
                            </td>
                        </tr>
                    </table>
                </div>
                """, unsafe_allow_html=True)

                st.markdown(f"<span class='editor-info'>آخر تعديل: {row['updated_by'] or 'لا يوجد'}</span>", unsafe_allow_html=True)
                
                # --- تعديل البيانات الأساسية ---
                update_count_key = f"update_tick_{row['id']}"
                if update_count_key not in st.session_state:
                    st.session_state[update_count_key] = 0
                
                with st.popover("📝 تعديل البيانات الأساسية", key=f"pop_{row['id']}_{st.session_state[update_count_key]}"):
                    with st.form(key=f"edit_base_{row['id']}"):
                        st.write(f"تعديل بيانات: {row['name']}")
                        u_name = st.text_input("الاسم", value=row['name'])
                        u_phone = st.text_input("رقم الجوال", value=row['phone'])
                        u_school = st.text_input("المدرسة", value=row['school'])
                        u_city = st.text_input("السكن", value=row['city'])
                        u_job = st.text_input("الوظيفة الأساسية", value=row['current_job'])
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], 
                                             index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2))
                        u_abil = st.selectbox("صلاحية المراقبة", ["يصلح", "لا يصلح", "لم تحدد"], 
                                              index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2))
                        
                        if st.form_submit_button("💾 تحديث وحفظ"):
                            c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, updated_by=? 
                                         WHERE id=?""", (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                            conn.commit()
                            add_log("تعديل بيانات أساسية", f"تعديل بيانات المعلم {u_name}")
                            st.session_state[update_count_key] += 1
                            
                            st.success("✅ تم التحديث بنجاح!")
                            time.sleep(2) # الانتظار لثانيتين
                            st.rerun()

                # --- جزء تحديد القاعة والمهمة (تم التأكد من مكانه ليبقى ظاهراً) ---
                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        add_log("حفظ تكليف", f"تم تكليف {row['name']} في {sel_h}")
                        conn.commit()
                        st.rerun()
                    if row['hall']:
                        if st.button("❌ إلغاء التكليف", key=f"del_search_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", (st.session_state.username, row['id']))
                            add_log("إلغاء تكليف", f"تم إلغاء تكليف {row['name']}")
                            conn.commit()
                            st.rerun()
                        f_word = generate_single_doc(row)
                        if f_word: st.download_button("📥 تحميل الكتاب", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_s_{row['id']}")

# الأجزاء الأخرى من الكود (Upload, Manage, Logs) تتبع هنا...
# (بقيت كما هي في الكود الأصلي لتعمل بشكل سليم)
