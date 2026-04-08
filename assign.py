import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
from datetime import datetime
from fpdf import FPDF

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
    .stDownloadButton button { background-color: #007bff !important; color: white !important; width: 100%; }
    [data-testid="stMetricValue"] { color: #00ffcc !important; }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect("data_system_v26.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT)''')
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

def generate_simple_pdf(row):
    """توليد ملف PDF بسيط لتجنب تعليق السيرفر ومشاكل المكتبات"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(200, 10, txt="Assignment Card - 2026", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", size=12)
    data = [
        ("Employee Name", str(row['name'])),
        ("ID Number", str(row['id'])),
        ("Duty / Role", str(row['role'])),
        ("Assigned Hall", str(row['hall'])),
        ("Hall Location", str(row['hall_city'])),
        ("Original School", str(row['school']))
    ]
    for label, val in data:
        pdf.cell(50, 10, txt=f"{label}:", border=1)
        pdf.cell(140, 10, txt=val, border=1, ln=True)
    return pdf.output()

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

tab_search, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "📥 رفع البيانات", "📊 الإحصائيات", "📜 السجل"])

with tab_search:
    st.subheader("إدارة الموظفين")
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        add_log("حفظ", f"تكليف {row['name']} في {sel_h}")
                        conn.commit(); st.success("تم الحفظ"); st.rerun()
                    
                    if row['hall']:
                        f_word = generate_single_doc(row)
                        if f_word:
                            col1, col2 = st.columns(2)
                            col1.download_button("📥 وورد (رسمي)", data=f_word, file_name=f"تكليف_{row['name']}.docx")
                            
                            # زر PDF بسيط
                            pdf_bytes = generate_simple_pdf(row)
                            col2.download_button("📕 PDF (سريع)", data=pdf_bytes, file_name=f"تكليف_{row['name']}.pdf", mime="application/pdf")

with tab_upload:
    st.subheader("تحديث القالب والبيانات")
    up_tpl = st.file_uploader("ارفع قالب الوورد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f:
            f.write(up_tpl.getbuffer())
        st.success("تم تحديث القالب")
    
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL); dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            for col in ['phone', 'role', 'hall', 'hall_city', 'updated_by']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.to_sql('halls', conn, if_exists='replace', index=False)
            st.success("تم التحديث بنجاح"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    df_all = pd.read_sql("SELECT id, hall FROM teachers", conn)
    total_t = len(df_all)
    assigned_t = len(df_all[df_all['hall'] != ''])
    
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي المعلمين", total_t)
    c_m2.metric("تم تكليفهم", assigned_t)
    c_m3.metric("المتبقي", total_t - assigned_t)
    
    st.divider()

    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    if not df_active.empty:
        h_choice = st.selectbox("عرض بيانات القاعة:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_m_full = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            c_exc, c_wrd = st.columns(2)
            # زر إكسل
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_m_full[['id', 'name', 'phone', 'school', 'role']].to_excel(writer, index=False)
            c_exc.download_button(f"📊 إكسل {h_choice}", data=output.getvalue(), file_name=f"كشف_{h_choice}.xlsx")
            
            # زر وورد مجمع
            bulk_f = generate_bulk_word(df_m_full, h_choice)
            if bulk_f:
                c_wrd.download_button(f"📄 وورد مجمع {h_choice}", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")

with tab_logs:
    st.subheader("📜 سجل العمليات")
    df_l = pd.read_sql("SELECT user, action, details, timestamp FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
