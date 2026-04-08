import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
from datetime import datetime
import aspose.words as aw  # المكتبة المسؤولة عن التحويل الدقيق

# =====================================
# 1. نظام تسجيل الدخول
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
                        st.error("❌ اسم المستخدم غير معرف")
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
    .stDownloadButton button { background-color: #007bff !important; color: white !important; width: 100%; }
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
    c.execute("INSERT INTO logs (user, action, details, timestamp) VALUES (?, ?, ?, ?)", (st.session_state.username, action, details, now))
    conn.commit()

TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 3. وظائف معالجة الملفات (التحويل الحقيقي)
# =====================================

def word_to_pdf_bytes(word_io):
    """تحويل محتوى BytesIO الخاص بـ Word إلى PDF Bytes"""
    word_io.seek(0)
    doc = aw.Document(word_io)
    pdf_io = io.BytesIO()
    doc.save(pdf_io, aw.SaveFormat.PDF)
    return pdf_io.getvalue()

def process_doc(doc_obj, row, h_name, h_city):
    repls = {'<NAME>': str(row.get('name', '')), '<ID>': str(row.get('id', '')), 
             '<PHONE>': str(row.get('phone', '')), '<JOB>': str(row.get('role', '')), 
             '<HALL_NAME>': str(h_name), '<HALL_LOCATION>': str(h_city), 
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
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        add_log("حفظ تكليف", f"تم تكليف {row['name']}")
                        conn.commit(); st.success("تم الحفظ"); st.rerun()
                    
                    if row['hall']:
                        # استخراج ملف الوورد أولاً
                        word_file = generate_single_doc(row)
                        if word_file:
                            col_w, col_p = st.columns(2)
                            # زر الوورد
                            col_w.download_button("📥 Word", data=word_file, file_name=f"تكليف_{row['name']}.docx")
                            # زر الـ PDF (تحويل حقيقي لنفس ملف الوورد)
                            with st.spinner("جاري التحويل لـ PDF..."):
                                pdf_bytes = word_to_pdf_bytes(word_file)
                                col_p.download_button("📕 PDF", data=pdf_bytes, file_name=f"تكليف_{row['name']}.pdf")

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
            st.success("تم التحديث"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    # إحصائيات سريعة
    df_count = pd.read_sql("SELECT id, hall FROM teachers", conn)
    total_t = len(df_count)
    assigned_t = len(df_count[df_count['hall'] != ''])
    
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("الإجمالي", total_t)
    c_m2.metric("تم إنجازهم", assigned_t)
    c_m3.metric("المتبقي", total_t - assigned_t)
    
    st.divider()

    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_m_full = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            c_exc, c_wrd, c_pdf = st.columns(3)
            # زر إكسل القاعة
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_m_full[['id', 'name', 'role']].to_excel(writer, index=False)
            c_exc.download_button(f"📊 إكسل {h_choice}", data=output.getvalue(), file_name=f"كشف_{h_choice}.xlsx")
            
            # زر وورد القاعة (الكل في ملف واحد)
            bulk_word = generate_bulk_word(df_m_full, h_choice)
            if bulk_word:
                c_wrd.download_button(f"📄 وورد {h_choice}", data=bulk_word, file_name=f"تكليفات_{h_choice}.docx")
                # زر PDF القاعة (تحويل حقيقي للوورد المجمع)
                with st.spinner("جاري تحويل الكل لـ PDF..."):
                    bulk_pdf = word_to_pdf_bytes(bulk_word)
                    c_pdf.download_button(f"📕 PDF {h_choice}", data=bulk_pdf, file_name=f"تكليفات_{h_choice}.pdf")

with tab_logs:
    st.subheader("📜 سجل العمليات")
    df_l = pd.read_sql("SELECT user, action, details, timestamp FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
