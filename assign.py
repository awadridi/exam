import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
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

# تحديث: إضافة preference, current_job, ability لقاعدة البيانات
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
                
                # --- مستطيل معلومات الموظف المحدث مع إضافة الصلاحية ---
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
                
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), st.session_state.username, row['id']))
                        add_log("حفظ تكليف", f"تم تكليف {row['name']} في {sel_h}")
                        conn.commit(); st.success("تم الحفظ"); st.rerun()
                    if row['hall']:
                        if st.button("❌ إلغاء التكليف", key=f"del_search_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", (st.session_state.username, row['id']))
                            add_log("إلغاء تكليف", f"تم إلغاء تكليف {row['name']}")
                            conn.commit(); st.rerun()
                        f_word = generate_single_doc(row)
                        if f_word: st.download_button("📥 تحميل الكتاب", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_s_{row['id']}")

with tab_upload:
    st.subheader("تحديث القالب والبيانات")
    
    up_tpl = st.file_uploader("ارفع قالب الوورد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f:
            f.write(up_tpl.getbuffer())
        add_log("تحديث قالب", "تم رفع قالب وورد جديد")
        st.success("تم تحديث قالب الوورد بنجاح")
    
    st.divider()
    
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL); dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            
            # تحديث: إضافة الأعمدة الجديدة لقائمة الفحص لضمان سحبها
            for col in ['phone', 'role', 'hall', 'hall_city', 'updated_by', 'preference', 'current_job', 'ability']: 
                if col not in dft.columns: dft[col] = ""
            
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.to_sql('halls', conn, if_exists='replace', index=False)
            add_log("تحديث بيانات", "تحديث من جوجل شيت")
            st.success("تم التحديث"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    try:
        df_all_source = pd.read_csv(TEACHERS_URL)
        total_count = len(df_all_source)
    except:
        total_count = len(pd.read_sql("SELECT id FROM teachers", conn))

    assigned_count = len(pd.read_sql("SELECT id FROM teachers WHERE hall != '' AND hall IS NOT NULL", conn))
    remaining_count = total_count - assigned_count

    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي المعلمين", total_count)
    c_m2.metric("تم إنجازهم", assigned_count)
    c_m3.metric("المتبقي", remaining_count)
    
    st.divider()

    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT role FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            count_chief = len(df_hall_details[df_hall_details['role'] == 'رئيس قاعة'])
            count_assistant = len(df_hall_details[df_hall_details['role'] == 'مساعد رئيس قاعة'])
            count_proctor = len(df_hall_details[df_hall_details['role'] == 'مراقب'])
            count_servant = len(df_hall_details[df_hall_details['role'] == 'آذن'])

            st.markdown(f"##### 📊 توزيع الكادر في قاعة: {h_choice}")
            c_stat1, c_stat2, c_stat3, c_stat4 = st.columns(4)
            c_stat1.metric("رئيس قاعة", count_chief)
            c_stat2.metric("مساعد رئيس", count_assistant)
            c_stat3.metric("مراقبين", count_proctor)
            c_stat4.metric("آذنة", count_servant)
            
            st.divider()
            
            df_m = pd.read_sql("SELECT id as 'رقم الهوية', name as 'الاسم', phone as 'رقم الجوال', school as 'المدرسة', role as 'المهمة', updated_by as 'الموظف' FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            c_exc, c_wrd = st.columns(2)
            with c_exc:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_m.to_excel(writer, index=False, sheet_name='العاملين')
                    workbook, worksheet = writer.book, writer.sheets['العاملين']
                    h_fmt = workbook.add_format({'bold':True,'font_size':14,'border':1,'align':'center','bg_color':'#D7E4BC'})
                    c_fmt = workbook.add_format({'font_size':14,'border':1,'align':'right'})
                    worksheet.right_to_left()
                    for col_num, col_name in enumerate(df_m.columns):
                        worksheet.write(0, col_num, col_name, h_fmt)
                        worksheet.set_column(col_num, col_num, 20, c_fmt)
                st.download_button(f"📊 إكسل {h_choice}", data=output.getvalue(), file_name=f"كشف_{h_choice}.xlsx")
            
            with c_wrd:
                df_m_full = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
                bulk_f = generate_bulk_word(df_m_full, h_choice)
                if bulk_f: st.download_button(f"📄 وورد {h_choice}", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")

            for _, m in df_m_full.iterrows():
                with st.expander(f"📝 {m['name']} | {m['role']}"):
                    col1, col2, col3 = st.columns([2, 2, 1])
                    with col1: new_h = st.selectbox("نقل", list(hall_map.keys()), index=list(hall_map.keys()).index(m['hall']), key=f"m_h_{m['id']}")
                    with col2: new_r = st.selectbox("مهمة", ["رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"m_r_{m['id']}")
                    with col3:
                        if st.button("تحديث", key=f"m_up_{m['id']}"):
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", (new_h, new_r, hall_map.get(new_h, ""), st.session_state.username, m['id']))
                            add_log("تعديل", f"نقل {m['name']} إلى {new_h}")
                            conn.commit(); st.rerun()
                        if st.button("حذف", key=f"m_del_{m['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", (st.session_state.username, m['id']))
                            add_log("حذف", f"حذف تكليف {m['name']}")
                            conn.commit(); st.rerun()

with tab_logs:
    st.subheader("📜 سجل العمليات")
    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
    if st.button("🗑️ مسح السجل"):
        c.execute("DELETE FROM logs"); conn.commit(); st.rerun()
