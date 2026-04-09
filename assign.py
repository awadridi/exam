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
                
                # --- تعديل البيانات الأساسية (تم تصحيح آلية الإغلاق) ---
                with st.popover("📝 تعديل البيانات الأساسية", key=f"pop_{row['id']}"):
                    with st.form(key=f"edit_form_{row['id']}"):
                        st.write(f"تعديل بيانات: {row['name']}")
                        u_name = st.text_input("الاسم", value=row['name'], key=f"in_n_{row['id']}")
                        u_phone = st.text_input("رقم الجوال", value=row['phone'], key=f"in_p_{row['id']}")
                        u_school = st.text_input("المدرسة", value=row['school'], key=f"in_s_{row['id']}")
                        u_city = st.text_input("السكن", value=row['city'], key=f"in_c_{row['id']}")
                        u_job = st.text_input("الوظيفة الأساسية", value=row['current_job'], key=f"in_j_{row['id']}")
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], 
                                             index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2),
                                             key=f"in_pr_{row['id']}")
                        u_abil = st.selectbox("صلاحية المراقبة", ["يصلح", "لا يصلح", "لم تحدد"], 
                                              index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2),
                                              key=f"in_ab_{row['id']}")
                        
                        btn_update = st.form_submit_button("💾 تحديث وحفظ")
                        
                        if btn_update:
                            # 1. تحديث قاعدة البيانات
                            c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, updated_by=? 
                                         WHERE id=?""", (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                            conn.commit()
                            # 2. إضافة لوج
                            add_log("تعديل بيانات أساسية", f"تعديل بيانات المعلم {u_name}")
                            # 3. إشعار سريع
                            st.toast(f"✅ تم تحديث {u_name}")
                            # 4. إعادة تحميل الصفحة (هذا يغلق الـ popover حتماً)
                            st.rerun()

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
            for col in ['phone', 'role', 'hall', 'hall_city', 'updated_by', 'preference', 'current_job', 'ability']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.to_sql('halls', conn, if_exists='replace', index=False)
            add_log("تحديث بيانات", "تحديث من جوجل شيت")
            st.success("تم التحديث")
            st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    df_all_teachers = pd.read_sql("SELECT * FROM teachers", conn)
    total_count = len(df_all_teachers)
    assigned_count = len(df_all_teachers[df_all_teachers['hall'].str.len() > 0])
    remaining_count = total_count - assigned_count

    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي المعلمين", total_count)
    c_m2.metric("تم إنجازهم", assigned_count)
    c_m3.metric("المتبقي", remaining_count)
    
    st.divider()
    st.subheader("📦 تصدير البيانات المعدلة")
    df_export = df_all_teachers.copy()
    df_export.columns = ['رقم الهوية', 'الاسم كامل', 'رقم الجوال', 'المدرسة', 'السكن', 'المهمة المكلف بها', 'القاعة', 'مدينة القاعة', 'الموظف المعدل', 'الرغبة', 'الوظيفة', 'الصلاحية']
    
    output_all = io.BytesIO()
    with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='جميع المعلمين')
        workbook, worksheet = writer.book, writer.sheets['جميع المعلمين']
        h_fmt = workbook.add_format({'bold':True,'font_size':12,'border':1,'align':'center','bg_color':'#D7E4BC'})
        c_fmt = workbook.add_format({'font_size':11,'border':1,'align':'right'})
        worksheet.right_to_left()
        for col_num, col_name in enumerate(df_export.columns):
            worksheet.write(0, col_num, col_name, h_fmt)
            worksheet.set_column(col_num, col_num, 18, c_fmt)
    
    st.download_button("📥 تحميل كافة المعلمين (إكسل معدل)", data=output_all.getvalue(), file_name=f"كشف_المعلمين_المعدل_{datetime.now().strftime('%Y%m%d')}.xlsx")

    st.divider()
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_hall_details = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            st.markdown(f"##### 📊 توزيع الكادر في قاعة: {h_choice}")
            c_stat1, c_stat2, c_stat3, c_stat4 = st.columns(4)
            c_stat1.metric("رئيس قاعة", len(df_hall_details[df_hall_details['role'] == 'رئيس قاعة']))
            c_stat2.metric("مساعد رئيس", len(df_hall_details[df_hall_details['role'] == 'مساعد رئيس قاعة']))
            c_stat3.metric("مراقبين", len(df_hall_details[df_hall_details['role'] == 'مراقب']))
            c_stat4.metric("آذنة", len(df_hall_details[df_hall_details['role'] == 'آذن']))
            
            c_exc, c_wrd = st.columns(2)
            with c_exc:
                st.write("استخدم زر التصدير الشامل أعلاه.")
            
            with c_wrd:
                bulk_f = generate_bulk_word(df_hall_details, h_choice)
                if bulk_f: st.download_button(f"📄 وورد قاعة {h_choice}", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")

with tab_logs:
    st.subheader("📜 سجل العمليات")
    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
    if st.button("🗑️ مسح السجل"):
        c.execute("DELETE FROM logs"); conn.commit(); st.rerun()
