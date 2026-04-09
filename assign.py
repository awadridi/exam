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
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .main, .stApp { 
        direction: rtl; 
        text-align: right; 
        background-color: #0e1117; 
    }
    div[data-baseweb="select"], div[data-baseweb="input"], .stMultiSelect {
        direction: rtl !important;
        text-align: right !important;
    }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    .editor-info { color: #ffc107 !important; font-size: 0.9rem; font-weight: bold; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; color: #00ffcc !important; }
    .stat-card {
        flex: 1;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        min-width: 150px;
        border: 1px solid #333;
    }
    .stat-wants { border-top: 5px solid #28a745; background-color: #1a2e1f; }
    .stat-no-wants { border-top: 5px solid #dc3545; background-color: #2e1a1a; }
    
    /* الفئة الجديدة لإجبار النص على الانتقال لليسار (الجهة المقابلة) */
    .move-to-right {
        text-align: right !important;
        direction: ltr !important;
        display: block;
        width: 100%;
        color: white;
        font-family: sans-serif;
    }
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

tab_search, tab_auto, tab_upload, tab_manage, tab_logs = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📥 رفع البيانات", "📊 الإدارة والإحصائيات", "📜 سجل العمليات"])

with tab_search:
    # 1. إدارة الموظفين للجهة المقابلة
    st.markdown('<h2 class="move-to-right">إدارة الموظفين</h2>', unsafe_allow_html=True)
    df_h_data = get_cached_halls()
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم، الهوية، أو الجوال")
    if q:
        df_teachers = get_cached_teachers()
        results = df_teachers[df_teachers['name'].str.contains(q, na=False, case=False) | df_teachers['id'].astype(str).str.contains(q) | df_teachers['phone'].astype(str).str.contains(q)]
        for _, row in results.iterrows():
            display_phone = str(row['phone'])
            if display_phone.startswith('5') and len(display_phone) == 9:
                display_phone = '0' + display_phone

            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير مكلف'}"):
                st.markdown(f"""
                <div style="background-color: #1a1c23; padding: 15px; border-radius: 10px; border: 1px solid #444; border-right: 5px solid #00ffcc; margin-bottom: 15px; text-align: right;">
                    <table style="width:100%; color: white; border: none; direction: rtl;">
                        <tr>
                            <td style="padding: 5px;"><b>🆔 الهوية:</b> {row.get('id', '---')}</td>
                            <td style="padding: 5px;"><b>📱 الجوال:</b> {display_phone}</td>
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
                
                with st.popover("📝 تعديل البيانات الأساسية"):
                    with st.form(key=f"edit_base_{row['id']}"):
                        u_name = st.text_input("الاسم", value=row['name'])
                        u_phone = st.text_input("رقم الجوال", value=display_phone)
                        u_school = st.text_input("المدرسة", value=row['school'])
                        u_city = st.text_input("السكن", value=row['city'])
                        u_job = st.text_input("الوظيفة الأساسية", value=row['current_job'])
                        u_pref = st.selectbox("الرغبة", ["يرغب", "لا يرغب", "غير محدد"], index=0 if row['preference']=="يرغب" else (1 if row['preference']=="لا يرغب" else 2))
                        u_abil = st.selectbox("صلاحية المراقبة", ["يصلح", "لا يصلح", "لم تحدد"], index=0 if row['ability']=="يصلح" else (1 if row['ability']=="لا يصلح" else 2))
                        
                        if st.form_submit_button("💾 تحديث وحفظ"):
                            c.execute("""UPDATE teachers SET name=?, phone=?, school=?, city=?, current_job=?, preference=?, ability=?, updated_by=? 
                                         WHERE id=?""", (u_name, u_phone, u_school, u_city, u_job, u_pref, u_abil, st.session_state.username, row['id']))
                            conn.commit()
                            add_log("تعديل بيانات أساسية", f"تعديل بيانات {u_name}")
                            st.success("✅ تم التحديث بنجاح!")
                            time.sleep(1)
                            st.rerun()

                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    current_hall = row['hall'] if row['hall'] and str(row['hall']).lower() != 'nan' else ""
                    sel_h = st.selectbox("القاعة", [""] + list(hall_map.keys()), 
                                         index=(list(hall_map.keys()).index(current_hall)+1 if current_hall in hall_map else 0), 
                                         key=f"q_h_{row['id']}")
                    
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], 
                                         index=(["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"].index(row['role']) if row['role'] in ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"] else 0),
                                         key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
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
                        if st.button("❌ إلغاء التكليف", key=f"del_search_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='', updated_by=? WHERE id=?", 
                                      (st.session_state.username, row['id']))
                            conn.commit()
                            add_log("إلغاء تكليف", f"تم إلغاء تكليف {row['name']}")
                            st.rerun()
                        
                        if st.button("📥 إنشاء الكتاب", key=f"gen_s_{row['id']}"):
                            f_word = generate_single_doc(row)
                            if f_word: 
                                st.download_button("📥 تحميل الآن", data=f_word, 
                                               file_name=f"تكليف_{row['name']}.docx", 
                                               key=f"dl_s_{row['id']}")

with tab_auto:
    # 2. نظام التوزيع التلقائي للجهة المقابلة
    st.markdown('<h2 class="move-to-right">🤖 نظام التوزيع التلقائي</h2>', unsafe_allow_html=True)
    df_all = get_cached_teachers()
    df_available = df_all[(df_all['hall'] == '') | (df_all['hall'].isna())]
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        target_h = st.selectbox("اختر القاعة المستهدفة:", [""] + list(hall_map.keys()), key="auto_target_h")
        selected_cities = st.multiselect("السحب من مناطق سكن معينة:", sorted(df_available['city'].unique().tolist()))
        
    with col_a2:
        pool_stats = df_available
        if selected_cities:
            pool_stats = pool_stats[pool_stats['city'].isin(selected_cities)]
            
        df_auto_pool = pool_stats[pool_stats['ability'] == 'يصلح']
        num_to_assign = st.number_input("العدد المطلوب توزيعه:", min_value=1, max_value=max(1, len(df_auto_pool)), value=min(1, len(df_auto_pool)))

        if st.button("🚀 ابدأ التوزيع التلقائي الآن", use_container_width=True):
            if not target_h:
                st.error("الرجاء اختيار قاعة أولاً")
            elif len(df_auto_pool) < num_to_assign:
                st.error(f"العدد المتاح ({len(df_auto_pool)}) أقل من المطلوب.")
            else:
                selected_sample = df_auto_pool.sample(n=int(num_to_assign))
                for _, r in selected_sample.iterrows():
                    c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=?, updated_by='توزيع تلقائي' WHERE id=?", 
                              (target_h, hall_map[target_h], r['id']))
                conn.commit()
                add_log("توزيع تلقائي", f"توزيع {num_to_assign} مراقب على قاعة {target_h}")
                st.success(f"✅ تم توزيع {num_to_assign} بنجاح!")
                time.sleep(1)
                st.rerun()

    can_and_wants = len(pool_stats[(pool_stats['ability'] == 'يصلح') & (pool_stats['preference'] == 'يرغب')])
    can_not_wants = len(pool_stats[(pool_stats['ability'] == 'يصلح') & (pool_stats['preference'] == 'لا يرغب')])
    
    st.markdown(f"""
    <div style="display: flex; gap: 15px; margin-bottom: 20px; direction: rtl;">
        <div class="stat-card stat-wants">
            <span style="color: #bbb; font-size: 0.9rem;">يصلح ويرغب (في المناطق المختارة)</span><br>
            <strong style="font-size: 2rem; color: #28a745;">{can_and_wants}</strong>
        </div>
        <div class="stat-card stat-no-wants">
            <span style="color: #bbb; font-size: 0.9rem;">يصلح ولا يرغب (في المناطق المختارة)</span><br>
            <strong style="font-size: 2rem; color: #dc3545;">{can_not_wants}</strong>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        with st.expander("📍 تفاصيل (يصلح ويرغب) لكل منطقة"):
            df_w_sub = pool_stats[(pool_stats['ability'] == 'يصلح') & (pool_stats['preference'] == 'يرغب')]
            if not df_w_sub.empty:
                st.write(df_w_sub['city'].value_counts())
            else: st.info("لا توجد بيانات")
            
    with col_exp2:
        with st.expander("📍 تفاصيل (يصلح ولا يرغب) لكل منطقة"):
            df_nw_sub = pool_stats[(pool_stats['ability'] == 'يصلح') & (pool_stats['preference'] == 'لا يرغب')]
            if not df_nw_sub.empty:
                st.write(df_nw_sub['city'].value_counts())
            else: st.info("لا توجد بيانات")

with tab_upload:
    # 3. تحديث القالب والبيانات للجهة المقابلة
    st.markdown('<h2 class="move-to-right">تحديث القالب والبيانات</h2>', unsafe_allow_html=True)
    up_tpl = st.file_uploader("ارفع قالب الوورد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f:
            f.write(up_tpl.getbuffer())
        add_log("تحديث قالب", "تم رفع قالب وورد جديد")
        st.success("تم تحديث قالب الوورد بنجاح")
    
    st.divider()
    if st.button("🔄 تحديث من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL, dtype={'id': str, 'phone': str}) 
            dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            for col in ['phone', 'role', 'hall', 'hall_city', 'updated_by', 'preference', 'current_job', 'ability']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            
            dfh = pd.read_csv(HALLS_URL)
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            
            add_log("تحديث بيانات", "تحديث من جوجل شيت")
            st.success("تم التحديث بنجاح")
            st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

with tab_manage:
    df_all_teachers = get_cached_teachers()
    total_count = len(df_all_teachers)
    assigned_count = len(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0])
    remaining_count = total_count - assigned_count

    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.metric("إجمالي المعلمين", total_count)
    c_m2.metric("تم إنجازهم", assigned_count)
    c_m3.metric("المتبقي", remaining_count)
    
    st.divider()
    # 4. تصدير البيانات المعدلة للجهة المقابلة
    st.markdown('<h3 class="move-to-right">📦 تصدير البيانات المعدلة</h3>', unsafe_allow_html=True)
    df_export = df_all_teachers.copy()
    df_export.columns = ['رقم الهوية', 'الاسم كامل', 'رقم الجوال', 'المدرسة', 'السكن', 'المهمة المكلف بها', 'القاعة', 'مدينة القاعة', 'الموظف المعدل', 'الرغبة', 'الوظيفة', 'الصلاحية']
    
    output_all = io.BytesIO()
    with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='جميع المعلمين')
        workbook, worksheet = writer.book, writer.sheets['جميع المعلمين']
        h_fmt = workbook.add_format({'bold':True,'font_size':12,'border':1,'align':'center','bg_color':'#D7E4BC'})
        c_fmt = workbook.add_format({'font_size':11,'border':1,'align':'right'})
        worksheet.right_to_right()
        for col_num, col_name in enumerate(df_export.columns):
            worksheet.write(0, col_num, col_name, h_fmt)
            worksheet.set_column(col_num, col_num, 18, c_fmt)
    
    st.download_button("📥 تحميل كافة المعلمين (إكسل معدل)", data=output_all.getvalue(), file_name=f"كشف_المعلمين_المعدل_{datetime.now().strftime('%Y%m%d')}.xlsx")

    st.divider()
    assigned_halls = sorted(df_all_teachers[df_all_teachers['hall'].astype(str).str.len() > 0]['hall'].unique().tolist())
    
    if assigned_halls:
        h_choice = st.selectbox("اختر قاعة لعرض الكادر (يظهر فقط القاعات التي بها موظفون):", [""] + assigned_halls)
        if h_choice:
            df_hall_details = df_all_teachers[df_all_teachers['hall'] == h_choice]
            
            # 5. توزيع الكادر في قاعة للجهة المقابلة
            st.markdown(f'<h4 class="move-to-right">📊 توزيع الكادر في قاعة: {h_choice}</h4>', unsafe_allow_html=True)
            
            if not df_hall_details.empty:
                st.table(df_hall_details[['name', 'role', 'school', 'city', 'phone']])
            else:
                st.info("لا يوجد موظفون مكلفون في هذه القاعة حالياً.")

            col_btns1, col_btns2, col_spacer = st.columns([1, 1.5, 2.5])
            
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
                    with st.spinner("جاري إنشاء الملف المجمع..."):
                        bulk_f = generate_bulk_word(df_hall_details, h_choice)
                        if bulk_f: 
                            st.download_button("📥 تحميل الآن (وورد)", 
                                           data=bulk_f, 
                                           file_name=f"تكليفات_{h_choice}.docx",
                                           key=f"bulk_dl_now_{h_choice}")
            
            st.markdown("---")
            c_stat1, c_stat2, c_stat3, c_stat4 = st.columns(4)
            c_stat1.metric("رئيس قاعة", len(df_hall_details[df_hall_details['role'] == 'رئيس قاعة']))
            c_stat2.metric("مساعد رئيس", len(df_hall_details[df_hall_details['role'] == 'مساعد رئيس قاعة']))
            c_stat3.metric("مراقبين", len(df_hall_details[df_hall_details['role'] == 'مراقب']))
            c_stat4.metric("آذنة", len(df_hall_details[df_hall_details['role'] == 'آذن']))
    else:
        st.warning("لا يوجد أي قاعات مكلفة حالياً ليتم عرضها.")

with tab_logs:
    # 6. سجل العمليات للجهة المقابلة
    st.markdown('<h2 class="move-to-right">📜 سجل العمليات</h2>', unsafe_allow_html=True)
    df_l = pd.read_sql("SELECT user as 'الموظف', action as 'الإجراء', details as 'التفاصيل', timestamp as 'الوقت' FROM logs ORDER BY id DESC LIMIT 100", conn)
    st.dataframe(df_l, use_container_width=True)
    if st.button("🗑️ مسح السجل"):
        c.execute("DELETE FROM logs"); conn.commit(); st.cache_data.clear(); st.rerun()
