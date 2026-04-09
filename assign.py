import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Mm
import io
import os
from datetime import datetime
from copy import deepcopy

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
    .main-info-box {
        background-color: #1e2129;
        padding: 15px;
        border-radius: 8px;
        border-right: 5px solid #00ffcc;
        margin-bottom: 15px;
    }
    .data-label { color: #888; font-size: 0.9rem; }
    .data-value { color: #fff; font-weight: bold; margin-left: 15px; }
    button[key^="save_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
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
conn.commit()

# =====================================
# 3. وظائف معالجة الملفات
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    phone_val = str(row.get('phone', ''))
    if phone_val.startswith('5') and len(phone_val) == 9: phone_val = '0' + phone_val
    
    repls = {
        'ZNAME': str(row.get('name', '')), 
        'ZID': str(row.get('id', '')), 
        'ZJOB': str(row.get('role', '') or 'مراقب'), 
        'ZHALL': str(h_name or ''), 
        'ZLOC': str(h_city or ''), 
        'ZWORK': str(row.get('school', '')), 
        'ZCITY': str(row.get('city', '')),
        'ZPHONE': phone_val
    }
    
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)
    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in repls.items():
                        if k in p.text:
                            for run in p.runs:
                                if k in run.text:
                                    run.text = run.text.replace(k, v)
    return doc_obj

def generate_bulk_word(df, h_name):
    if not os.path.exists("template.docx") or df.empty: return None
    final_doc = Document("template.docx")
    final_doc._body.clear_content()
    
    for idx, row in df.iterrows():
        temp_doc = Document("template.docx")
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        for element in temp_doc.element.body:
            if element.tag.endswith('sectPr'): continue
            final_doc.element.body.append(deepcopy(element))
        if idx < len(df) - 1:
            final_doc.add_page_break()
            
    out = io.BytesIO()
    final_doc.save(out); out.seek(0)
    return out

def generate_single_doc(row):
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# =====================================
# 4. الواجهة الرئيسية والتبويبات
# =====================================
tab_search, tab_auto, tab_manage, tab_upload = st.tabs(["🔍 البحث والتعيين", "🤖 التوزيع التلقائي", "📊 الإدارة والإحصائيات", "📥 الرفع والمزامنة"])

with tab_search:
    st.markdown("<h3 class='right-align'>🔍 البحث عن موظف وتعيينه يدوياً</h3>", unsafe_allow_html=True)
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث بالاسم، الهوية، أو الجوال:")
    
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | 
                              df_teachers['id'].astype(str).str.contains(q) | 
                              df_teachers['phone'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            with st.container():
                st.markdown(f"""
                <div class='main-info-box'>
                    <h4 style='margin:0; color:#00ffcc;'>👤 {row['name']}</h4>
                    <span class='data-label'>🆔 الهوية:</span> <span class='data-value'>{row['id']}</span>
                    <span class='data-label'>📱 الجوال:</span> <span class='data-value'>{row['phone']}</span>
                    <span class='data-label'>🏫 المدرسة:</span> <span class='data-value'>{row['school']}</span><br>
                    <span class='data-label'>📍 السكن:</span> <span class='data-value'>{row['city']}</span>
                    <span class='data-label'>💼 الوظيفة:</span> <span class='data-value'>{row['current_job']}</span>
                    <span class='data-label'>⭐ الرغبة:</span> <span class='data-value'>{row['preference']}</span>
                    <hr style='border: 0.5px solid #333;'>
                    <span class='data-label'>🚩 القاعة الحالية:</span> <span class='data-value' style='color:#ffcc00;'>{row['hall'] or 'غير مكلف'}</span>
                    <span class='data-label'>🛠️ المهمة:</span> <span class='data-value'>{row['role'] or '---'}</span>
                </div>
                """, unsafe_allow_html=True)
                
                c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
                with c1:
                    new_h = st.selectbox("اختر القاعة:", [""] + list(hall_map.keys()), key=f"h_{row['id']}")
                with c2:
                    new_r = st.selectbox("اختر المهمة:", ["", "رئيس قاعة", "مساعد رئيس", "مراقب", "آذن"], key=f"r_{row['id']}")
                with c3:
                    if st.button("💾 حفظ", key=f"save_{row['id']}", use_container_width=True):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                                  (new_h, new_r, hall_map.get(new_h, ""), st.session_state.username, row['id']))
                        conn.commit(); st.success("تم الحفظ"); st.rerun()
                with c4:
                    if row['hall']:
                        if st.button("🗑️ إلغاء", key=f"del_{row['id']}", use_container_width=True):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit(); st.rerun()
                
                if row['hall']:
                    doc_file = generate_single_doc(row)
                    if doc_file:
                        st.download_button("📥 تحميل التكليف", data=doc_file, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}", use_container_width=True)
                st.divider()

with tab_auto:
    st.markdown("<h3 class='right-align'>🤖 التوزيع التلقائي الذكي</h3>", unsafe_allow_html=True)
    df_avail = pd.read_sql("SELECT * FROM teachers WHERE (hall = '' OR hall IS NULL)", conn)
    if not df_avail.empty:
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            target_h = st.selectbox("القاعة المستهدفة:", list(hall_map.keys()), key="auto_h")
            selected_cities = st.multiselect("السحب من مناطق:", df_avail['city'].unique())
        with col_a2:
            count = st.number_input("العدد المطلوب:", min_value=1, max_value=len(df_avail), value=1)
            if st.button("🚀 ابدأ التوزيع التلقائي", use_container_width=True):
                pool = df_avail[df_avail['city'].isin(selected_cities)].sample(frac=1).head(count)
                for _, r in pool.iterrows():
                    c.execute("UPDATE teachers SET hall=?, role='مراقب', hall_city=? WHERE id=?", (target_h, hall_map[target_h], r['id']))
                conn.commit(); st.success(f"تم توزيع {len(pool)} مراقب"); st.rerun()
    else:
        st.info("لا يوجد موظفون متاحون للتوزيع.")

with tab_manage:
    st.markdown("<h3 class='right-align'>📊 إحصائيات القاعات والكشوفات</h3>", unsafe_allow_html=True)
    df_all = pd.read_sql("SELECT * FROM teachers", conn)
    halls_list = df_all[df_all['hall'] != '']['hall'].unique()
    if len(halls_list) > 0:
        sel_h_view = st.selectbox("عرض كشف قاعة:", sorted(halls_list))
        df_view = df_all[df_all['hall'] == sel_h_view]
        st.dataframe(df_view[['name', 'id', 'role', 'school', 'phone']], use_container_width=True)
        bulk_doc = generate_bulk_word(df_view, sel_h_view)
        if bulk_doc:
            st.download_button(f"📥 تحميل كافة تكليفات قاعة {sel_h_view}", data=bulk_doc, file_name=f"كشوفات_{sel_h_view}.docx", use_container_width=True)

with tab_upload:
    st.markdown("<h3 class='right-align'>📥 رفع البيانات والقالب</h3>", unsafe_allow_html=True)
    up_docx = st.file_uploader("ارفع قالب الورد (template.docx):", type="docx")
    if up_docx:
        with open("template.docx", "wb") as f: f.write(up_docx.getbuffer())
        st.success("تم تحديث القالب.")
    
    if st.button("🔄 مزامنة من Google Sheets (الرابط المبرمج)"):
        st.warning("هذه العملية ستقوم بتحديث قاعدة البيانات بالكامل.")
        # هنا يمكن إضافة كود pd.read_csv المباشر من الروابط التي كانت في الكود السابق
