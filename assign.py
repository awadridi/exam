import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
from copy import deepcopy

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
                        # يتم جلب كلمة المرور من Secrets
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
st.set_page_config(page_title="إدارة الموظفين والتعيين اليدوي", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0b0e14; }
    [data-testid="stSidebar"] { direction: rtl; text-align: right; }
    
    /* تصميم البطاقة مثل الصور المرفقة */
    .teacher-card {
        background-color: #161b22;
        padding: 15px;
        border-radius: 4px;
        border-right: 4px solid #00f2ea; /* الشريط الجانبي الملون */
        margin-bottom: 10px;
        color: white;
        position: relative;
    }
    .card-header {
        display: flex;
        justify-content: space-between;
        border-bottom: 1px solid #30363d;
        padding-bottom: 5px;
        margin-bottom: 10px;
    }
    .info-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
    }
    .label-icon { color: #8b949e; margin-left: 5px; }
    .value-text { color: #c9d1d9; font-weight: bold; }
    .status-badge {
        color: #e3b341;
        font-weight: bold;
        font-size: 0.85rem;
    }
    </style>
    """, unsafe_allow_html=True)

conn = sqlite3.connect("data_system_v26.db", check_same_thread=False)
c = conn.cursor()

# تحديث الجدول ليشمل الحقول الظاهرة في الصور
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, phone TEXT, school TEXT, city TEXT, 
             role TEXT, hall TEXT, hall_city TEXT, updated_by TEXT,
             preference TEXT, current_job TEXT, ability TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
conn.commit()

# =====================================
# 3. معالجة المستندات (Word)
# =====================================
def process_doc(doc_obj, row, h_name, h_city):
    repls = {
        'ZNAME': str(row.get('name', '')), 
        'ZID': str(row.get('id', '')), 
        'ZHALL': str(h_name or ''), 
        'ZLOC': str(h_city or ''), 
        'ZWORK': str(row.get('school', '')), 
        'ZCITY': str(row.get('city', ''))
    }
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)
    return doc_obj

def generate_single_doc(row):
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# =====================================
# 4. الواجهة الرئيسية
# =====================================
tab_search, tab_auto, tab_manage, tab_upload = st.tabs([
    "🔍 إدارة الموظفين والتعيين اليدوي", "🤖 التوزيع التلقائي", "📊 الإحصائيات", "📥 رفع البيانات"
])

with tab_search:
    st.markdown("### 🔍 ابحث عن الاسم، الهوية، أو الجوال")
    q = st.text_input("", placeholder="اكتب هنا للبحث...", label_visibility="collapsed")
    
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}

    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | 
                              df_teachers['id'].astype(str).str.contains(q) | 
                              df_teachers['phone'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            # تنسيق حالة المراقبة كما في الصورة
            ability_status = "✅ يصلح" if row['ability'] != "لا يصلح" else "❌ لا يصلح"
            
            st.markdown(f"""
            <div class="teacher-card">
                <div class="card-header">
                    <span>👤 <b>{row['name']}</b> | القاعة: <span style="color:#58a6ff">{row['hall'] or 'غير محدد'}</span></span>
                </div>
                <div class="info-grid">
                    <div><span class="label-icon">🆔 الهوية:</span><span class="value-text">{row['id']}</span></div>
                    <div><span class="label-icon">📱 الجوال:</span><span class="value-text">{row['phone']}</span></div>
                    <div><span class="label-icon">🏫 المدرسة:</span><span class="value-text">{row['school']}</span></div>
                    <div><span class="label-icon">🏠 السكن:</span><span class="value-text">{row['city']}</span></div>
                    <div><span class="label-icon">💼 الوظيفة:</span><span class="value-text">{row['current_job'] or 'معلم'}</span></div>
                    <div><span class="label-icon">📝 الرغبة:</span><span class="value-text">{row['preference'] or 'غير مسجل'}</span></div>
                </div>
                <div style="margin-top:10px;">
                    <span class="status-badge">⚠️ صلاحية المراقبة: {ability_status}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # أدوات التحكم (التعيين)
            c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
            with c1:
                new_h = st.selectbox("القاعة:", [""] + list(hall_map.keys()), key=f"h_{row['id']}", index=0)
            with c2:
                new_r = st.selectbox("المهمة:", ["مراقب", "رئيس قاعة", "مساعد", "آذن"], key=f"r_{row['id']}")
            with c3:
                if st.button("💾 حفظ", key=f"sv_{row['id']}", use_container_width=True):
                    c.execute("UPDATE teachers SET hall=?, role=?, hall_city=?, updated_by=? WHERE id=?", 
                             (new_h, new_r, hall_map.get(new_h, ""), st.session_state.username, row['id']))
                    conn.commit(); st.success("تم التحديث"); st.rerun()
            with c4:
                if row['hall']:
                    doc_file = generate_single_doc(row)
                    if doc_file:
                        st.download_button("📄 عقد", data=doc_file, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

with tab_upload:
    st.info("ارفع ملف Excel يحتوي على الأعمدة: id, name, phone, school, city, preference, current_job")
    up_file = st.file_uploader("اختر ملف البيانات", type="xlsx")
    if up_file:
        df_new = pd.read_excel(up_file)
        # التأكد من أن الهوية نصية
        df_new['id'] = df_new['id'].astype(str)
        df_new.to_sql("teachers", conn, if_exists="replace", index=False)
        st.success("تم رفع بيانات المعلمين بنجاح!")
    
    st.divider()
    up_docx = st.file_uploader("ارفع قالب الورد (template.docx)", type="docx")
    if up_docx:
        with open("template.docx", "wb") as f: f.write(up_docx.getbuffer())
        st.success("تم تحديث قالب التكليف")

# إغلاق الاتصال عند الانتهاء
conn.close()
