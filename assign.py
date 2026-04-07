import streamlit as st
import pandas as pd
import sqlite3
# from docx import Document  # مؤجل لحين استخدامه
# import os
# import zipfile

# =====================================
# إعدادات الصفحة
# =====================================
st.set_page_config(page_title="نظام إدارة التكليف", page_icon="🎓")

# =====================================
# قاعدة البيانات SQLite
# =====================================
# يفضل استخدام دالة مع caching للاتصال لتجنب مشاكل الـ Threads
@st.cache_resource
def get_db_connection():
    return sqlite3.connect("data.db", check_same_thread=False)

conn = get_db_connection()
c = conn.cursor()

# إنشاء الجداول بشكل كامل من البداية
c.execute('''
CREATE TABLE IF NOT EXISTS teachers (
    id TEXT PRIMARY KEY,
    name TEXT,
    school TEXT,
    city TEXT,
    phone TEXT,
    role TEXT,
    hall TEXT,
    accept TEXT DEFAULT 'نعم'
)
''')

c.execute('''
CREATE TABLE IF NOT EXISTS halls (
    number TEXT PRIMARY KEY,
    hall TEXT,
    city TEXT
)
''')
conn.commit()

# =====================================
# تسجيل الدخول
# =====================================
st.title("🎓 نظام إدارة التكليف")

if "logged" not in st.session_state:
    st.session_state.logged = False

# يفضل جلب الباسورد من st.secrets في البيئة الحقيقية
# password_secret = st.secrets.get("password", "1234") 
password_secret = "1234" # غيرها لاحقاً لـ secrets

if not st.session_state.logged:
    password = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if password == password_secret:  
            st.session_state.logged = True
            st.rerun() # تحديث الصفحة فورا بعد تسجيل الدخول
        else:
            st.error("كلمة المرور غير صحيحة")
    st.stop() # إيقاف تنفيذ باقي الكود إذا لم يسجل دخول

if st.session_state.logged:
    if st.button("تسجيل خروج"):
        st.session_state.logged = False
        st.rerun()

# =====================================
# دوال تحميل البيانات
# =====================================
def get_teachers():
    return pd.read_sql("SELECT * FROM teachers WHERE accept='نعم'", conn)

def get_halls():
    df = pd.read_sql("SELECT * FROM halls", conn)
    df['number'] = df['number'].fillna("").astype(str)
    df['hall'] = df['hall'].fillna("").astype(str)
    return df

teachers = get_teachers()
halls = get_halls()

role_options = ["مراقب", "رئيس قاعة", "آذن", "مساعد رئيس قاعة"]

# =====================================
# واجهة التقسيم (Tabs) لتحسين شكل الموقع
# =====================================
tab1, tab2, tab3 = st.tabs(["🔍 البحث والتعيين", "👨‍🏫 إضافة معلم", "🏢 إضافة قاعة"])

# --- التبويب الأول: البحث والتعيين ---
with tab1:
    st.subheader("البحث والتعيين")
    search = st.text_input("ابحث بالاسم، الهوية، أو المدرسة")
    
    if search:
        result = teachers[
            teachers['name'].str.contains(search, case=False, na=False) |
            teachers['id'].astype(str).str.contains(search) |
            teachers['school'].str.contains(search, case=False, na=False)
        ]

        if not result.empty:
            # دمج الاسم والهوية عشان نميز لو في أسماء متشابهة
            teacher_options = result['name'] + " - " + result['id']
            selected_teacher_info = st.selectbox("اختر المعلم من نتائج البحث", teacher_options)
            
            # استخراج الهوية من الخيار المحدد لجلب بياناته
            selected_id = selected_teacher_info.split(" - ")[1]
            r = result[result['id'] == selected_id].iloc[0]
            
            # عرض بيانات المعلم المحدد
            st.info(f"المعلم الحالي: {r['name']} | القاعة الحالية: {r['hall']} | المهمة: {r['role']}")
            
            hall_options = [""] + [f"{row['number']} - {row['hall']}" for _, row in halls.iterrows()]
            
            col1, col2 = st.columns(2)
            with col1:
                selected_hall = st.selectbox("اختر القاعة", hall_options)
            with col2:
                role_assign = st.selectbox("المهمة", role_options, index=role_options.index(r['role']) if r['role'] in role_options else 0)

            col3, col4 = st.columns(2)
            with col3:
                if st.button("✅ تعيين", use_container_width=True):
                    hall_name = selected_hall.split(" - ")[1] if selected_hall else ""
                    c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (hall_name, role_assign, selected_id))
                    conn.commit()
                    st.success("تم التعيين بنجاح!")
                    st.rerun() # تحديث فوري
            with col4:
                if st.button("❌ إلغاء التعيين", use_container_width=True):
                    c.execute("UPDATE teachers SET hall='', role='' WHERE id=?", (selected_id,))
                    conn.commit()
                    st.warning("تم إزالة القاعة والمهمة.")
                    st.rerun()
        else:
            st.warning("لا يوجد نتائج مطابقة للبحث.")

# --- التبويب الثاني: إضافة معلم ---
with tab2:
    st.subheader("إضافة معلم جديد")
    with st.form("add_teacher_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("الاسم")
            idd = st.text_input("رقم الهوية")
            school = st.text_input("المدرسة")
        with col2:
            city = st.text_input("مكان السكن")
            phone = st.text_input("رقم الجوال")
            role = st.selectbox("المهمة", role_options)
        
        accept = st.selectbox("هل يرغب بالعمل؟", ["نعم", "لا"])
        
        if st.form_submit_button("💾 حفظ المعلم"):
            if not idd or not name:
                st.warning("⚠️ يرجى تعبئة الاسم ورقم الهوية على الأقل.")
            else:
                try:
                    c.execute("INSERT INTO teachers (id, name, school, city, phone, role, hall, accept) VALUES (?,?,?,?,?,?, '',?)",
                              (idd, name, school, city, phone, role, accept))
                    conn.commit()
                    st.success("تم الحفظ بنجاح! ✅")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("❌ رقم الهوية موجود مسبقاً في النظام!")

# --- التبويب الثالث: إضافة قاعة ---
with tab3:
    st.subheader("إضافة قاعة")
    with st.form("add_hall_form", clear_on_submit=True):
        hall_name = st.text_input("اسم القاعة")
        hall_number = st.text_input("رقم القاعة")
        hall_city = st.text_input("البلد")
        
        if st.form_submit_button("إضافة القاعة"):
            if hall_name and hall_number:
                try:
                    c.execute("INSERT INTO halls (number, hall, city) VALUES (?,?,?)",
                              (hall_number, hall_name, hall_city))
                    conn.commit()
                    st.success("تمت إضافة القاعة ✅")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("❌ رقم القاعة موجود مسبقاً!")
            else:
                st.warning("يرجى إدخال اسم ورقم القاعة.")
