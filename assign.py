import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import os
import zipfile

# =====================================
# قاعدة البيانات SQLite
# =====================================
conn = sqlite3.connect("data.db", check_same_thread=False)
c = conn.cursor()

# --- إنشاء جدول المعلمين إذا لم يكن موجود ---
c.execute('''
CREATE TABLE IF NOT EXISTS teachers (
    id TEXT PRIMARY KEY,
    name TEXT,
    school TEXT,
    city TEXT,
    phone TEXT,
    role TEXT,
    hall TEXT
)
''')
conn.commit()

# --- فحص العمود accept وإضافته إذا لم يكن موجود ---
c.execute("PRAGMA table_info(teachers)")
cols = [x[1] for x in c.fetchall()]
if "accept" not in cols:
    c.execute("ALTER TABLE teachers ADD COLUMN accept TEXT DEFAULT 'نعم'")
    conn.commit()

# --- جدول القاعات ---
c.execute('''
CREATE TABLE IF NOT EXISTS halls (
    number TEXT,
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

password = st.text_input("كلمة المرور", type="password", key="login_pass")
if st.button("دخول"):
    if password == "1234":  # غيّر كلمة المرور هنا
        st.session_state.logged = True
        st.success("تم الدخول ✅")
    else:
        st.error("كلمة المرور غلط")
if not st.session_state.logged:
    st.stop()

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

# =====================================
# قائمة المهام الجديدة
# =====================================
role_options = ["مراقب", "رئيس قاعة", "آذن", "مساعد رئيس قاعة"]

# =====================================
# إضافة قاعة
# =====================================
st.subheader("➕ إضافة قاعة")
with st.form("add_hall_form"):
    hall_name = st.text_input("اسم القاعة")
    hall_number = st.text_input("رقم القاعة")
    hall_city = st.text_input("البلد")
    if st.form_submit_button("إضافة القاعة"):
        if hall_name:
            try:
                c.execute("INSERT INTO halls VALUES (?,?,?)",
                          (hall_number, hall_name, hall_city))
                conn.commit()
                st.success("تمت إضافة القاعة ✅")
            except:
                st.error("❌ القاعة موجودة مسبقًا")
        else:
            st.warning("أدخل اسم القاعة")
    halls = get_halls()  # إعادة تحميل القاعات

# =====================================
# إضافة معلم جديد
# =====================================
with st.expander("➕ إضافة معلم جديد", expanded=True):
    with st.form("add_teacher_form"):
        name = st.text_input("الاسم", key="add_name")
        idd = st.text_input("الهوية", key="add_id")
        school = st.text_input("المدرسة", key="add_school")
        city = st.text_input("السكن", key="add_city")
        phone = st.text_input("الجوال", key="add_phone")
        role = st.selectbox("المهمة", role_options, key="add_role")
        accept = st.selectbox("هل يرغب بالعمل؟", ["نعم", "لا"], key="add_accept")
        if st.form_submit_button("💾 حفظ المعلم"):
            if not idd or not name:
                st.warning("⚠️ لازم تعبّي الاسم والهوية")
            else:
                try:
                    c.execute("INSERT INTO teachers VALUES (?,?,?,?,?,?,?,?)",
                              (idd, name, school, city, phone, role, "", accept))
                    conn.commit()
                    st.success("تم الحفظ ✅")
                    for k in ["add_name","add_id","add_school","add_city","add_phone"]:
                        st.session_state[k] = ""
                except:
                    st.error("❌ المعلم موجود مسبقًا")
    teachers = get_teachers()  # إعادة تحميل المعلمين

# =====================================
# البحث والتعيين
# =====================================
st.subheader("🔍 البحث والتعيين")
search = st.text_input("ابحث", key="search_box")
result = teachers.copy()
if search:
    result = teachers[
        teachers['name'].str.contains(search, case=False, na=False) |
        teachers['id'].astype(str).str.contains(search) |
        teachers['school'].str.contains(search, case=False, na=False)
    ]

if not result.empty:
    r = result.iloc[0]
    st.write(r)
    # إعداد قائمة القاعات للاختيار بشكل صحيح
    hall_options = [""] + [f"{row['number']} - {row['hall']}" for _, row in halls.iterrows()]
    selected_hall = st.selectbox("اختر القاعة", hall_options, key="assign_hall")
    role_assign = st.selectbox("المهمة", role_options, key="assign_role")

    if st.button("تعيين", key="assign_btn"):
        hall_name = selected_hall.split(" - ")[1] if selected_hall else ""
        c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?",
                  (hall_name, role_assign, r['id']))
        conn.commit()
        st.success("تم التعيين")
        teachers = get_teachers()

    if st.button("إلغاء", key="cancel_btn"):
        c.execute("UPDATE teachers SET hall='' WHERE id=?", (r['id'],))
        conn.commit()
        st.warning("تم الإلغاء")
        teachers = get_teachers()
