import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import os
import zipfile

# =====================================
# قاعدة البيانات
# =====================================
conn = sqlite3.connect("data.db", check_same_thread=False)
c = conn.cursor()

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

c.execute('''
CREATE TABLE IF NOT EXISTS halls (
    hall TEXT,
    number TEXT,
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

if st.button("دخول", key="login_btn"):
    if password == "1234":
        st.session_state.logged = True
        st.success("تم الدخول ✅")
    else:
        st.error("كلمة المرور غلط")

if not st.session_state.logged:
    st.stop()

# =====================================
# تحميل البيانات
# =====================================
def get_teachers():
    return pd.read_sql("SELECT * FROM teachers", conn)

def get_halls():
    return pd.read_sql("SELECT * FROM halls", conn)

teachers = get_teachers()
halls = get_halls()

# =====================================
# ➕ إضافة قاعة
# =====================================
st.subheader("➕ إضافة قاعة")
with st.form("add_hall_form"):
    hall_name = st.text_input("اسم القاعة", key="hall_name")
    hall_number = st.text_input("رقم القاعة", key="hall_number")
    hall_city = st.text_input("البلد", key="hall_city")

    if st.form_submit_button("إضافة القاعة"):
        if hall_name:
            c.execute("INSERT INTO halls VALUES (?,?,?)",
                      (hall_name, hall_number, hall_city))
            conn.commit()
            st.success("تمت إضافة القاعة ✅")
            st.experimental_rerun()
        else:
            st.warning("أدخل اسم القاعة")

# =====================================
# ➕ إضافة معلم
# =====================================
with st.expander("➕ إضافة معلم جديد", expanded=True):
    with st.form("add_teacher_form"):
        name = st.text_input("الاسم", key="add_name")
        idd = st.text_input("الهوية", key="add_id")
        school = st.text_input("المدرسة", key="add_school")
        city = st.text_input("السكن", key="add_city")
        phone = st.text_input("الجوال", key="add_phone")
        role = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"], key="add_role")

        submitted = st.form_submit_button("💾 حفظ المعلم")

        if submitted:
            if not idd or not name:
                st.warning("⚠️ لازم تعبّي الاسم والهوية")
            elif not idd.isdigit():
                st.error("رقم الهوية يجب أن يكون أرقام فقط")
            else:
                try:
                    c.execute("INSERT INTO teachers VALUES (?,?,?,?,?,?,?)",
                              (idd, name, school, city, phone, role, ""))
                    conn.commit()
                    st.success("تم الحفظ ✅")
                    # تفريغ الحقول
                    for k in ["add_name","add_id","add_school","add_city","add_phone"]:
                        st.session_state[k] = ""
                    st.experimental_rerun()
                except:
                    st.error("❌ المعلم موجود مسبقًا")

teachers = get_teachers()
halls = get_halls()

# =====================================
# فلترة حسب المدرسة
# =====================================
school_filter = st.selectbox(
    "فلترة حسب المدرسة",
    ["الكل"] + sorted(teachers['school'].dropna().unique().tolist()),
    key="school_filter"
)

filtered_teachers = teachers.copy()
if school_filter != "الكل":
    filtered_teachers = filtered_teachers[filtered_teachers['school'] == school_filter]

# =====================================
# ✏️ تعديل / حذف
# =====================================
st.subheader("✏️ تعديل أو حذف معلم")
if not filtered_teachers.empty:
    selected = st.selectbox("اختر معلم", filtered_teachers['name'], key="edit_select")
    row = filtered_teachers[filtered_teachers['name']==selected].iloc[0]

    new_name = st.text_input("الاسم الجديد", row['name'], key="edit_name")
    new_phone = st.text_input("الجوال الجديد", row['phone'], key="edit_phone")

    if st.button("تحديث", key="update_btn"):
        c.execute("UPDATE teachers SET name=?, phone=? WHERE id=?",
                  (new_name, new_phone, row['id']))
        conn.commit()
        st.success("تم التعديل")
        st.experimental_rerun()

    if st.button("حذف", key="delete_btn"):
        c.execute("DELETE FROM teachers WHERE id=?", (row['id'],))
        conn.commit()
        st.warning("تم الحذف")
        st.experimental_rerun()

# =====================================
# 🔍 البحث والتعيين
# =====================================
st.subheader("🔍 البحث والتعيين")
search = st.text_input("ابحث", key="search_box")
result = filtered_teachers.copy()
if search:
    result = filtered_teachers[
        filtered_teachers['name'].str.contains(search, case=False, na=False) |
        filtered_teachers['id'].astype(str).str.contains(search) |
        filtered_teachers['school'].str.contains(search, case=False, na=False)
    ]

if not result.empty:
    r = result.iloc[0]
    st.write(r)
    # إصلاح TypeError في القاعات
    halls['number'] = halls['number'].fillna("").astype(str)
    halls['hall'] = halls['hall'].fillna("").astype(str)
    hall_options = [""] + [f"{n} - {h}" for n,h in zip(halls['number'], halls['hall'])]
    selected_hall = st.selectbox("اختر القاعة", hall_options, key="assign_hall")
    role_assign = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"], key="assign_role")

    if st.button("تعيين", key="assign_btn"):
        hall_name = selected_hall.split(" - ")[1] if selected_hall else ""
        c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?",
                  (hall_name, role_assign, r['id']))
        conn.commit()
        st.success("تم التعيين")
        st.experimental_rerun()

    if st.button("إلغاء", key="cancel_btn"):
        c.execute("UPDATE teachers SET hall='' WHERE id=?", (r['id'],))
        conn.commit()
        st.warning("تم الإلغاء")
        st.experimental_rerun()

# =====================================
# 🏛️ إدارة القاعات + الطباعة
# =====================================
st.subheader("🏛️ القاعات")
if not halls.empty:
    hall_select = st.selectbox("اختر قاعة", halls['hall'], key="hall_select")
    hall_teachers = teachers[teachers['hall']==hall_select]
    st.dataframe(hall_teachers)

    if not hall_teachers.empty:
        single = st.selectbox("طباعة لمعلم", hall_teachers['name'], key="print_single")

        if st.button("📄 طباعة فردي", key="print_btn"):
            row = hall_teachers[hall_teachers['name']==single].iloc[0]
            doc = Document("doc.docx")
            for p in doc.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace("<NAME>", row['name'])\
                                       .replace("<ID>", row['id'])\
                                       .replace("<CITY>", row['city'])\
                                       .replace("<WORKPLACE>", row['school'])\
                                       .replace("<HALL_NAME>", hall_select)
            path = f"{row['name']}.docx"
            doc.save(path)
            with open(path, "rb") as f:
                st.download_button("تحميل", f, key="download_single")

        if st.button("📦 طباعة جماعي", key="print_all"):
            os.makedirs("out", exist_ok=True)
            files = []
            for _, row in hall_teachers.iterrows():
                doc = Document("doc.docx")
                for p in doc.paragraphs:
                    for run in p.runs:
                        run.text = run.text.replace("<NAME>", row['name'])\
                                           .replace("<HALL_NAME>", hall_select)
                path = f"out/{row['name']}.docx"
                doc.save(path)
                files.append(path)
            zip_path = "out/all.zip"
            with zipfile.ZipFile(zip_path, 'w') as z:
                for f in files:
                    z.write(f, os.path.basename(f))
            with open(zip_path, "rb") as f:
                st.download_button("تحميل الكل", f, key="download_all")
