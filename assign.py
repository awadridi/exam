import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import os
import zipfile

# --- قاعدة البيانات ---
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

# --- تسجيل الدخول ---
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

# --- تحميل البيانات ---
def get_teachers():
    return pd.read_sql("SELECT * FROM teachers", conn)

def get_halls():
    return pd.read_sql("SELECT * FROM halls", conn)

teachers = get_teachers()
halls = get_halls()

# =====================================
# ➕ إضافة معلم (تم إصلاح مشكلة TAB)
# =====================================
with st.expander("➕ إضافة معلم جديد", expanded=True):
    with st.form("add_teacher_form"):

        name = st.text_input("الاسم", key="add_name")
        idd = st.text_input("الهوية", key="add_id")
        school = st.text_input("المدرسة", key="add_school")
        city = st.text_input("السكن", key="add_city")
        phone = st.text_input("الجوال", key="add_phone")

        role = st.selectbox(
            "المهمة",
            ["مراقب","رئيس قاعة","آذن","سكرتير"],
            key="add_role"
        )

        submitted = st.form_submit_button("💾 حفظ المعلم")

        if submitted:
            if not idd or not name:
                st.warning("⚠️ لازم تعبّي الاسم والهوية")
            else:
                try:
                    c.execute("INSERT INTO teachers VALUES (?,?,?,?,?,?,?)",
                              (idd, name, school, city, phone, role, ""))
                    conn.commit()
                    st.success("تم الحفظ ✅")

                    # تفريغ الحقول بعد الحفظ
                    for k in ["add_name","add_id","add_school","add_city","add_phone"]:
                        st.session_state[k] = ""

                except:
                    st.error("❌ المعلم موجود مسبقًا")

# =====================================
# ✏️ تعديل / حذف
# =====================================
st.subheader("✏️ تعديل أو حذف معلم")

if not teachers.empty:
    selected = st.selectbox("اختر معلم", teachers['name'], key="edit_select")

    row = teachers[teachers['name']==selected].iloc[0]

    new_name = st.text_input("الاسم الجديد", row['name'], key="edit_name")
    new_phone = st.text_input("الجوال الجديد", row['phone'], key="edit_phone")

    if st.button("تحديث", key="update_btn"):
        c.execute("UPDATE teachers SET name=?, phone=? WHERE id=?",
                  (new_name, new_phone, row['id']))
        conn.commit()
        st.success("تم التعديل")

    if st.button("حذف", key="delete_btn"):
        c.execute("DELETE FROM teachers WHERE id=?", (row['id'],))
        conn.commit()
        st.warning("تم الحذف")

# =====================================
# 🔍 البحث والتعيين
# =====================================
st.subheader("🔍 البحث والتعيين")

search = st.text_input("ابحث", key="search_box")

if search:
    result = teachers[
        teachers['name'].str.contains(search, case=False, na=False) |
        teachers['id'].astype(str).str.contains(search)
    ]

    if not result.empty:
        r = result.iloc[0]

        st.write(r)

        hall_options = [""] + list(halls['number'].astype(str) + " - " + halls['hall'])
        selected_hall = st.selectbox("اختر القاعة", hall_options, key="assign_hall")

        role_assign = st.selectbox(
            "المهمة",
            ["مراقب","رئيس قاعة","آذن","سكرتير"],
            key="assign_role"
        )

        if st.button("تعيين", key="assign_btn"):
            hall_name = selected_hall.split(" - ")[1] if selected_hall else ""
            c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?",
                      (hall_name, role_assign, r['id']))
            conn.commit()
            st.success("تم التعيين")

        if st.button("إلغاء", key="cancel_btn"):
            c.execute("UPDATE teachers SET hall='' WHERE id=?", (r['id'],))
            conn.commit()
            st.warning("تم الإلغاء")

# =====================================
# 🏛️ القاعات + الطباعة
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
