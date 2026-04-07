import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import os
import zipfile

# --- قاعدة البيانات ---
conn = sqlite3.connect("data.db", check_same_thread=False)
c = conn.cursor()

# إنشاء الجداول
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

password = st.text_input("كلمة المرور", type="password")

if st.button("دخول"):
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

# --- إضافة معلم ---
with st.expander("➕ إضافة معلم"):
    name = st.text_input("الاسم")
    idd = st.text_input("الهوية")
    school = st.text_input("المدرسة")
    city = st.text_input("السكن")
    phone = st.text_input("الجوال")
    role = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"])

    if st.button("حفظ المعلم"):
        try:
            c.execute("INSERT INTO teachers VALUES (?,?,?,?,?,?,?)",
                      (idd, name, school, city, phone, role, ""))
            conn.commit()
            st.success("تم الحفظ ✅")
        except:
            st.warning("المعلم موجود مسبقًا")

# --- تعديل / حذف ---
st.subheader("✏️ تعديل أو حذف معلم")

if not teachers.empty:
    selected = st.selectbox("اختر معلم", teachers['name'])

    row = teachers[teachers['name']==selected].iloc[0]

    new_name = st.text_input("الاسم", row['name'])
    new_phone = st.text_input("الجوال", row['phone'])

    if st.button("تحديث"):
        c.execute("UPDATE teachers SET name=?, phone=? WHERE id=?",
                  (new_name, new_phone, row['id']))
        conn.commit()
        st.success("تم التعديل")

    if st.button("حذف"):
        c.execute("DELETE FROM teachers WHERE id=?", (row['id'],))
        conn.commit()
        st.warning("تم الحذف")

# --- البحث والتعيين ---
st.subheader("🔍 البحث والتعيين")

search = st.text_input("ابحث")

if search:
    result = teachers[
        teachers['name'].str.contains(search, case=False, na=False) |
        teachers['id'].astype(str).str.contains(search)
    ]

    if not result.empty:
        r = result.iloc[0]

        st.write(r)

        hall_options = [""] + list(halls['number'] + " - " + halls['hall'])
        selected_hall = st.selectbox("اختر القاعة", hall_options)

        role = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"])

        if st.button("تعيين"):
            hall_name = selected_hall.split(" - ")[1] if selected_hall else ""
            c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?",
                      (hall_name, role, r['id']))
            conn.commit()
            st.success("تم التعيين")

        if st.button("إلغاء"):
            c.execute("UPDATE teachers SET hall='' WHERE id=?", (r['id'],))
            conn.commit()
            st.warning("تم الإلغاء")

# --- القاعات ---
st.subheader("🏛️ إدارة القاعات")

if not halls.empty:
    hall_select = st.selectbox("اختر قاعة", halls['hall'])

    hall_teachers = teachers[teachers['hall']==hall_select]

    st.dataframe(hall_teachers)

    if not hall_teachers.empty:

        # طباعة فردي
        single = st.selectbox("طباعة لمعلم", hall_teachers['name'])

        if st.button("طباعة كتاب فردي"):
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
                st.download_button("تحميل", f)

        # طباعة جماعي
        if st.button("طباعة كل القاعة"):
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
                st.download_button("تحميل الكل", f)
