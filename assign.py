import streamlit as st
import pandas as pd
from docx import Document
import sqlite3
import os
import zipfile

# ================== إعدادات ==================
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
PASSWORD = st.secrets["PASSWORD"]
empty_doc = "doc.docx"

# قاعدة البيانات
conn = sqlite3.connect("data.db", check_same_thread=False)
c = conn.cursor()

# إنشاء جدول إذا مش موجود
c.execute("""
CREATE TABLE IF NOT EXISTS assignments (
    id TEXT PRIMARY KEY,
    hall TEXT
)
""")
conn.commit()

# ================== واجهة ==================
st.title("📋 تطبيق التكليف")

# RTL
st.markdown("""
<style>
body { direction: RTL; text-align: right; }
</style>
""", unsafe_allow_html=True)

# تسجيل الدخول
password_input = st.text_input("أدخل كلمة المرور:", type="password")
if password_input != PASSWORD:
    st.stop()

# ================== تحميل البيانات ==================
teachers = pd.read_excel(exam_file)
halls = pd.read_excel(halls_file)

teachers = teachers.rename(columns={
    'هوية': 'هوية',
    'اسم المعلم': 'اسم',
    'اسم المدرسة': 'مدرسة',
    'سكن': 'سكن'
})

halls = halls.rename(columns={
    'اسم القاعة': 'قاعة',
    'البلد': 'بلد'
})

# ================== تحميل التعيينات من DB ==================
c.execute("SELECT * FROM assignments")
rows = c.fetchall()

assign_dict = {r[0]: r[1] for r in rows}

teachers['قاعة مختارة'] = teachers['هوية'].astype(str).map(assign_dict).fillna("")

# ================== دوال ==================
def assign_teacher(teacher_id, hall):
    c.execute("REPLACE INTO assignments (id, hall) VALUES (?, ?)", (str(teacher_id), hall))
    conn.commit()

def remove_teacher(teacher_id):
    c.execute("DELETE FROM assignments WHERE id=?", (str(teacher_id),))
    conn.commit()

def get_hall_info(hall_name):
    return halls[halls['قاعة'] == hall_name].iloc[0]

def generate_doc(row, hall_info):
    doc = Document(empty_doc)
    for p in doc.paragraphs:
        for run in p.runs:
            run.text = run.text.replace("<NAME>", str(row['اسم']))\
                               .replace("<ID>", str(row['هوية']))\
                               .replace("<CITY>", str(row['سكن']))\
                               .replace("<WORKPLACE>", str(row['مدرسة']))\
                               .replace("<HALL_NAME>", str(hall_info['قاعة']))\
                               .replace("<HALL_LOCATION>", str(hall_info['بلد']))
    return doc

# ================== البحث ==================
st.subheader("🔍 البحث")

search = st.text_input("اسم أو هوية")

if search:
    results = teachers[
        teachers['اسم'].astype(str).str.contains(search, case=False, na=False) |
        teachers['هوية'].astype(str).str.contains(search, na=False)
    ]

    if not results.empty:
        name = st.selectbox("اختر المعلم", results['اسم'])

        row = teachers[teachers['اسم'] == name].iloc[0]
        teacher_id = str(row['هوية'])

        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة", hall_options)

        if hall != "اختر القاعة...":

            col1, col2, col3 = st.columns(3)

            with col1:
                if st.button("✅ تعيين"):
                    assign_teacher(teacher_id, hall)
                    st.success("تم الحفظ ✅")
                    st.rerun()

            with col2:
                if st.button("📄 توليد"):
                    hall_info = get_hall_info(hall)
                    doc = generate_doc(row, hall_info)

                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)

                    with open(path, "rb") as f:
                        st.download_button("تحميل", f, file_name=row['اسم'] + ".docx")

            with col3:
                if st.button("❌ إلغاء"):
                    remove_teacher(teacher_id)
                    st.warning("تم الحذف")
                    st.rerun()

# ================== إدارة القاعات ==================
st.subheader("🏫 القاعات")

hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
hall = st.selectbox("اختر القاعة", hall_options, key="hall_main")

if hall != "اختر القاعة...":
    selected = teachers[teachers['قاعة مختارة'] == hall]

    st.write(f"عدد المعلمين: {len(selected)}")

    if not selected.empty:
        name = st.selectbox("حذف", selected['اسم'])

        if st.button("❌ حذف من القاعة"):
            teacher_id = teachers[teachers['اسم'] == name]['هوية'].iloc[0]
            remove_teacher(teacher_id)
            st.rerun()

    if st.button("📦 توليد القاعة"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []

        for _, row in selected.iterrows():
            hall_info = get_hall_info(hall)
            doc = generate_doc(row, hall_info)

            path = f"تكليفات/{row['اسم']}.docx"
            doc.save(path)
            files.append(path)

        zip_path = f"تكليفات/{hall}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in files:
                zipf.write(f, os.path.basename(f))

        with open(zip_path, "rb") as f:
            st.download_button("تحميل ZIP", f, file_name=hall + ".zip")
