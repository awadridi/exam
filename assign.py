import streamlit as st
import pandas as pd
from docx import Document
from fpdf import FPDF
import sqlite3
import os
import zipfile

# ================== إعدادات ==================
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
PASSWORD = st.secrets["PASSWORD"]
empty_doc = "doc.docx"

# ================== قاعدة البيانات ==================
conn = sqlite3.connect("data.db", check_same_thread=False)
c = conn.cursor()

c.execute("""
CREATE TABLE IF NOT EXISTS assignments (
    id TEXT PRIMARY KEY,
    hall TEXT,
    role TEXT
)
""")
conn.commit()

# ================== واجهة ==================
st.title("🎓 نظام إدارة التكليف")
st.markdown("""
<style>
body { direction: RTL; text-align: right; font-family: Arial; }
</style>
""", unsafe_allow_html=True)

# تسجيل دخول
password_input = st.text_input("كلمة المرور", type="password")
if password_input != PASSWORD:
    st.stop()

# ================== تحميل البيانات ==================
teachers = pd.read_excel(exam_file)
halls = pd.read_excel(halls_file)

teachers = teachers.rename(columns={
    'هوية': 'هوية',
    'اسم المعلم': 'اسم',
    'اسم المدرسة': 'مدرسة',
    'سكن': 'سكن',
    'الوظيفة': 'وظيفة'
})

halls = halls.rename(columns={
    'اسم القاعة': 'قاعة',
    'البلد': 'بلد'
})

# ================== تحويل أرقام الوظائف إلى نصوص ==================
role_map = {
    1: "رئيس قاعة",
    2: "سكرتير",
    3: "آذن",
    4: "مراقب"
}
teachers['وظيفة'] = teachers['وظيفة'].map(role_map).fillna("مراقب")

# ================== تحميل التعيينات ==================
c.execute("SELECT * FROM assignments")
rows = c.fetchall()

assign_dict = {}
for r in rows:
    if len(r) == 3:
        assign_dict[r[0]] = (r[1], r[2])

teachers['قاعة مختارة'] = teachers['هوية'].astype(str).map(lambda x: assign_dict.get(x, ("",""))[0])
teachers['مهمة'] = teachers['هوية'].astype(str).map(lambda x: assign_dict.get(x, ("",""))[1])

# ================== دوال ==================
def assign_teacher(id, hall, role):
    c.execute("REPLACE INTO assignments VALUES (?, ?, ?)", (id, hall, role))
    conn.commit()

def remove_teacher(id):
    c.execute("DELETE FROM assignments WHERE id=?", (id,))
    conn.commit()

def generate_doc(row):
    hall_info = halls[halls['قاعة'] == row['قاعة مختارة']].iloc[0]
    doc = Document(empty_doc)
    for p in doc.paragraphs:
        for run in p.runs:
            run.text = run.text.replace("<NAME>", str(row['اسم']))\
                               .replace("<ID>", str(row['هوية']))\
                               .replace("<CITY>", str(row['سكن']))\
                               .replace("<WORKPLACE>", str(row['مدرسة']))\
                               .replace("<HALL_NAME>", str(row['قاعة مختارة']))\
                               .replace("<HALL_LOCATION>", str(hall_info['بلد']))\
                               .replace("<ROLE>", str(row['مهمة']))
    return doc

def generate_pdf(all_teachers):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for _, row in all_teachers.iterrows():
        pdf.multi_cell(0, 8, txt=f"الاسم: {row['اسم']}\n"
                                  f"الهوية: {row['هوية']}\n"
                                  f"المدرسة: {row['مدرسة']}\n"
                                  f"السكن: {row['سكن']}\n"
                                  f"الوظيفة: {row['وظيفة']}\n"
                                  f"قاعة: {row['قاعة مختارة']}\n"
                                  f"المهمة: {row['مهمة']}\n\n")
    os.makedirs("تكليفات", exist_ok=True)
    pdf_path = "تكليفات/جميع_المعلمين.pdf"
    pdf.output(pdf_path)
    return pdf_path

# ================== إضافة معلم ==================
st.subheader("➕ إضافة معلم جديد")
with st.expander("إضافة"):
    name = st.text_input("الاسم")
    id = st.text_input("الهوية")
    school = st.text_input("المدرسة")
    city = st.text_input("السكن")
    role = st.selectbox("الوظيفة", ["مراقب","رئيس قاعة","آذن","سكرتير"])

    if st.button("إضافة المعلم"):
        new_row = pd.DataFrame([{
            'هوية': id,
            'اسم': name,
            'مدرسة': school,
            'سكن': city,
            'وظيفة': role
        }])
        teachers = pd.concat([teachers, new_row], ignore_index=True)
        st.success("تمت الإضافة")

# ================== توزيع تلقائي ==================
st.subheader("⚡ توزيع تلقائي")
max_per_hall = st.number_input("عدد المعلمين لكل قاعة", 1, 50, 5)

if st.button("توزيع تلقائي"):
    halls_list = list(halls['قاعة'])
    i = 0
    for _, row in teachers.iterrows():
        teacher_id = str(row['هوية'])
        if teacher_id not in assign_dict:
            hall = halls_list[i % len(halls_list)]
            assign_teacher(teacher_id, hall, row['وظيفة'])
            i += 1
    st.success("تم التوزيع")
    st.experimental_rerun()

# ================== البحث ==================
st.subheader("🔍 بحث")
search = st.text_input("اسم أو هوية")
if search:
    results = teachers[
        teachers['اسم'].astype(str).str.contains(search, case=False) |
        teachers['هوية'].astype(str).str.contains(search)
    ]
    if not results.empty:
        name = st.selectbox("اختر", results['اسم'])
        row = teachers[teachers['اسم'] == name].iloc[0]
        teacher_id = str(row['هوية'])
        hall = st.selectbox("قاعة", [""] + list(halls['قاعة']))
        role = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"])
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين"):
                assign_teacher(teacher_id, hall, role)
                st.experimental_rerun()
        with col2:
            if st.button("توليد Word"):
                doc = generate_doc(row)
                os.makedirs("تكليفات", exist_ok=True)
                path = f"تكليفات/{row['اسم']}.docx"
                doc.save(path)
                with open(path, "rb") as f:
                    st.download_button("تحميل Word", f, file_name=row['اسم'] + ".docx")
        with col3:
            if st.button("حذف"):
                remove_teacher(teacher_id)
                st.experimental_rerun()

# ================== القاعات ==================
st.subheader("🏫 القاعات")
hall = st.selectbox("اختر القاعة", [""] + list(halls['قاعة']))
if hall:
    selected = teachers[teachers['قاعة مختارة'] == hall]
    st.write(f"عدد: {len(selected)}")
    if st.button("توليد القاعة Word + ZIP"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []
        for _, row in selected.iterrows():
            doc = generate_doc(row)
            path = f"تكليفات/{row['اسم']}.docx"
            doc.save(path)
            files.append(path)
        zip_path = f"تكليفات/{hall}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in files:
                zipf.write(f, os.path.basename(f))
        with open(zip_path, "rb") as f:
            st.download_button("تحميل ZIP", f, file_name=hall + ".zip")

# ================== توليد PDF لجميع المعلمين ==================
st.subheader("📄 توليد PDF لجميع المعلمين")
if st.button("توليد PDF"):
    pdf_path = generate_pdf(teachers)
    with open(pdf_path, "rb") as f:
        st.download_button("تحميل PDF لجميع المعلمين", f, file_name="جميع_المعلمين.pdf")
