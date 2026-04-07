import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# ================== إعدادات ==================
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
assignments_file = st.secrets["ASSIGNMENTS_FILE"]
PASSWORD = st.secrets["PASSWORD"]
empty_doc = "doc.docx"

# ================== واجهة ==================
st.title("📋 تطبيق التكليف")

# دعم العربية RTL
st.markdown("""
<style>
body { direction: RTL; text-align: right; }
.stTextInput, .stSelectbox, .stButton { direction: RTL; text-align: right; }
</style>
""", unsafe_allow_html=True)

# تسجيل الدخول
password_input = st.text_input("أدخل كلمة المرور:", type="password")
if password_input != PASSWORD:
    st.error("كلمة المرور غير صحيحة")
    st.stop()
st.success("تم تسجيل الدخول بنجاح ✅")

# ================== قراءة البيانات ==================
teachers = pd.read_excel(exam_file).rename(columns={
    'هوية': 'هوية',
    'اسم المعلم': 'اسم',
    'اسم المدرسة': 'مدرسة',
    'سكن': 'سكن',
    'الوظيفة': 'وظيفة'
})

halls = pd.read_excel(halls_file).rename(columns={
    'رقم': 'رقم',
    'اسم القاعة': 'قاعة',
    'البلد': 'بلد'
})

# ================== إصلاح المشكلة الأساسية ==================
if 'قاعة مختارة' not in teachers.columns:
    teachers['قاعة مختارة'] = ""

# إجبار العمود أن يكون نص (حل الخطأ)
teachers['قاعة مختارة'] = teachers['قاعة مختارة'].astype("object")

# تحميل التعيينات
if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    if 'هوية' in assignments.columns:
        assignments['قاعة مختارة'] = assignments['قاعة مختارة'].astype("object")
        teachers = teachers.merge(assignments, on="هوية", how="left", suffixes=('', '_old'))
        teachers['قاعة مختارة'] = teachers['قاعة مختارة_old'].fillna(teachers['قاعة مختارة'])
        teachers.drop(columns=['قاعة مختارة_old'], inplace=True)

# ================== دوال ==================
def save_assignments():
    teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)

def get_hall_info(hall_name):
    hall_info = halls[halls['قاعة'] == hall_name]
    if hall_info.empty:
        st.error("⚠️ خطأ في بيانات القاعة")
        st.stop()
    return hall_info.iloc[0]

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
st.subheader("🔍 البحث عن معلم")

search = st.text_input("اكتب الاسم أو الهوية")

if search:
    results = teachers[
        teachers['اسم'].astype(str).str.contains(search, case=False, na=False) |
        teachers['هوية'].astype(str).str.contains(search, na=False)
    ]

    if not results.empty:
        selected_name = st.selectbox("اختر المعلم", results['اسم'])

        row = teachers[teachers['اسم'] == selected_name].iloc[0]
        teacher_id = row['هوية']

        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة", hall_options)

        if hall != "اختر القاعة...":
            hall_info = get_hall_info(hall)

            col1, col2, col3 = st.columns(3)

            # تعيين
            with col1:
                if st.button("✅ تعيين"):
                    teachers.loc[teachers['هوية'] == teacher_id, 'قاعة مختارة'] = hall
                    save_assignments()
                    st.success("تم التعيين")

            # توليد كتاب
            with col2:
                if st.button("📄 توليد"):
                    doc = generate_doc(row, hall_info)
                    os.makedirs("تكليفات", exist_ok=True)

                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)

                    with open(path, "rb") as f:
                        st.download_button("⬇️ تحميل", f, file_name=row['اسم'] + ".docx")

            # إلغاء
            with col3:
                if st.button("❌ إلغاء"):
                    teachers.loc[teachers['هوية'] == teacher_id, 'قاعة مختارة'] = ""
                    save_assignments()
                    st.warning("تم الإلغاء")

# ================== إدارة القاعات ==================
st.subheader("🏫 إدارة القاعات")

hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
hall = st.selectbox("اختر القاعة")

if hall != "اختر القاعة...":
    selected_teachers = teachers[teachers['قاعة مختارة'] == hall]

    st.write(f"عدد المعلمين: {len(selected_teachers)}")

    if not selected_teachers.empty:
        t_remove = st.selectbox("اختر للحذف", selected_teachers['اسم'])

        if st.button("❌ حذف من القاعة"):
            teacher_id = teachers[teachers['اسم'] == t_remove]['هوية'].iloc[0]
            teachers.loc[teachers['هوية'] == teacher_id, 'قاعة مختارة'] = ""
            save_assignments()
            st.warning("تم الحذف")

    if st.button("📦 توليد ملفات القاعة"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []

        for _, row in selected_teachers.iterrows():
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
            st.download_button("⬇️ تحميل ZIP", f, file_name=hall + ".zip")

# ================== توليد الكل ==================
st.subheader("📦 كل التكليفات")

if st.button("توليد الكل"):
    os.makedirs("تكليفات", exist_ok=True)
    files = []

    all_assigned = teachers[teachers['قاعة مختارة'] != ""]

    for _, row in all_assigned.iterrows():
        hall_info = get_hall_info(row['قاعة مختارة'])
        doc = generate_doc(row, hall_info)

        path = f"تكليفات/{row['اسم']}.docx"
        doc.save(path)
        files.append(path)

    zip_path = "تكليفات/all.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for f in files:
            zipf.write(f, os.path.basename(f))

    with open(zip_path, "rb") as f:
        st.download_button("⬇️ تحميل الكل", f, file_name="all.zip")
