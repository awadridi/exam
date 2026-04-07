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

# RTL دعم العربية
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

# تحميل التعيينات
if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    if 'هوية' in assignments.columns:
        teachers = teachers.merge(assignments, on="هوية", how="left")
else:
    teachers['قاعة مختارة'] = None

if 'قاعة مختارة' not in teachers.columns:
    teachers['قاعة مختارة'] = None

# ================== دالة توليد الملف ==================
def generate_doc(row, hall_info):
    doc = Document(empty_doc)
    for p in doc.paragraphs:
        for run in p.runs:
            run.text = run.text.replace("<NAME>", row['اسم'])\
                               .replace("<ID>", str(row['هوية']))\
                               .replace("<CITY>", row['سكن'])\
                               .replace("<WORKPLACE>", row['مدرسة'])\
                               .replace("<HALL_NAME>", hall_info['قاعة'])\
                               .replace("<HALL_LOCATION>", hall_info['بلد'])
    return doc

# ================== عرض المعلمين الموزعين ==================
assigned = teachers.dropna(subset=['قاعة مختارة'])

if not assigned.empty:
    st.subheader("📌 المعلمون الموزعون")
    t_remove = st.selectbox("اختر معلم:", assigned['اسم'])
    if st.button("❌ إلغاء التكليف"):
        teachers.loc[teachers['اسم'] == t_remove, 'قاعة مختارة'] = None
        teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
        st.warning(f"تم إلغاء تكليف {t_remove}")

# ================== البحث بالاسم ==================
st.subheader("🔍 البحث بالاسم")
search_name = st.text_input("اكتب اسم المعلم")

if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, case=False, na=False)]

    if not results.empty:
        selected = st.selectbox("اختر المعلم:", results['اسم'])
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="name")

        if hall != "اختر القاعة...":

            hall_info = halls[halls['قاعة'] == hall]
            if hall_info.empty:
                st.error("خطأ في القاعة")
                st.stop()
            hall_info = hall_info.iloc[0]

            if st.button("✅ تعيين"):
                teachers.loc[teachers['اسم'] == selected, 'قاعة مختارة'] = hall
                teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
                st.success("تم التعيين")

            if st.button("📄 توليد الكتاب"):
                row = teachers[teachers['اسم'] == selected].iloc[0]
                doc = generate_doc(row, hall_info)

                os.makedirs("تكليفات", exist_ok=True)
                path = f"تكليفات/{row['اسم']}.docx"
                doc.save(path)

                with open(path, "rb") as f:
                    st.download_button("⬇️ تحميل", f, file_name=row['اسم'] + ".docx")

            if st.button("❌ إلغاء"):
                teachers.loc[teachers['اسم'] == selected, 'قاعة مختارة'] = None
                teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
                st.warning("تم الإلغاء")

# ================== البحث بالهوية ==================
st.subheader("🆔 البحث بالهوية")
search_id = st.text_input("رقم الهوية")

if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]

    if not result.empty:
        row = result.iloc[0]
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="id")

        if hall != "اختر القاعة...":

            hall_info = halls[halls['قاعة'] == hall]
            if hall_info.empty:
                st.error("خطأ في القاعة")
                st.stop()
            hall_info = hall_info.iloc[0]

            if st.button("✅ تعيين بالهوية"):
                teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = hall
                teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
                st.success("تم التعيين")

            if st.button("📄 توليد بالهوية"):
                doc = generate_doc(row, hall_info)

                os.makedirs("تكليفات", exist_ok=True)
                path = f"تكليفات/{row['اسم']}.docx"
                doc.save(path)

                with open(path, "rb") as f:
                    st.download_button("⬇️ تحميل", f, file_name=row['اسم'] + ".docx")

            if st.button("❌ إلغاء بالهوية"):
                teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = None
                teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
                st.warning("تم الإلغاء")

# ================== حسب القاعة ==================
st.subheader("🏫 إدارة القاعات")

hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
hall = st.selectbox("اختر القاعة", hall_options)

if hall != "اختر القاعة...":

    selected_teachers = teachers[teachers['قاعة مختارة'] == hall]

    if not selected_teachers.empty:
        st.write(f"عدد المعلمين: {len(selected_teachers)}")

        t_remove = st.selectbox("إلغاء من القاعة:", selected_teachers['اسم'])
        if st.button("❌ حذف من القاعة"):
            teachers.loc[teachers['اسم'] == t_remove, 'قاعة مختارة'] = None
            teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
            st.warning("تم الحذف")

    if st.button("📦 توليد كل القاعة"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []

        for _, row in selected_teachers.iterrows():
            hall_info = halls[halls['قاعة'] == hall].iloc[0]
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
st.subheader("📦 جميع التكليفات")

if st.button("توليد الكل"):
    os.makedirs("تكليفات", exist_ok=True)
    files = []

    for _, row in teachers.dropna(subset=['قاعة مختارة']).iterrows():
        hall_info = halls[halls['قاعة'] == row['قاعة مختارة']].iloc[0]
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
