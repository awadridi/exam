import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# --- إعداد الملفات ---
exam_file = "teachers.xlsx"
halls_file = "halls.xlsx"
assignments_file = "assignments.xlsx"
empty_doc = "doc.docx"
PASSWORD = "1234"

# --- تسجيل الدخول ---
st.title("🎓 نظام إدارة التكليف")
password_input = st.text_input("أدخل كلمة المرور:", type="password")

if password_input != PASSWORD:
    st.stop()

st.success("تم تسجيل الدخول ✅")

# --- تحميل البيانات مرة واحدة ---
if "teachers" not in st.session_state:

    if os.path.exists(exam_file):
        teachers = pd.read_excel(exam_file)
    else:
        teachers = pd.DataFrame(columns=[
            'هوية','اسم','مدرسة','سكن','جوال','وظيفة','قاعة مختارة','مهمة'
        ])

    if os.path.exists(halls_file):
        halls = pd.read_excel(halls_file)
    else:
        halls = pd.DataFrame(columns=['قاعة','بلد'])

    # أعمدة مضمونة
    for col in ['قاعة مختارة','مهمة']:
        if col not in teachers.columns:
            teachers[col] = ""

    # تحميل التعيينات
    if os.path.exists(assignments_file):
        assignments = pd.read_excel(assignments_file)
        teachers = teachers.merge(assignments, on='هوية', how='left', suffixes=('','_old'))

        if 'قاعة مختارة_old' in teachers.columns:
            teachers['قاعة مختارة'] = teachers['قاعة مختارة_old'].fillna(teachers['قاعة مختارة'])
            teachers.drop(columns=['قاعة مختارة_old'], inplace=True)

        if 'مهمة_old' in teachers.columns:
            teachers['مهمة'] = teachers['مهمة_old'].fillna(teachers['مهمة'])
            teachers.drop(columns=['مهمة_old'], inplace=True)

    st.session_state.teachers = teachers
    st.session_state.halls = halls

teachers = st.session_state.teachers
halls = st.session_state.halls

# --- حفظ البيانات ---
def save_all():
    st.session_state.teachers.to_excel(exam_file, index=False)
    st.session_state.teachers[['هوية','قاعة مختارة','مهمة']].to_excel(assignments_file, index=False)

# --- إضافة معلم ---
with st.expander("➕ إضافة معلم جديد"):
    with st.form("add_teacher"):
        name = st.text_input("الاسم")
        identity = st.text_input("رقم الهوية")
        school = st.text_input("المدرسة")
        residence = st.text_input("السكن")
        phone = st.text_input("رقم الجوال")
        role = st.selectbox("المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"])

        submitted = st.form_submit_button("إضافة")

        if submitted:
            if identity in teachers['هوية'].astype(str).values:
                st.warning("المعلم موجود مسبقًا")
            else:
                new_teacher = pd.DataFrame([{
                    'هوية': identity,
                    'اسم': name,
                    'مدرسة': school,
                    'سكن': residence,
                    'جوال': phone,
                    'وظيفة': role,
                    'قاعة مختارة': "",
                    'مهمة': role
                }])

                st.session_state.teachers = pd.concat([teachers, new_teacher], ignore_index=True)
                save_all()
                st.success("تمت الإضافة ✅")

# --- البحث ---
st.subheader("🔍 البحث")
search = st.text_input("ابحث بالاسم أو الهوية")

if search:
    result = teachers[
        teachers['اسم'].str.contains(search, case=False, na=False) |
        (teachers['هوية'].astype(str) == search)
    ]

    if not result.empty:
        row = result.iloc[0]

        st.write(f"👤 {row['اسم']} | 📞 {row['جوال']} | 🏫 {row['مدرسة']} | 🏛️ {row['قاعة مختارة']}")

        hall = st.selectbox("اختر القاعة", [""] + list(halls['قاعة']))
        role = st.selectbox("اختر المهمة", ["مراقب","رئيس قاعة","آذن","سكرتير"])

        if st.button("تعيين"):
            st.session_state.teachers.loc[
                teachers['هوية']==row['هوية'],
                ['قاعة مختارة','مهمة']
            ] = [hall, role]

            save_all()
            st.success("تم التعيين ✅")

        if st.button("إلغاء التكليف"):
            st.session_state.teachers.loc[
                teachers['هوية']==row['هوية'],
                ['قاعة مختارة','مهمة']
            ] = ["",""]

            save_all()
            st.warning("تم الإلغاء ⚠️")

# --- إدارة القاعات ---
st.subheader("🏛️ القاعات")

selected_hall = st.selectbox("اختر القاعة", [""] + list(halls['قاعة']))

if selected_hall:
    hall_teachers = teachers[teachers['قاعة مختارة']==selected_hall]

    st.write("المعلمين داخل القاعة:")
    st.dataframe(hall_teachers[['اسم','مهمة','جوال']])

    if not hall_teachers.empty:
        remove = st.selectbox("إزالة معلم", hall_teachers['اسم'])

        if st.button("إزالة من القاعة"):
            st.session_state.teachers.loc[
                teachers['اسم']==remove,
                ['قاعة مختارة','مهمة']
            ] = ["",""]

            save_all()
            st.success("تمت الإزالة")

    # --- توليد ملفات ---
    if st.button("توليد كتب التكليف"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []

        for _, row in hall_teachers.iterrows():
            hall_info = halls[halls['قاعة']==selected_hall].iloc[0]

            doc = Document(empty_doc)

            for p in doc.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace("<NAME>", row['اسم'])\
                                       .replace("<ID>", str(row['هوية']))\
                                       .replace("<CITY>", row['سكن'])\
                                       .replace("<WORKPLACE>", row['مدرسة'])\
                                       .replace("<HALL_NAME>", hall_info['قاعة'])\
                                       .replace("<HALL_LOCATION>", hall_info['بلد'])

            path = f"تكليفات/{row['اسم']}.docx"
            doc.save(path)
            files.append(path)

        zip_path = f"تكليفات/{selected_hall}.zip"
        with zipfile.ZipFile(zip_path, 'w') as z:
            for f in files:
                z.write(f, os.path.basename(f))

        with open(zip_path, "rb") as f:
            st.download_button("تحميل ZIP", f, file_name="تكليفات.zip")
