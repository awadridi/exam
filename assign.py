import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# قراءة أسماء الملفات من Secrets
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
assignments_file = st.secrets["ASSIGNMENTS_FILE"]
empty_doc = "doc.docx"   # اسم ملف القالب الجديد

# كلمة السر
PASSWORD = st.secrets["PASSWORD"]

# واجهة التطبيق
st.title("تطبيق التكليف")
password_input = st.text_input("أدخل كلمة المرور:", type="password")
if password_input != PASSWORD:
    st.error("كلمة المرور غير صحيحة")
    st.stop()
st.success("تم تسجيل الدخول بنجاح ✅")

# --- قراءة الملفات ---
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

if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    if 'هوية' in assignments.columns:
        teachers = teachers.merge(assignments, on="هوية", how="left")
    else:
        teachers['قاعة مختارة'] = None
else:
    teachers['قاعة مختارة'] = None

if 'قاعة مختارة' not in teachers.columns:
    teachers['قاعة مختارة'] = None

# --- البحث بالاسم ---
search_name = st.text_input("ابحث عن المراقب بالاسم:")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        selected_teacher = st.selectbox("اختر المراقب:", results['اسم'])
        hall_choice = st.selectbox("اختر القاعة:", halls['قاعة'], key="hall_by_name")

        if st.button("تعيين القاعة بالاسم", key="assign_by_name"):
            teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
            st.success(f"تم تعيين القاعة {hall_choice} للمعلم {selected_teacher}")

       if st.button("توليد كتاب التكليف بالاسم", key="generate_by_name"):
    row = teachers[teachers['اسم'] == selected_teacher].iloc[0]
    hall_info = halls[halls['قاعة'] == hall_choice].iloc[0]
    doc = Document(empty_doc)
    for p in doc.paragraphs:
        for run in p.runs:
            run.text = run.text.replace("<NAME>", row['اسم'])\
                               .replace("<ID>", str(row['هوية']))\
                               .replace("<CITY>", row['سكن'])\
                               .replace("<WORKPLACE>", row['مدرسة'])\
                               .replace("<HALL_NAME>", hall_info['قاعة'])\
                               .replace("<HALL_LOCATION>", hall_info['بلد'])
    os.makedirs("تكليفات", exist_ok=True)
    word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
    doc.save(word_path)

    with open(word_path, "rb") as f:
        st.download_button(
            label="⬇️ تنزيل كتاب التكليف Word",
            data=f,
            file_name=f"تكليف_{row['اسم']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


            with open(word_path, "rb") as f:
                st.download_button(
                    label="⬇️ تنزيل كتاب التكليف Word",
                    data=f,
                    file_name=f"تكليف_{selected_teacher}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        if st.button("إلغاء التعيين بالاسم", key="remove_by_name"):
            teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = None
            teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
            st.warning(f"تم إلغاء تكليف المعلم {selected_teacher}")

# --- البحث برقم الهوية ---
search_id = st.text_input("اكتب رقم هوية المعلم:")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        hall_choice = st.selectbox("اختر أو غيّر القاعة:", halls['قاعة'], key="hall_by_id")

        if st.button("تعيين القاعة بالهوية", key="assign_by_id"):
            teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
            st.success(f"تم تعيين القاعة {hall_choice} للمعلم {row['اسم']}")

        if st.button("توليد كتاب التكليف بالهوية", key="generate_by_id"):
            hall_info = halls[halls['قاعة'] == hall_choice].iloc[0]
            doc = Document(empty_doc)
            for p in doc.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace("<NAME>", row['اسم'])\
                                       .replace("<ID>", str(row['هوية']))\
                                       .replace("<CITY>", row['سكن'])\
                                       .replace("<WORKPLACE>", row['مدرسة'])\
                                       .replace("<HALL_NAME>", hall_info['قاعة'])\
                                       .replace("<HALL_LOCATION>", hall_info['بلد'])
            os.makedirs("تكليفات", exist_ok=True)
            word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
            doc.save(word_path)

            with open(word_path, "rb") as f:
                st.download_button(
                    label="⬇️ تنزيل كتاب التكليف Word",
                    data=f,
                    file_name=f"تكليف_{row['اسم']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        if st.button("إلغاء التعيين بالهوية", key="remove_by_id"):
            teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = None
            teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
            st.warning(f"تم إلغاء تكليف المعلم {row['اسم']}")

# --- إدارة القاعة ---
hall_filter = st.selectbox("اختر قاعة:", halls['قاعة'], key="hall_select")

teacher_in_hall = teachers[teachers['قاعة مختارة'] == hall_filter]
if not teacher_in_hall.empty:
    teacher_to_remove = st.selectbox("اختر معلم لإلغاء تكليفه:", teacher_in_hall['اسم'], key="remove_teacher_select")
    if st.button("إلغاء تكليف هذا المعلم", key="remove_teacher_button"):
        teachers.loc[teachers['اسم'] == teacher_to_remove, 'قاعة مختارة'] = None
        teachers[['هوية','قاعة مختارة']].to_excel(assignments_file, index=False)
        st.warning(f"تم إلغاء تكليف المعلم {teacher_to_remove} من القاعة {hall_filter}")

if st.button("توليد كتب التكليف لهذه القاعة", key="generate_by_hall"):
    selected_teachers = teachers[teachers['قاعة مختارة'] == hall_filter]
    if not selected_teachers.empty:
        os.makedirs("تكليفات", exist_ok=True)
        word_files = []
        for _, row in selected_teachers.iterrows():
            hall_info = halls[halls['قاعة'] == hall_filter].iloc[0]
            doc = Document(empty_doc)
            for p in doc.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace("<NAME>", row['اسم'])\
                                       .replace("<ID>", str(row['هوية']))\
                                       .replace("<CITY>", row['سكن'])\
                                       .replace("<WORKPLACE>", row['مدرسة'])\
                                       .replace("<HALL_NAME>", hall_info['قاعة'])\
                                       .replace("<HALL_LOCATION>", hall_info['بلد'])
            word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
            doc.save(word_path)
            word_files.append(word_path)

        zip_path = f"تكليفات/تكليفات_{hall_filter}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for word_file in word_files:
                zipf.write(word_file, os.path.basename(word_file))

        with open(zip_path, "rb") as f:
            st.download_button(
                label="⬇️ تنزيل جميع كتب القاعة كـ ZIP",
                data=f,
                file_name=f"تكليفات_{hall_filter}.zip",
                mime="application/zip"
            )

# --- توليد جميع الكتب دفعة واحدة ---
if st.button("توليد جميع كتب التكليف", key="generate_all"):
    os.makedirs("تكليفات", exist_ok=True)
    word_files = []
    for _, row in teachers.dropna(subset=['قاعة مختارة']).iterrows():
        hall_info = halls[halls['قاعة'] == row['قاعة مختارة']].iloc[0]
        doc = Document(empty_doc)
        for p in doc.paragraphs:
            for run in p.runs:
                run.text = run.text.replace("<NAME>", row['اسم'])\
                                   .replace("<ID>", str(row['هوية']))\
                                   .replace("<CITY>", row['سكن'])\
                                   .replace("<WORKPLACE>", row['مدرسة'])\
                                   .replace("<HALL_NAME>", hall_info['قاعة'])\
                                   .replace("<HALL_LOCATION>", hall_info['بلد'])
        word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
        doc.save(word_path)
        word_files.append(word_path)

    zip_path = "تكليفات/جميع_التكليفات.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for word_file in word_files:
            zipf.write(word_file, os.path.basename(word_file))

    with open(zip_path, "rb") as f:
        st.download_button(
            label="⬇️ تنزيل جميع كتب التكليف كـ ZIP",
            data=f,
            file_name="جميع_التكليفات.zip",
            mime="application/zip"
        )
