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

# --- البحث بالاسم وتوزيع ---
search_name = st.text_input("ابحث عن المراقب بالاسم:")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        selected_teacher = st.selectbox("اختر المراقب:", results['اسم'])
        hall_choice = st.selectbox("اختر القاعة:", halls['قاعة'])
        if st.button("إضافة/تحديث التوزيع"):
            teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
if st.button("إضافة/تحديث التوزيع"):
    teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = hall_choice
    teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)
    st.success(f"تم تعيين {hall_choice} لـ {selected_teacher}")

    # توليد كتاب التكليف مباشرة
    hall_info = halls[halls['قاعة'] == hall_choice].iloc[0]
    doc = Document(empty_doc)
    for p in doc.paragraphs:
        for run in p.runs:
            run.text = run.text.replace("<NAME>", selected_teacher)\
                               .replace("<HALL_NAME>", hall_info['قاعة'])\
                               .replace("<HALL_LOCATION>", hall_info['بلد'])
    os.makedirs("تكليفات", exist_ok=True)
    word_path = f"تكليفات/تكليف_{selected_teacher}.docx"
    doc.save(word_path)

    with open(word_path, "rb") as f:
        st.download_button(
            label="⬇️ تنزيل كتاب التكليف Word",
            data=f,
            file_name=f"تكليف_{selected_teacher}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


# --- البحث برقم الهوية ---
search_id = st.text_input("اكتب رقم هوية المعلم:")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        hall_choice = st.selectbox("اختر أو غيّر القاعة:", halls['قاعة'], key="hall_by_id")
        if st.button("تعيين القاعة"):
            teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)

        if st.button("توليد كتاب التكليف لهذا المعلم (Word)"):
            selected_hall = teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة']
            if not selected_hall.empty:
                hall_info = halls[halls['قاعة'] == selected_hall.iloc[0]].iloc[0]
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

# --- توليد كتب لكل المراقبين في قاعة معينة ---
hall_filter = st.selectbox("اختر قاعة لتوليد كتب التكليف لكل المراقبين فيها:", halls['قاعة'])
if st.button("توليد كتب التكليف لهذه القاعة (Word)"):
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
if st.button("توليد جميع كتب التكليف (Word)"):
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
