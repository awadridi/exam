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

# --- البحث بالاسم وتوليد كتاب ---
search_name = st.text_input("ابحث عن المراقب بالاسم:")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        selected_teacher = st.selectbox("اختر المراقب:", results['اسم'])
        hall_choice = st.selectbox("اختر القاعة:", halls['قاعة'], key="hall_by_name")
        if st.button("توليد كتاب التكليف بالاسم", key="generate_by_name"):
            teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)

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

# --- البحث برقم الهوية وتوليد كتاب ---
search_id = st.text_input("اكتب رقم هوية المعلم:")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        hall_choice = st.selectbox("اختر أو غيّر القاعة:", halls['قاعة'], key="hall_by_id")
        if st.button("توليد كتاب التكليف بالهوية", key="generate_by_id"):
            teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assignments_file, index=False)

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
