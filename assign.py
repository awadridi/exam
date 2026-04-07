import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import platform

# استدعاء pythoncom فقط إذا كنت على Windows
if platform.system() == "Windows":
    import pythoncom
import streamlit as st
import streamlit as st

# قراءة كلمة السر من Secrets
PASSWORD = st.secrets["PASSWORD"]

st.title("تطبيق التكليف")
password_input = st.text_input("أدخل كلمة المرور:", type="password")

if password_input != PASSWORD:
    st.error("كلمة المرور غير صحيحة")
    st.stop()


# --- قراءة الملفات ---
teachers = pd.read_excel("exam.xlsx").rename(columns={
    'هوية': 'هوية',
    'اسم المعلم': 'اسم',
    'اسم المدرسة': 'مدرسة',
    'سكن': 'سكن',
    'الوظيفة': 'وظيفة'
})
halls = pd.read_excel("halls.xlsx").rename(columns={
    'رقم': 'رقم',
    'اسم القاعة': 'قاعة',
    'البلد': 'بلد'
})

assign_file = "assignments.xlsx"
if os.path.exists(assign_file):
    assignments = pd.read_excel(assign_file)
    if 'هوية' in assignments.columns:
        teachers = teachers.merge(assignments, on="هوية", how="left")
    else:
        teachers['قاعة مختارة'] = None
else:
    teachers['قاعة مختارة'] = None

if 'قاعة مختارة' not in teachers.columns:
    teachers['قاعة مختارة'] = None

st.title("إدارة كتب التكليف")

# --- البحث بالاسم وتوزيع ---
search_name = st.text_input("ابحث عن المراقب بالاسم:")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        selected_teacher = st.selectbox("اختر المراقب:", results['اسم'])
        hall_choice = st.selectbox("اختر القاعة:", halls['قاعة'])
        if st.button("إضافة/تحديث التوزيع"):
            teachers.loc[teachers['اسم'] == selected_teacher, 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assign_file, index=False)
            st.success(f"تم تعيين {hall_choice} لـ {selected_teacher}")
    else:
        st.warning("لا يوجد نتائج مطابقة")

# --- البحث برقم الهوية ---
search_id = st.text_input("اكتب رقم هوية المعلم:")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        current_hall = row['قاعة مختارة'] if pd.notna(row['قاعة مختارة']) else None

        if current_hall:
            st.info(f"المعلم {row['اسم']} معين في القاعة: {current_hall}")
        else:
            st.warning(f"المعلم {row['اسم']} لم يتم تعيين قاعة له بعد")

        hall_choice = st.selectbox("اختر أو غيّر القاعة:", halls['قاعة'], key="hall_by_id")
        if st.button("تحديث القاعة"):
            teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'] = hall_choice
            teachers[['هوية','قاعة مختارة']].dropna().to_excel(assign_file, index=False)
            st.success(f"تم تعيين القاعة {hall_choice} للمعلم {row['اسم']}")

        if st.button("توليد كتاب التكليف لهذا المعلم (PDF)"):
            hall_info = halls[halls['قاعة'] == teachers.loc[teachers['هوية'] == row['هوية'], 'قاعة مختارة'].iloc[0]].iloc[0]
            doc = Document("كتاب التكليف مراقبة فارغ.docx")
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
            pdf_path = f"تكليفات/تكليف_{row['اسم']}.pdf"
            doc.save(word_path)

            pythoncom.CoInitialize()
            convert(word_path, pdf_path)
            pythoncom.CoUninitialize()

            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="⬇️ تنزيل كتاب التكليف PDF",
                    data=f,
                    file_name=f"تكليف_{row['اسم']}.pdf",
                    mime="application/pdf"
                )
            st.success(f"تم توليد كتاب التكليف PDF للمعلم {row['اسم']}")
    else:
        st.error("لا يوجد معلم بهذا الرقم")

# --- توليد كتب لكل المراقبين في قاعة معينة ---
hall_filter = st.selectbox("اختر قاعة لتوليد كتب التكليف لكل المراقبين فيها:", halls['قاعة'])
if st.button("توليد كتب التكليف لهذه القاعة (PDF)"):
    selected_teachers = teachers[teachers['قاعة مختارة'] == hall_filter]
    if not selected_teachers.empty:
        os.makedirs("تكليفات", exist_ok=True)
        word_files = []
        for _, row in selected_teachers.iterrows():
            hall_info = halls[halls['قاعة'] == hall_filter].iloc[0]
            doc = Document("كتاب التكليف مراقبة فارغ.docx")
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

        # تحويل كل الملفات دفعة واحدة
        pythoncom.CoInitialize()
        convert("تكليفات")
        pythoncom.CoUninitialize()

        # توليد ملف ZIP للتنزيل
        zip_path = f"تكليفات/تكليفات_{hall_filter}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for word_file in word_files:
                pdf_file = word_file.replace(".docx", ".pdf")
                zipf.write(pdf_file, os.path.basename(pdf_file))

        with open(zip_path, "rb") as f:
            st.download_button(
                label="⬇️ تنزيل جميع كتب القاعة كـ ZIP",
                data=f,
                file_name=f"تكليفات_{hall_filter}.zip",
                mime="application/zip"
            )
        st.success(f"تم توليد كتب التكليف لجميع المراقبين في القاعة {hall_filter} بصيغة PDF")
    else:
        st.warning("لا يوجد مراقبين معينين لهذه القاعة")

# --- توليد جميع الكتب دفعة واحدة ---
if st.button("توليد جميع كتب التكليف (PDF)"):
    os.makedirs("تكليفات", exist_ok=True)
    word_files = []
    for _, row in teachers.dropna(subset=['قاعة مختارة']).iterrows():
        hall_info = halls[halls['قاعة'] == row['قاعة مختارة']].iloc[0]
        doc = Document("كتاب التكليف مراقبة فارغ.docx")
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

    # تحويل كل الملفات دفعة واحدة
    pythoncom.CoInitialize()
    convert("تكليفات")
    pythoncom.CoUninitialize()

    # توليد ملف ZIP للتنزيل
    zip_path = "تكليفات/جميع_التكليفات.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for word_file in word_files:
            pdf_file = word_file.replace(".docx", ".pdf")
            zipf.write(pdf_file, os.path.basename(pdf_file))

    with open(zip_path, "rb") as f:
        st.download_button(
            label="⬇️ تنزيل جميع كتب التكليف كـ ZIP",
            data=f,
            file_name="جميع_التكليفات.zip",
            mime="application/zip"
        )
    st.success("تم توليد جميع كتب التكليف بصيغة PDF")
