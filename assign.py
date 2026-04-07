import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# --- إعداد الملفات من Secrets ---
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
assignments_file = st.secrets["ASSIGNMENTS_FILE"]
empty_doc = "doc.docx"
PASSWORD = st.secrets["PASSWORD"]

# --- واجهة تسجيل الدخول ---
st.title("🎓 نظام إدارة التكليف")
password_input = st.text_input("أدخل كلمة المرور:", type="password")
if password_input != PASSWORD:
    st.error("كلمة المرور غير صحيحة")
    st.stop()
st.success("تم تسجيل الدخول ✅")

# --- قراءة البيانات ---
teachers = pd.read_excel(exam_file).rename(columns={
    'هوية':'هوية',
    'اسم المعلم':'اسم',
    'اسم المدرسة':'مدرسة',
    'سكن':'سكن',
    'رقم الجوال':'جوال',
    'الوظيفة':'وظيفة'
})
halls = pd.read_excel(halls_file).rename(columns={
    'رقم':'رقم',
    'اسم القاعة':'قاعة',
    'البلد':'بلد'
})

# --- التأكد من الأعمدة الضرورية ---
for col in ['قاعة مختارة', 'مهمة']:
    if col not in teachers.columns:
        teachers[col] = ""
    else:
        teachers[col] = teachers[col].fillna("").astype(str)

# --- تحميل التعيينات السابقة إذا موجود ---
if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    teachers = teachers.merge(assignments, on='هوية', how='left', suffixes=('','_old'))
    if 'قاعة مختارة_old' in teachers.columns:
        teachers['قاعة مختارة'] = teachers['قاعة مختارة_old'].fillna(teachers['قاعة مختارة'])
        teachers.drop(columns=['قاعة مختارة_old'], inplace=True)
    if 'مهمة_old' in teachers.columns:
        teachers['مهمة'] = teachers['مهمة_old'].fillna(teachers['مهمة'])
        teachers.drop(columns=['مهمة_old'], inplace=True)

# --- دالة لحفظ التعيينات ---
def save_assignments():
    teachers[['هوية','قاعة مختارة','مهمة']].to_excel(assignments_file, index=False)

# --- إضافة معلم جديد داخل Expander ---
with st.expander("➕ إضافة معلم جديد"):
    with st.form("add_teacher_form"):
        name = st.text_input("اسم المعلم")
        identity = st.text_input("رقم الهوية")
        school = st.text_input("اسم المدرسة")
        residence = st.text_input("السكن")
        phone = st.text_input("رقم الجوال")
        role = st.selectbox("المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"])
        submitted = st.form_submit_button("إضافة المعلم")
        if submitted:
            if identity in teachers['هوية'].astype(str).values:
                st.warning("المعلم موجود مسبقًا")
            else:
                new_teacher = {
                    'هوية': identity,
                    'اسم': name,
                    'مدرسة': school,
                    'سكن': residence,
                    'جوال': phone,
                    'وظيفة': role,
                    'قاعة مختارة': "",
                    'مهمة': role
                }
                teachers = pd.concat([teachers, pd.DataFrame([new_teacher])], ignore_index=True)
                save_assignments()
                st.success(f"تم إضافة المعلم {name}")

# --- البحث بالاسم أو رقم الهوية ---
st.subheader("🔍 البحث بالاسم أو رقم الهوية")
search_input = st.text_input("اكتب الاسم أو رقم الهوية:")

matched_teachers = pd.DataFrame()
if search_input:
    matched_teachers = teachers[
        teachers['اسم'].str.contains(search_input, case=False, na=False) |
        (teachers['هوية'].astype(str) == search_input)
    ]

if not matched_teachers.empty:
    if len(matched_teachers) > 1:
        selected_teacher_name = st.selectbox("المعلمون المطابقون:", matched_teachers['اسم'])
        row = matched_teachers[matched_teachers['اسم'] == selected_teacher_name].iloc[0]
    else:
        row = matched_teachers.iloc[0]
    st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']} - القاعة: {row['قاعة مختارة']}")

    # --- تعيين القاعة والمهمة ---
    hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
    hall = st.selectbox("اختر القاعة:", hall_options)
    role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"])

    if hall != "اختر القاعة...":
        if st.button("تعيين القاعة والمهمة لهذا المعلم"):
            teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [str(hall), str(role)]
            save_assignments()
            st.success(f"تم تعيين {row['اسم']} في القاعة {hall} كمهمة {role}")

    if st.button("إلغاء تكليف هذا المعلم"):
        teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = ["",""]
        save_assignments()
        st.warning(f"تم إلغاء تكليف {row['اسم']}")

# --- إدارة القاعات ---
st.subheader("🏛️ إدارة القاعات")
hall_options_all = ["اختر القاعة..."] + list(halls['قاعة'])
selected_hall = st.selectbox("اختر قاعة:", hall_options_all, key="hall_manage")

if selected_hall != "اختر القاعة...":
    teachers_in_hall = teachers[teachers['قاعة مختارة']==selected_hall]
    if not teachers_in_hall.empty:
        remove_teacher = st.selectbox("اختر معلم لإلغاء تكليفه من القاعة:", teachers_in_hall['اسم'], key="remove_teacher_select")
        if st.button("إلغاء تكليف هذا المعلم من القاعة"):
            teachers.loc[teachers['اسم']==remove_teacher, ['قاعة مختارة','مهمة']] = ["",""]
            save_assignments()
            st.warning(f"تم إلغاء تكليف {remove_teacher} من القاعة {selected_hall}")

    if st.button("توليد كتب التكليف لهذه القاعة"):
        os.makedirs("تكليفات", exist_ok=True)
        word_files = []
        for _, row in teachers_in_hall.iterrows():
            hall_info = halls[halls['قاعة'] == selected_hall].iloc[0]
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

        zip_path = f"تكليفات/تكليفات_{selected_hall}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for word_file in word_files:
                zipf.write(word_file, os.path.basename(word_file))

        with open(zip_path, "rb") as f:
            st.download_button(
                label="⬇️ تنزيل جميع كتب القاعة كـ ZIP",
                data=f,
                file_name=f"تكليفات_{selected_hall}.zip",
                mime="application/zip"
            )
