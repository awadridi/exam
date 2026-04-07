import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# ================== إعدادات ==================
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
assignments_file = "assignments.xlsx"  # حفظ التعيينات
PASSWORD = st.secrets["PASSWORD"]
empty_doc = "doc.docx"  # قالب Word

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
st.success("تم تسجيل الدخول ✅")

# ================== تحميل البيانات ==================
teachers = pd.read_excel(exam_file)
halls = pd.read_excel(halls_file)

# إعادة تسمية الأعمدة لتوافق الكود
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

# تحويل أرقام الوظائف إلى نصوص
role_map = {1: "رئيس قاعة", 2: "سكرتير", 3: "آذن", 4: "مراقب"}
teachers['وظيفة'] = teachers['وظيفة'].map(role_map).fillna("مراقب")

# ================== تحميل التعيينات السابقة ==================
if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    teachers = teachers.merge(assignments, on="هوية", how="left")
    if 'قاعة مختارة' not in teachers.columns:
        teachers['قاعة مختارة'] = None
    if 'مهمة' not in teachers.columns:
        teachers['مهمة'] = teachers['وظيفة']
else:
    teachers['قاعة مختارة'] = None
    teachers['مهمة'] = teachers['وظيفة']

# ================== دوال ==================
def save_assignments():
    teachers[['هوية','قاعة مختارة','مهمة']].to_excel(assignments_file, index=False)

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

# ================== إضافة معلم جديد ==================
st.subheader("➕ إضافة معلم جديد")
with st.expander("إضافة"):
    name = st.text_input("الاسم", key="new_name")
    id = st.text_input("الهوية", key="new_id")
    school = st.text_input("المدرسة", key="new_school")
    city = st.text_input("السكن", key="new_city")
    role = st.selectbox("الوظيفة", ["مراقب","رئيس قاعة","آذن","سكرتير"], key="new_role")

    if st.button("إضافة المعلم"):
        if name and id and school and city:
            if id in teachers['هوية'].astype(str).values:
                st.warning("رقم الهوية موجود مسبقًا!")
            else:
                new_row = pd.DataFrame([{
                    'هوية': id,
                    'اسم': name,
                    'مدرسة': school,
                    'سكن': city,
                    'وظيفة': role,
                    'قاعة مختارة': None,
                    'مهمة': role
                }])
                teachers = pd.concat([teachers, new_row], ignore_index=True)
                save_assignments()
                st.success("تمت إضافة المعلم بنجاح ✅")
        else:
            st.warning("املأ جميع الحقول قبل الإضافة!")

# ================== البحث والتعيين ==================
st.subheader("🔍 البحث والتعيين")
search_name = st.text_input("ابحث عن المعلم بالاسم:", key="search_name")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        selected_teacher = st.selectbox("اختر المعلم:", results['اسم'], key="select_teacher")
        row = teachers[teachers['اسم'] == selected_teacher].iloc[0]

        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="select_hall")
        role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"], index=["مراقب","رئيس قاعة","آذن","سكرتير"].index(row['مهمة']), key="select_role")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين"):
                if hall != "اختر القاعة..." and role != "":
                    teachers.loc[teachers['اسم']==selected_teacher, ['قاعة مختارة','مهمة']] = [hall, role]
                    try:
                        save_assignments()
                        st.success(f"تم تعيين {selected_teacher} في {hall} كمهمة {role}")
                    except Exception as e:
                        st.error(f"حدث خطأ أثناء حفظ التعيين: {e}")
                else:
                    st.warning("اختر القاعة والمهمة قبل التعيين!")

        with col2:
            if st.button("توليد كتاب Word"):
                if row['قاعة مختارة']:
                    doc = generate_doc(row)
                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)
                    with open(path, "rb") as f:
                        st.download_button("تحميل Word", f, file_name=f"{row['اسم']}.docx")
                else:
                    st.warning("المعلم غير معين في أي قاعة بعد!")

        with col3:
            if st.button("إلغاء التعيين"):
                teachers.loc[teachers['اسم']==selected_teacher, ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {selected_teacher}")

# ================== إدارة القاعات ==================
st.subheader("🏫 إدارة القاعات")
hall_filter = st.selectbox("اختر القاعة:", [""] + list(halls['قاعة']), key="hall_filter")
if hall_filter:
    assigned = teachers[teachers['قاعة مختارة'] == hall_filter]
    st.write(f"عدد المعلمين في القاعة: {len(assigned)}")
    if not assigned.empty:
        teacher_to_remove = st.selectbox("اختر معلم لإلغاء تكليفه:", assigned['اسم'], key="remove_teacher")
        if st.button("إلغاء تكليف هذا المعلم"):
            teachers.loc[teachers['اسم']==teacher_to_remove, ['قاعة مختارة','مهمة']] = [None, assigned.loc[assigned['اسم']==teacher_to_remove, 'وظيفة'].values[0]]
            save_assignments()
            st.warning(f"تم إلغاء تكليف {teacher_to_remove} من القاعة {hall_filter}")

    if st.button("توليد كتب Word + ZIP لهذه القاعة"):
        os.makedirs("تكليفات", exist_ok=True)
        files = []
        for _, row in assigned.iterrows():
            doc = generate_doc(row)
            path = f"تكليفات/{row['اسم']}.docx"
            doc.save(path)
            files.append(path)
        zip_path = f"تكليفات/{hall_filter}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in files:
                zipf.write(f, os.path.basename(f))
        with open(zip_path, "rb") as f:
            st.download_button("تحميل ZIP", f, file_name=f"{hall_filter}.zip")

# ================== توليد جميع الكتب دفعة واحدة ==================
if st.button("توليد جميع كتب التكليف دفعة واحدة"):
    os.makedirs("تكليفات", exist_ok=True)
    files = []
    for _, row in teachers.dropna(subset=['قاعة مختارة']).iterrows():
        doc = generate_doc(row)
        path = f"تكليفات/{row['اسم']}.docx"
        doc.save(path)
        files.append(path)
    zip_path = "تكليفات/جميع_التكليفات.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for f in files:
            zipf.write(f, os.path.basename(f))
    with open(zip_path, "rb") as f:
        st.download_button("تحميل جميع التكليفات", f, file_name="جميع_التكليفات.zip")
