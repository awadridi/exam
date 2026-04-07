import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile

# ================== إعدادات ==================
exam_file = st.secrets["EXAM_FILE"]
halls_file = st.secrets["HALLS_FILE"]
assignments_file = "assignments.xlsx"
PASSWORD = st.secrets["PASSWORD"]
empty_doc = "doc.docx"

# ================== واجهة ==================
st.title("🎓 نظام إدارة التكليف")
st.markdown("""
<style>
body { direction: RTL; text-align: right; font-family: Arial; }
</style>
""", unsafe_allow_html=True)

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

# إضافة عمود رقم الجوال إذا لم يكن موجودًا
if 'جوال' not in teachers.columns:
    teachers['جوال'] = None

halls = halls.rename(columns={'اسم القاعة': 'قاعة','البلد': 'بلد'})

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
    if 'جوال' not in teachers.columns:
        teachers['جوال'] = None
else:
    teachers['قاعة مختارة'] = None
    teachers['مهمة'] = teachers['وظيفة']
    teachers['جوال'] = None

# ================== دوال ==================
def save_assignments():
    teachers[['هوية','اسم','مدرسة','سكن','وظيفة','جوال','قاعة مختارة','مهمة']].to_excel(assignments_file, index=False)

def generate_doc(row):
    if not row['قاعة مختارة']:
        return None
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
                               .replace("<ROLE>", str(row['مهمة']))\
                               .replace("<PHONE>", str(row['جوال']))
    return doc

# ================== إضافة معلم جديد ==================
st.subheader("➕ إضافة معلم جديد")
with st.expander("إضافة معلم"):
    name = st.text_input("الاسم", key="new_name")
    id = st.text_input("الهوية", key="new_id")
    school = st.text_input("المدرسة", key="new_school")
    city = st.text_input("السكن", key="new_city")
    phone = st.text_input("رقم الجوال", key="new_phone")
    role = st.selectbox("الوظيفة", ["مراقب","رئيس قاعة","آذن","سكرتير"], key="new_role")

    if st.button("إضافة المعلم"):
        if not all([name, id, school, city, phone]):
            st.warning("جميع الحقول إلزامية، يجب إدخال رقم الجوال!")
        elif id in teachers['هوية'].astype(str).values:
            st.warning("رقم الهوية موجود مسبقًا!")
        else:
            new_row = pd.DataFrame([{
                'هوية': id,
                'اسم': name,
                'مدرسة': school,
                'سكن': city,
                'وظيفة': role,
                'جوال': phone,
                'قاعة مختارة': None,
                'مهمة': role
            }])
            teachers = pd.concat([teachers, new_row], ignore_index=True)
            save_assignments()
            st.success("تمت إضافة المعلم بنجاح ✅")

# ================== البحث بالاسم ==================
st.subheader("🔍 البحث بالاسم")
search_name = st.text_input("ابحث عن المراقب بالاسم:")
if search_name:
    results = teachers[teachers['اسم'].str.contains(search_name, na=False)]
    if not results.empty:
        row = results.iloc[0]
        st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']} - القاعة: {row['قاعة مختارة']}")
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="search_name_hall")
        role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"], index=["مراقب","رئيس قاعة","آذن","سكرتير"].index(row['مهمة']), key="search_name_role")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين", key="search_name_assign"):
                if hall != "اختر القاعة...":
                    teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [hall, role]
                    save_assignments()
                    st.success(f"تم تعيين {row['اسم']} في {hall} كمهمة {role}")
                else:
                    st.warning("اختر القاعة قبل التعيين!")
        with col2:
            if st.button("توليد Word", key="search_name_word"):
                doc = generate_doc(row)
                if doc:
                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)
                    with open(path, "rb") as f:
                        st.download_button("تحميل Word", f, file_name=f"{row['اسم']}.docx")
                else:
                    st.warning("المعلم غير معين في أي قاعة!")
        with col3:
            if st.button("إلغاء التعيين", key="search_name_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")

# ================== البحث برقم الهوية ==================
st.subheader("🔍 البحث برقم الهوية")
search_id = st.text_input("اكتب رقم هوية المعلم:")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']} - القاعة: {row['قاعة مختارة']}")
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="search_id_hall")
        role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"], index=["مراقب","رئيس قاعة","آذن","سكرتير"].index(row['مهمة']), key="search_id_role")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين", key="search_id_assign"):
                if hall != "اختر القاعة...":
                    teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [hall, role]
                    save_assignments()
                    st.success(f"تم تعيين {row['اسم']} في {hall} كمهمة {role}")
                else:
                    st.warning("اختر القاعة قبل التعيين!")
        with col2:
            if st.button("توليد Word", key="search_id_word"):
                doc = generate_doc(row)
                if doc:
                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)
                    with open(path, "rb") as f:
                        st.download_button("تحميل Word", f, file_name=f"{row['اسم']}.docx")
                else:
                    st.warning("المعلم غير معين في أي قاعة!")
        with col3:
            if st.button("إلغاء التعيين", key="search_id_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")

# ================== إدارة القاعات ==================
st.subheader("🏫 إدارة القاعات")
hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
selected_hall = st.selectbox("اختر القاعة لإدارة المعلمين:", hall_options, key="hall_manage")

if selected_hall != "اختر القاعة...":
    hall_teachers = teachers[teachers['قاعة مختارة']==selected_hall]
    if not hall_teachers.empty:
        teacher_to_remove = st.selectbox("اختر معلم لإلغاء تكليفه:", hall_teachers['اسم'], key="remove_teacher_select")
        if st.button("إلغاء تكليف هذا المعلم", key="remove_teacher_button"):
            teachers.loc[teachers['اسم']==teacher_to_remove, 'قاعة مختارة'] = None
            save_assignments()
            st.warning(f"تم إلغاء تكليف {teacher_to_remove} من القاعة {selected_hall}")

        if st.button("توليد كتب التكليف لهذه القاعة", key="generate_by_hall"):
            os.makedirs("تكليفات", exist_ok=True)
            word_files = []
            for _, row in hall_teachers.iterrows():
                doc = generate_doc(row)
                if doc:
                    path = f"تكليفات/تكليف_{row['اسم']}.docx"
                    doc.save(path)
                    word_files.append(path)
            if word_files:
                zip_path = f"تكليفات/تكليفات_{selected_hall}.zip"
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in word_files:
                        zipf.write(file, os.path.basename(file))
                with open(zip_path, "rb") as f:
                    st.download_button("تحميل ZIP للقاعة", f, file_name=f"تكليفات_{selected_hall}.zip")

    st.write("➕ إضافة معلم موجود لهذه القاعة")
    unassigned_teachers = teachers[teachers['قاعة مختارة'].isna()]
    if not unassigned_teachers.empty:
        teacher_to_assign = st.selectbox("اختر معلم للتعيين:", unassigned_teachers['اسم'], key="assign_teacher_select")
        role_options = ["مراقب","رئيس قاعة","آذن","سكرتير"]
        assign_role = st.selectbox("اختر المهمة:", role_options, key="assign_role_select")
        if st.button("تعيين المعلم للقاعة", key="assign_teacher_button"):
            teachers.loc[teachers['اسم']==teacher_to_assign, ['قاعة مختارة','مهمة']] = [selected_hall, assign_role]
            save_assignments()
            st.success(f"تم تعيين {teacher_to_assign} في {selected_hall} كمهمة {assign_role}")
