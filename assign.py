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
with st.expander("إضافة"):
    name = st.text_input("الاسم", key="new_name")
    id = st.text_input("الهوية", key="new_id")
    school = st.text_input("المدرسة", key="new_school")
    city = st.text_input("السكن", key="new_city")
    phone = st.text_input("رقم الجوال", key="new_phone")
    role = st.selectbox("الوظيفة", ["مراقب","رئيس قاعة","آذن","سكرتير"], key="new_role")

    if st.button("إضافة المعلم"):
        if name and id and school and city and phone:
            if id in teachers['هوية'].astype(str).values:
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
        else:
            st.warning("املأ جميع الحقول قبل الإضافة!")

# ================== البحث بالهوية ==================
st.subheader("🔍 البحث والتعيين بالهوية")
search_id = st.text_input("أدخل رقم الهوية:", key="search_id")
if search_id:
    result = teachers[teachers['هوية'].astype(str) == search_id]
    if not result.empty:
        row = result.iloc[0]
        st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']}")
