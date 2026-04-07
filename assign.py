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

# ================== البحث بالاسم أو الهوية في حقل واحد ==================
st.subheader("🔍 البحث بالاسم أو رقم الهوية")
search_input = st.text_input("اكتب الاسم أو رقم الهوية:")

if search_input:
    # البحث بالهوية أولاً إذا كانت كل الأرقام
    if search_input.isdigit():
        results = teachers[teachers['هوية'].astype(str).str.contains(search_input)]
    else:
        # البحث بالاسم جزئيًا
        results = teachers[teachers['اسم'].str.contains(search_input, na=False)]

    if results.empty:
        st.warning("لم يتم العثور على أي معلم مطابق")
    elif len(results) == 1:
        row = results.iloc[0]
        st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']} - القاعة: {row['قاعة مختارة']}")
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="single_search_hall")
        role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"], index=["مراقب","رئيس قاعة","آذن","سكرتير"].index(row['مهمة']), key="single_search_role")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين", key="single_search_assign"):
                if hall != "اختر القاعة...":
                    teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [hall, role]
                    save_assignments()
                    st.success(f"تم تعيين {row['اسم']} في {hall} كمهمة {role}")
                else:
                    st.warning("اختر القاعة قبل التعيين!")
        with col2:
            if st.button("توليد Word", key="single_search_word"):
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
            if st.button("إلغاء التعيين", key="single_search_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")

    else:
        # أكثر من معلم مطابق → عرض قائمة للاختيار
        selected_name = st.selectbox("المعلمون المطابقون:", results['اسم'].tolist())
        row = results[results['اسم'] == selected_name].iloc[0]
        st.write(f"المعلم: {row['اسم']} - المدرسة: {row['مدرسة']} - المهمة: {row['مهمة']} - الجوال: {row['جوال']} - القاعة: {row['قاعة مختارة']}")
        hall_options = ["اختر القاعة..."] + list(halls['قاعة'])
        hall = st.selectbox("اختر القاعة:", hall_options, key="multi_search_hall")
        role = st.selectbox("اختر المهمة:", ["مراقب","رئيس قاعة","آذن","سكرتير"], index=["مراقب","رئيس قاعة","آذن","سكرتير"].index(row['مهمة']), key="multi_search_role")

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تعيين", key="multi_search_assign"):
                if hall != "اختر القاعة...":
                    teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [hall, role]
                    save_assignments()
                    st.success(f"تم تعيين {row['اسم']} في {hall} كمهمة {role}")
                else:
                    st.warning("اختر القاعة قبل التعيين!")
        with col2:
            if st.button("توليد Word", key="multi_search_word"):
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
            if st.button("إلغاء التعيين", key="multi_search_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")
