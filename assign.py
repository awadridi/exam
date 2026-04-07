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

# تحويل أرقام الوظائف إلى نصوص
role_map = {1: "رئيس قاعة", 2: "سكرتير", 3: "آذن", 4: "مراقب"}
teachers['وظيفة'] = teachers['وظيفة'].map(role_map).fillna("مراقب")

# الأعمدة الأساسية
if 'قاعة مختارة' not in teachers.columns:
    teachers['قاعة مختارة'] = None
if 'مهمة' not in teachers.columns:
    teachers['مهمة'] = teachers['وظيفة']

# القاعات
halls = halls.rename(columns={'اسم القاعة': 'قاعة','البلد': 'بلد'})

# تحميل التعيينات السابقة
if os.path.exists(assignments_file):
    assignments = pd.read_excel(assignments_file)
    teachers = teachers.merge(assignments, on="هوية", how="left", suffixes=("","_old"))
    if 'قاعة مختارة_old' in teachers.columns:
        teachers['قاعة مختارة'] = teachers['قاعة مختارة_old']
        teachers.drop(columns=['قاعة مختارة_old'], inplace=True)
    if 'مهمة_old' in teachers.columns:
        teachers['مهمة'] = teachers['مهمة_old']
        teachers.drop(columns=['مهمة_old'], inplace=True)

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
                'قاعة مختارة': None,
                'مهمة': role
            }
            teachers = pd.concat([teachers, pd.DataFrame([new_teacher])], ignore_index=True)
            save_assignments()
            st.success(f"تم إضافة المعلم {name}")

# ================== البحث الموحد ==================
st.subheader("🔍 البحث بالاسم أو رقم الهوية")
search_input = st.text_input("اكتب الاسم أو رقم الهوية:")

if search_input:
    if search_input.isdigit():
        results = teachers[teachers['هوية'].astype(str).str.contains(search_input)]
    else:
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
        with col2:
            if st.button("توليد Word", key="single_search_word"):
                doc = generate_doc(row)
                if doc:
                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)
                    with open(path, "rb") as f:
                        st.download_button("تحميل Word", f, file_name=f"{row['اسم']}.docx")
        with col3:
            if st.button("إلغاء التعيين", key="single_search_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")

    else:
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
        with col2:
            if st.button("توليد Word", key="multi_search_word"):
                doc = generate_doc(row)
                if doc:
                    os.makedirs("تكليفات", exist_ok=True)
                    path = f"تكليفات/{row['اسم']}.docx"
                    doc.save(path)
                    with open(path, "rb") as f:
                        st.download_button("تحميل Word", f, file_name=f"{row['اسم']}.docx")
        with col3:
            if st.button("إلغاء التعيين", key="multi_search_remove"):
                teachers.loc[teachers['هوية']==row['هوية'], ['قاعة مختارة','مهمة']] = [None,row['وظيفة']]
                save_assignments()
                st.warning(f"تم إلغاء التعيين للمعلم {row['اسم']}")

# ================== إدارة القاعات ==================
st.subheader("🏫 إدارة القاعات")
hall_selected = st.selectbox("اختر قاعة:", ["اختر القاعة..."] + list(halls['قاعة']))
if hall_selected != "اختر القاعة...":
    teachers_in_hall = teachers[teachers['قاعة مختارة']==hall_selected]
    if not teachers_in_hall.empty:
        st.write("المعلمون في القاعة:")
        st.table(teachers_in_hall[['اسم','مهمة','جوال']])
        # حذف معلم
        remove_teacher = st.selectbox("اختر معلم لإلغاء تكليفه:", teachers_in_hall['اسم'], key="remove_teacher_hall")
        if st.button("إلغاء تكليف هذا المعلم", key="remove_teacher_button"):
            teachers.loc[teachers['اسم']==remove_teacher, 'قاعة مختارة'] = None
            save_assignments()
            st.warning(f"تم إلغاء تكليف {remove_teacher} من {hall_selected}")

        # توليد كتب القاعة
        if st.button("توليد كتب التكليف لهذه القاعة", key="generate_hall"):
            os.makedirs("تكليفات", exist_ok=True)
            word_files = []
            for _, row in teachers_in_hall.iterrows():
                doc = generate_doc(row)
                if doc:
                    word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
                    doc.save(word_path)
                    word_files.append(word_path)
            zip_path = f"تكليفات/تكليفات_{hall_selected}.zip"
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for word_file in word_files:
                    zipf.write(word_file, os.path.basename(word_file))
            with open(zip_path,"rb") as f:
                st.download_button("تحميل ZIP للقاعة", f, file_name=f"تكليفات_{hall_selected}.zip", mime="application/zip")

# ================== توليد كل الكتب دفعة واحدة ==================
if st.button("توليد جميع كتب التكليف", key="generate_all"):
    os.makedirs("تكليفات", exist_ok=True)
    word_files = []
    for _, row in teachers.dropna(subset=['قاعة مختارة']).iterrows():
        doc = generate_doc(row)
        if doc:
            word_path = f"تكليفات/تكليف_{row['اسم']}.docx"
            doc.save(word_path)
            word_files.append(word_path)
    zip_path = "تكليفات/جميع_التكليفات.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for word_file in word_files:
            zipf.write(word_file, os.path.basename(word_file))
    with open(zip_path,"rb") as f:
        st.download_button("تحميل ZIP لجميع الكتب", f, file_name="جميع_التكليفات.zip", mime="application/zip")
