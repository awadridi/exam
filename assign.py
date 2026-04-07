import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os

# =====================================
# 1. إعدادات قاعدة البيانات
# =====================================
conn = sqlite3.connect("final_exam_2026.db", check_same_thread=False)
c = conn.cursor()

c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. دالة تعبئة نموذج Word الخاص بك
# =====================================
def generate_official_doc(row):
    # تحميل القالب الذي أرسلته (يجب تسميته template.docx)
    try:
        doc = Document("template.docx")
    except:
        # إذا لم يجد القالب، ينشئ ملفاً بسيطاً للتنبيه
        doc = Document()
        doc.add_paragraph("خطأ: لم يتم العثور على ملف template.docx")
        return doc

    # استبدال البيانات في كل فقرات المستند
    for paragraph in doc.paragraphs:
        if '<NAME>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<NAME>', str(row['name']))
        if '<ID>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<ID>', str(row['id']))
        if '<HALL_NAME>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<HALL_NAME>', str(row['hall']))
        if '<HALL_LOCATION>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<HALL_LOCATION>', str(row['hall_city']))
        if '<WORKPLACE>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<WORKPLACE>', str(row['school']))
        if '<CITY>' in paragraph.text:
            paragraph.text = paragraph.text.replace('<CITY>', str(row['city']))

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =====================================
# 3. الواجهة الرئيسية
# =====================================
st.set_page_config(page_title="نظام امتحانات 2026", layout="wide")
st.title("📑 إصدار تكليفات الثانوية العامة 2026")

tab_search, tab_upload = st.tabs(["🔍 بحث وإصدار", "📥 رفع البيانات"])

with tab_search:
    q = st.text_input("ابحث عن موظف (اسم أو هوية)")
    
    # جلب بيانات القاعات للربط
    df_h = pd.read_sql("SELECT * FROM halls", conn)
    h_map = {row['hall_name']: row['city'] for _, row in df_h.iterrows()}
    hall_names = [""] + list(h_map.keys())

    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']}"):
                col1, col2 = st.columns(2)
                
                # اختيار القاعة وتحديث مكانها تلقائياً
                selected_hall = col1.selectbox("تعيين القاعة", hall_names, 
                                             index=hall_names.index(row['hall']) if row['hall'] in hall_names else 0,
                                             key=f"h_{row['id']}")
                
                if col1.button("💾 حفظ التعيين", key=f"s_{row['id']}"):
                    h_city = h_map.get(selected_hall, "")
                    c.execute("UPDATE teachers SET hall=?, hall_city=? WHERE id=?", (selected_hall, h_city, row['id']))
                    conn.commit()
                    st.success("تم الحفظ")
                    st.rerun()

                if row['hall']:
                    doc_bytes = generate_official_doc(row)
                    col2.download_button("📥 تحميل الكتاب الرسمي", data=doc_bytes, 
                                       file_name=f"تكليف_{row['name']}.docx",
                                       key=f"dl_{row['id']}")

with tab_upload:
    st.subheader("رفع ملفات الإكسل")
    
    up_h = st.file_uploader("1. رفع ملف القاعات", type="xlsx")
    if up_h and st.button("تثبيت القاعات"):
        df = pd.read_excel(up_h)
        c.execute("DELETE FROM halls")
        for _, r in df.iterrows():
            # نفترض الأعمدة: رقم القاعة، اسم القاعة، المدينة
            c.execute("INSERT INTO halls VALUES (?,?,?)", (str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2])))
        conn.commit()
        st.success("تم رفع القاعات")

    up_t = st.file_uploader("2. رفع ملف الموظفين", type="xlsx")
    if up_t and st.button("تثبيت الموظفين"):
        df = pd.read_excel(up_t)
        c.execute("DELETE FROM teachers")
        for _, r in df.iterrows():
            # مطابقة أعمدة الصورة: id, name, school, city, phone, role...
            c.execute("INSERT INTO teachers (id, name, school, city, phone, role, hall, hall_city) VALUES (?,?,?,?,?,?,?,?)",
                      (str(r['id']), str(r['name']), str(r['school']), str(r['city']), str(r['phone']), str(r['role']), "", ""))
        conn.commit()
        st.success("تم رفع الموظفين")
