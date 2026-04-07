import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import re

# =====================================
# 1. إعدادات الواجهة
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; width: 100% !important; }
    label, .stMarkdown p { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# الاتصال بقاعدة البيانات
conn = sqlite3.connect("data_system_v18.db", check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
conn.commit()

# روابط جوجل (للمزامنة)
TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 2. وظائف معالجة الملفات (Word & Excel)
# =====================================

def process_doc(doc_obj, row, h_name, h_city):
    """دالة استبدال النصوص في القالب مع الحفاظ على التنسيق"""
    repls = {
        '<NAME>': str(row.get('name', '')),
        '<ID>': str(row.get('id', '')),
        '<JOB>': str(row.get('role', '')),
        '<HALL_NAME>': str(h_name),
        '<HALL_LOCATION>': str(h_city),
        '<WORKPLACE>': str(row.get('school', '')),
        '<CITY>': str(row.get('city', ''))
    }
    for p in doc_obj.paragraphs:
        for k, v in repls.items():
            if k in p.text:
                for run in p.runs:
                    if k in run.text: run.text = run.text.replace(k, v); run.bold = True
    for table in doc_obj.tables:
        for r in table.rows:
            for cell in r.cells:
                for p in cell.paragraphs:
                    for k, v in repls.items():
                        if k in p.text:
                            for run in p.runs:
                                if k in run.text: run.text = run.text.replace(k, v); run.bold = True
    return doc_obj

def generate_single_doc(row):
    """توليد ملف واحد"""
    if not os.path.exists("template.docx"): return None
    doc = Document("template.docx")
    doc = process_doc(doc, row, row['hall'], row['hall_city'])
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_bulk_word(df, h_name):
    """توليد ملف Word واحد يحتوي على كل التكليفات"""
    if not os.path.exists("template.docx"): return None
    final_doc = Document("template.docx")
    final_doc._body.clear_content()
    for idx, row in df.iterrows():
        temp_doc = Document("template.docx")
        temp_doc = process_doc(temp_doc, row, h_name, row['hall_city'])
        if idx > 0: final_doc.add_page_break()
        for element in temp_doc.element.body:
            if not element.tag.endswith('sectPr'):
                final_doc.element.body.append(element)
    out = io.BytesIO()
    final_doc.save(out)
    out.seek(0)
    return out

# =====================================
# 3. واجهة المستخدم
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

# --- تبويب البحث والتعيين ---
with tab_search:
    st.subheader("إدارة الموظفين")
    df_h_data = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h_data.iterrows()}
    
    q = st.text_input("ابحث عن الاسم أو الهوية")
    if q:
        df_teachers = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_teachers[df_teachers['name'].str.contains(q, na=False) | df_teachers['id'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة الحالية: {row['hall'] or 'غير مكلف'}"):
                c1, c2 = st.columns(2)
                with c1:
                    sel_h = st.selectbox("اختر القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"q_h_{row['id']}")
                    sel_r = st.selectbox("المهمة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"q_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"btn_save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_h, sel_r, hall_map.get(sel_h, ""), row['id']))
                        conn.commit(); st.success("تم الحفظ"); st.rerun()
                    
                    if row['hall']:
                        if st.button("❌ حذف التكليف", key=f"del_single_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit(); st.rerun()
                        
                        f_word = generate_single_doc(row)
                        if f_word:
                            st.download_button("📥 تحميل كتاب التكليف", data=f_word, file_name=f"تكليف_{row['name']}.docx", key=f"dl_s_{row['id']}")

# --- تبويب رفع الملفات ---
with tab_upload:
    st.subheader("إعدادات القالب والبيانات")
    up_tpl = st.file_uploader("ارفع قالب الوورد (template.docx)", type="docx")
    if up_tpl:
        with open("template.docx", "wb") as f: f.write(up_tpl.getbuffer())
        st.success("تم تحديث القالب")

    if st.button("🔄 تحديث الأسماء من Google Sheets"):
        try:
            dft = pd.read_csv(TEACHERS_URL); dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            for col in ['role', 'hall', 'hall_city']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            dfh = pd.read_csv(HALLS_URL); dfh.columns = dfh.columns.str.strip().str.lower()
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            st.success("تم التحديث بنجاح"); st.rerun()
        except Exception as e: st.error(f"خطأ: {e}")

# --- تبويب الإدارة ---
with tab_manage:
    st.subheader("إدارة القاعات والمخرجات الجماعية")
    df_active = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    
    if not df_active.empty:
        h_choice = st.selectbox("اختر قاعة للعرض والإدارة:", [""] + sorted(df_active['hall'].tolist()))
        if h_choice:
            df_m = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(h_choice,))
            
            # --- منطقة أزرار التحميل الجماعي ---
            c_exc, c_wrd = st.columns(2)
            with c_exc:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_m.to_excel(writer, index=False, sheet_name='العاملين')
                st.download_button(f"📊 تحميل كشف {h_choice} (Excel)", data=output.getvalue(), file_name=f"كشف_{h_choice}.xlsx")
            
            with c_wrd:
                bulk_f = generate_bulk_word(df_m, h_choice)
                if bulk_f:
                    st.download_button(f"📄 تحميل كتب تكليف {h_choice} (Word)", data=bulk_f, file_name=f"تكليفات_{h_choice}.docx")

            st.write(f"عدد الموظفين في القاعة: **{len(df_m)}**")
            st.divider()

            # --- عرض الموظفين مع إمكانية التعديل أو الحذف ---
            for _, m in df_m.iterrows():
                with st.expander(f"📝 {m['name']} | الوظيفة: {m['role']}"):
                    col1, col2, col3 = st.columns([2, 2, 1])
                    with col1:
                        new_h = st.selectbox("نقل لقاعة", list(hall_map.keys()), index=list(hall_map.keys()).index(m['hall']), key=f"m_h_{m['id']}")
                    with col2:
                        new_r = st.selectbox("تغيير مهمة", ["رئيس قاعة", "مراقب", "آذن"], index=0, key=f"m_r_{m['id']}")
                    with col3:
                        if st.button("تحديث", key=f"m_up_{m['id']}"):
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (new_h, new_r, hall_map.get(new_h, ""), m['id']))
                            conn.commit(); st.rerun()
                        if st.button("حذف", key=f"m_del_{m['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (m['id'],))
                            conn.commit(); st.rerun()

    st.divider()
    if st.button("⚠️ مسح شامل لجميع التكليفات", type="secondary"):
        c.execute("UPDATE teachers SET hall='', role='', hall_city=''")
        conn.commit(); st.success("تم تصفير النظام"); st.rerun()
