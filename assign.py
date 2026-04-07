import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
import io
import os
import re

# =====================================
# 1. إعدادات الواجهة والألوان
# =====================================
st.set_page_config(page_title="نظام التكليفات المتكامل 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    div[data-testid="stExpander"] { border: 1px solid #444 !important; background-color: #1a1c23 !important; }
    button[key^="btn_"] { background-color: #28a745 !important; color: white !important; }
    button[key^="del_"] { background-color: #dc3545 !important; color: white !important; }
    .stDownloadButton button { background-color: #007bff !important; color: white !important; }
    label, .stMarkdown p { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات
conn = sqlite3.connect("data_system_v17.db", check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls (hall_name TEXT PRIMARY KEY, city TEXT)''')
conn.commit()

# روابط جوجل شيت
TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

# =====================================
# 2. وظائف المعالجة
# =====================================

def generate_word_doc(row, hall_name, hall_city):
    """توليد ملف Word بناءً على القالب المرفوع"""
    try:
        # التأكد من وجود قالب
        template_source = "template.docx"
        if not os.path.exists(template_source):
            return None
            
        doc = Document(template_source)
        repls = {
            '<NAME>': str(row['name']),
            '<ID>': str(row['id']),
            '<JOB>': str(row['role']),
            '<HALL_NAME>': str(hall_name),
            '<HALL_LOCATION>': str(hall_city),
            '<WORKPLACE>': str(row['school']),
            '<CITY>': str(row['city'])
        }

        for p in doc.paragraphs:
            for k, v in repls.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text: run.text = run.text.replace(k, v); run.bold = True
        
        for table in doc.tables:
            for r in table.rows:
                for cell in r.cells:
                    for p in cell.paragraphs:
                        for k, v in repls.items():
                            if k in p.text:
                                for run in p.runs:
                                    if k in run.text: run.text = run.text.replace(k, v); run.bold = True
        
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except: return None

# =====================================
# 3. بناء التبويبات
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

# --- تبويب البحث والتعيين ---
with tab_search:
    st.subheader("تعيين الموظفين")
    df_h = pd.read_sql("SELECT * FROM halls", conn)
    hall_map = {r['hall_name']: r['city'] for _, r in df_h.iterrows()}
    
    q = st.text_input("ابحث عن اسم الموظف أو رقم الهوية:")
    if q:
        df_t = pd.read_sql("SELECT * FROM teachers", conn)
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        
        for _, row in results.iterrows():
            with st.expander(f"👤 {row['name']} | القاعة: {row['hall'] or 'غير معين'}"):
                c1, c2 = st.columns(2)
                with c1:
                    new_h = st.selectbox("اختر القاعة", [""] + list(hall_map.keys()), index=(list(hall_map.keys()).index(row['hall'])+1 if row['hall'] in hall_map else 0), key=f"sel_h_{row['id']}")
                    new_r = st.selectbox("الوظيفة", ["", "رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن"], index=0, key=f"sel_r_{row['id']}")
                with c2:
                    if st.button("💾 حفظ التكليف", key=f"save_{row['id']}"):
                        c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (new_h, new_r, hall_map.get(new_h, ""), row['id']))
                        conn.commit()
                        st.success("تم التحديث")
                        st.rerun()
                    
                    if row['hall'] and row['role']:
                        doc_file = generate_word_doc(row, row['hall'], row['hall_city'])
                        if doc_file:
                            st.download_button("📥 تحميل كتاب التكليف", data=doc_file, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# --- تبويب رفع الملفات ---
with tab_upload:
    st.subheader("إدارة ملفات النظام")
    
    # 1. رفع قالب الوورد
    uploaded_tpl = st.file_uploader("ارفع قالب التكليف (template.docx)", type="docx")
    if uploaded_tpl:
        with open("template.docx", "wb") as f:
            f.write(uploaded_tpl.getbuffer())
        st.success("✅ تم تحديث ملف القالب بنجاح")

    st.divider()
    
    # 2. التزامن مع جوجل
    if st.button("🔄 تحديث البيانات من Google Sheets"):
        try:
            # تحديث المعلمين
            dft = pd.read_csv(TEACHERS_URL)
            dft.columns = dft.columns.str.strip().str.lower()
            if 'id_number' in dft.columns: dft.rename(columns={'id_number': 'id'}, inplace=True)
            for col in ['role', 'hall', 'hall_city']: 
                if col not in dft.columns: dft[col] = ""
            dft.to_sql('teachers', conn, if_exists='replace', index=False)
            
            # تحديث القاعات
            dfh = pd.read_csv(HALLS_URL)
            dfh.columns = dfh.columns.str.strip().str.lower()
            dfh.to_sql('halls', conn, if_exists='replace', index=False)
            
            st.success("✅ تمت المزامنة بنجاح")
            st.rerun()
        except Exception as e: st.error(f"خطأ في المزامنة: {e}")

# --- تبويب الإدارة ---
with tab_manage:
    st.subheader("إدارة القاعات والمخرجات")
    
    df_act = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall != ''", conn)
    if not df_act.empty:
        selected_hall = st.selectbox("اختر قاعة لإدارتها:", [""] + sorted(df_act['hall'].tolist()))
        
        if selected_hall:
            df_members = pd.read_sql("SELECT * FROM teachers WHERE hall = ?", conn, params=(selected_hall,))
            
            # أزرار التحميل الجماعي
            col_a, col_b = st.columns(2)
            with col_a:
                # تصدير Excel
                output_exc = io.BytesIO()
                with pd.ExcelWriter(output_exc, engine='xlsxwriter') as writer:
                    df_members.to_excel(writer, index=False, sheet_name='المكلفين')
                st.download_button("📊 تحميل كشف العاملين (Excel)", data=output_exc.getvalue(), file_name=f"كشف_{selected_hall}.xlsx")
            
            with col_b:
                st.write(f"عدد الموظفين: {len(df_members)}")

            st.divider()
            
            # عرض وتعديل الموظفين داخل القاعة
            for _, m in df_members.iterrows():
                with st.expander(f"📝 {m['name']} - ({m['role']})"):
                    c1, c2, c3 = st.columns([2,2,1])
                    all_halls = list(hall_map.keys())
                    with c1:
                        edit_h = st.selectbox("نقل لقاعة أخرى", all_halls, index=all_halls.index(m['hall']), key=f"ed_h_{m['id']}")
                    with c2:
                        edit_r = st.selectbox("تغيير الوظيفة", ["رئيس قاعة", "مراقب", "آذن"], index=0, key=f"ed_r_{m['id']}")
                    with c3:
                        if st.button("تحديث", key=f"up_m_{m['id']}"):
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (edit_h, edit_r, hall_map.get(edit_h, ""), m['id']))
                            conn.commit(); st.rerun()
                        if st.button("حذف التكليف", key=f"rm_m_{m['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (m['id'],))
                            conn.commit(); st.rerun()

    st.divider()
    if st.button("⚠️ مسح شامل لجميع التكليفات"):
        c.execute("UPDATE teachers SET hall='', role='', hall_city=''")
        conn.commit(); st.success("تم تصفير النظام"); st.rerun()
