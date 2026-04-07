import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
import io
import os
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# =====================================
# 1. إعدادات الواجهة وتثبيت الألوان (CSS)
# =====================================
st.set_page_config(page_title="نظام التكليفات 2026", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; background-color: #0e1117; }
    
    /* تنسيق مستطيل الاسم (Expander) */
    div[data-testid="stExpander"] {
        border: 1px solid #444 !important;
        background-color: #1a1c23 !important;
    }
    div[data-testid="stExpander"] summary p {
        color: #ffffff !important;
        font-weight: bold !important;
    }

    /* ألوان الأزرار إجبارية */
    /* حفظ - أخضر */
    button[key^="btn_"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
    }
    /* إلغاء - أحمر */
    button[key^="del_"] {
        background-color: #dc3545 !important;
        color: white !important;
        border: none !important;
    }
    /* تحميل - أزرق */
    .stDownloadButton button {
        background-color: #007bff !important;
        color: white !important;
        width: 100% !important;
    }
    
    label, .stMarkdown p { color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# قاعدة البيانات - محاولة جعلها ثابتة قدر الإمكان
db_path = "data_system_v16.db"
conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS teachers 
             (id TEXT PRIMARY KEY, name TEXT, school TEXT, city TEXT, phone TEXT, role TEXT, hall TEXT, hall_city TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS halls 
             (number TEXT PRIMARY KEY, hall_name TEXT, city TEXT)''')
conn.commit()

# =====================================
# 2. وظيفة الاستبدال (البولد الذكي)
# =====================================
def generate_from_template(row):
    try:
        doc = Document("template.docx")
        # قائمة الاستبدالات مع التأكد من جلب الوظيفة بشكل صحيح
        replacements = {
            '<NAME>': str(row.get('name', '')),
            '<ID>': str(row.get('id_number', '')),
            
            # يحاول جلب 'role' وإذا لم يجدها يبحث عن 'job'
            '<JOB>': str(row.get('role', row.get('job', ''))), 
            
            '<HALL_NAME>': str(hall_name),
            '<HALL_LOCATION>': str(row.get('hall_city', '')),
            '<WORKPLACE>': str(row.get('school', '')),
            '<CITY>': str(row.get('city', ''))
                        }
        def apply_smart_bold_replace(paragraph, data_map):
            text = paragraph.text
            if any(key in text for key in data_map):
                # 1. حفظ حجم الخط الأصلي من أول جزء في الفقرة قبل مسحه
                original_size = Pt(14) # القيمة الافتراضية إذا فشل البرنامج في القراءة
                if paragraph.runs and paragraph.runs[0].font.size:
                    original_size = paragraph.runs[0].font.size
                
                # 2. تفريغ الـ runs الحالية
                for run in paragraph.runs: run.text = ""
                
                # 3. تقسيم النص وإعادة البناء بنفس الحجم الأصلي
                parts = re.split(r'(<[^>]+>)', text)
                for part in parts:
                    run = paragraph.add_run()
                    if part in data_map:
                        run.text = str(data_map[part]) if data_map[part] else ""
                        run.bold = True
                    else:
                        run.text = part
                    
                    # 4. إعادة تطبيق الحجم الأصلي المحفوظ على كل أجزاء الفقرة
                    run.font.size = original_size

        for p in doc.paragraphs: apply_smart_bold_replace(p, replacements)
        for table in doc.tables:
            for r in table.rows:
                for cell in r.cells:
                    for p in cell.paragraphs: apply_smart_bold_replace(p, replacements)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    except: return None

# =====================================
# 3. الواجهة الرئيسية
# =====================================
tab_search, tab_upload, tab_manage = st.tabs(["🔍 البحث والتعيين", "📥 رفع الملفات", "⚙️ الإدارة"])

with tab_search:
    st.subheader("إدارة الموظفين")
    try:
        df_halls = pd.read_sql("SELECT * FROM halls", conn)
        df_halls.columns = df_halls.columns.str.strip().str.lower()
        if 'hall_name' in df_halls.columns:
            hall_map = {str(r['hall_name']): str(r['city']) for _, r in df_halls.iterrows()}
        else:
            hall_map = {}
    except:
        hall_map = {}
    hall_list = [""] + list(hall_map.keys())
    role_list = ["", "رئيس قاعة", "مراقب", "مساعد رئيس قاعة", "آذن", "عضو لجنة"]

    q = st.text_input("ابحث عن الاسم أو الهوية")
    df_t = pd.read_sql("SELECT * FROM teachers", conn)
    
    if q and not df_t.empty:
        results = df_t[df_t['name'].str.contains(q, na=False) | df_t['id'].astype(str).str.contains(q)]
        for i, row in results.iterrows():
            title = f"👤 {row['name']} | {row['role'] if row['role'] else '-'} | {row['hall'] if row['hall'] else '-'}"
            with st.expander(title):
                c1, c2 = st.columns(2)
                with c1:
                    sel_hall = st.selectbox(f"اختر القاعة لـ {row['id']}", hall_list, index=hall_list.index(row['hall']) if row['hall'] in hall_list else 0, key=f"h_{row['id']}")
                    sel_role = st.selectbox(f"اختر الوظيفة لـ {row['id']}", role_list, index=role_list.index(row['role']) if row['role'] in role_list else 0, key=f"r_{row['id']}")
                with c2:
                    st.write(f"المدرسة: {row['school']}")
                    col_save, col_del = st.columns(2)
                    with col_save:
                        if st.button("💾 حفظ البيانات", key=f"btn_{row['id']}"):
                            h_city = hall_map.get(sel_hall, "")
                            c.execute("UPDATE teachers SET hall=?, role=?, hall_city=? WHERE id=?", (sel_hall, sel_role, h_city, row['id']))
                            conn.commit(); st.success("تم الحفظ!"); st.rerun()
                    with col_del:
                        if row['role'] or row['hall']:
                            if st.button("❌ إلغاء التكليف", key=f"del_{row['id']}"):
                                c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                                conn.commit(); st.warning("تم الإلغاء"); st.rerun()
                    
                    if row['hall'] and row['role']:
                        f_data = generate_from_template(row)
                        if f_data:
                            st.download_button("📥 تحميل التكليف", data=f_data, file_name=f"تكليف_{row['name']}.docx", key=f"dl_{row['id']}")

# تبويب الرفع (تم إرجاع زر القاعات)
# =====================================
# 3. الواجهة الرئيسية والمزامنة مع Google Sheets
# =====================================

# --- إعدادات جوجل شيت (تأكد من وضع القيم الصحيحة هنا) ---
# الروابط الناتجة من "النشر على الويب" (Publish to web)

TEACHERS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=264504938&single=true&output=csv"
HALLS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSubFlcocaWSvF7GU14hNGx1cuLJBwF5SchDxzeaNMJnSy6T_b0Hu5aDMnc-OM9u7EnNIATUui12H9L/pub?gid=1364805271&single=true&output=csv"

def sync_data():
    try:
        # 1. سحب بيانات المعلمين وتنظيف أسماء الأعمدة
        df_t = pd.read_csv(TEACHERS_URL)
        df_t.columns = df_t.columns.str.strip().str.lower()
        
        # 2. إضافة أعمدة التكليف (hall, role, hall_city) إذا لم تكن موجودة في ملف جوجل
        # هذه الخطوة تحل مشكلة sqlite3.OperationalError عند الحفظ
        for col in ['hall', 'role', 'hall_city']:
            if col not in df_t.columns:
                df_t[col] = None
        
        # حفظ بيانات المعلمين في قاعدة البيانات
        df_t.to_sql('teachers', conn, if_exists='replace', index=False)
        
        # 3. سحب بيانات القاعات وتنظيفها
        df_h = pd.read_csv(HALLS_URL)
        df_h.columns = df_h.columns.str.strip().str.lower()
        df_h.to_sql('halls', conn, if_exists='replace', index=False)
        
        st.success("✅ تم التحديث من جوجل وإعداد الجداول بنجاح!")
        st.rerun()
    except Exception as e:
        st.error(f"❌ خطأ أثناء المزامنة: {e}")

    # أضف 4 مسافات في بداية الأسطر التالية:
    if st.button("🔄 تحديث الأسماء الآن"): 
            sync_data()

    st.subheader("إدارة الموظفين")
    # ... وبقية الكود الذي يليه
    # ... باقي الكود
    # ... (باقي كود البحث والتعيين كما هو عندك دون تغيير) ...
# 1. جلب قائمة القاعات التي بها تكليفات حالياً
    df_active_halls = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall IS NOT NULL AND hall != ''", conn)
    
    if not df_active_halls.empty:
        hall_to_manage = st.selectbox("اختر قاعة لعرض المكلفين بها وحذفهم:", [""] + sorted(df_active_halls['hall'].tolist()))
        
        if hall_to_manage and hall_to_manage != "":
            # جلب الموظفين في هذه القاعة
            df_members = pd.read_sql("SELECT id, name, role, school, city, hall_city FROM teachers WHERE hall = ?", conn, params=(hall_to_manage,))
            
            if not df_members.empty:
                st.write(f"👥 عدد المكلفين في {hall_to_manage}: **{len(df_members)}**")
                
                # جلب قائمة القاعات للنقل
                all_halls_df = pd.read_sql("SELECT DISTINCT hall_name FROM halls", conn)
                halls_list = sorted(all_halls_df['hall_name'].tolist()) if not all_halls_df.empty else [hall_to_manage]

                # --- حلقة عرض الموظفين (فقط للعرض والحذف الفردي) ---
                for index, row in df_members.iterrows():
                    with st.expander(f"👤 {row['name']} - {row['role']}"):
                        c1, c2, c3 = st.columns([2, 2, 1])
                        with c1:
                            roles_list = ["رئيس قاعة", "مساعد رئيس قاعة", "مراقب", "آذن", "عضو لجنة"]
                            curr_role_idx = roles_list.index(row['role']) if row['role'] in roles_list else 2
                            new_role = st.selectbox(f"تغيير المهمة", roles_list, index=curr_role_idx, key=f"r_edit_{row['id']}")
                        with c2:
                            curr_hall_idx = halls_list.index(hall_to_manage) if hall_to_manage in halls_list else 0
                            new_hall = st.selectbox(f"نقل إلى قاعة", halls_list, index=curr_hall_idx, key=f"h_edit_{row['id']}")
                        with c3:
                            st.write("") 
                            if st.button("✅ تحديث", key=f"upd_btn_{row['id']}"):
                                c.execute("UPDATE teachers SET hall=?, role=? WHERE id=?", (new_hall, new_role, row['id']))
                                conn.commit()
                                st.rerun()
                            if st.button("🗑️ حذف", key=f"del_btn_{row['id']}"):
                                c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                                conn.commit()
                                st.rerun()

                # --- الآن: تعريف دالة الطباعة الجماعية (خارج حلقة الـ for) ---
                def generate_bulk_docs(df, h_name):
                    from docx import Document
                    import io
                    template_p = "template.docx"
                    if not os.path.exists(template_p): return None
                    
                    final_doc = Document(template_p)
                    # مسح محتوى الصفحة الأولى من الملف الأساسي لبدء الدمج نظيفاً
                    final_doc._body.clear_content()
                    
                    for idx, r in df.iterrows():
                        temp_doc = Document(template_p)
                        # حل مشكلة الـ JOB: نبحث في كل الاحتمالات
                        job_val = r.get('role') or r.get('job') or ""
                        
                        repls = {
                            '<NAME>': str(r.get('name', '')),
                            '<ID>': str(r.get('id', r.get('id_number', ''))),
                            '<JOB>': str(job_val),
                            '<HALL_NAME>': str(h_name),
                            '<HALL_LOCATION>': str(r.get('hall_city', '')),
                            '<WORKPLACE>': str(r.get('school', '')),
                            '<CITY>': str(r.get('city', ''))
                        }

                        def smart_rep(doc_obj):
                            for p in doc_obj.paragraphs:
                                for k, v in repls.items():
                                    if k in p.text:
                                        for run in p.runs:
                                            if k in run.text: run.text = run.text.replace(k, v)
                            for table in doc_obj.tables:
                                for row_t in table.rows:
                                    for cell in row_t.cells:
                                        for p in cell.paragraphs:
                                            for k, v in repls.items():
                                                if k in p.text:
                                                    for run in p.runs:
                                                        if k in run.text: run.text = run.text.replace(k, v)

                        smart_rep(temp_doc)
                        
                        # دمج الصفحات
                        if idx > 0: final_doc.add_page_break()
                        for element in temp_doc.element.body:
                            if not element.tag.endswith('sectPr'):
                                final_doc.element.body.append(element)
                    
                    out = io.BytesIO()
                    final_doc.save(out)
                    out.seek(0)
                    return out

                # --- الزر النهائي الوحيد لإصدار الكل ---
                st.markdown("### 📄 إصدار جميع التكليفات")
                try:
                    with st.spinner("جاري تجهيز الملف..."):
                        final_word = generate_bulk_docs(df_members, hall_to_manage)
                        if final_word:
                            st.download_button(
                                label=f"📥 تحميل جميع تكليفات قاعة {hall_to_manage} (ملف واحد)",
                                data=final_word,
                                file_name=f"تكليفات_{hall_to_manage}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"bulk_dl_{hall_to_manage}",
                                use_container_width=True
                            )
                except Exception as e:
                    st.error(f"حدث خطأ أثناء الإصدار: {e}")
            else:
                st.info("لا يوجد موظفون في هذه القاعة.")
    
    st.divider()
    st.subheader("⚠️ منطقة الخطر")
    if st.button("⚠️ مسح شامل للتكليفات"):
        c.execute("UPDATE teachers SET hall='', role='', hall_city=''")
        conn.commit()
        st.success("تم مسح التكليفات بنجاح.")
        st.rerun()
