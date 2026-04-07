import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt
import io
import os
import re

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
        replacements = {
            '<NAME>': str(row['name']), '<ID>': str(row['id']),
            '<JOB>': str(row['role']), '<HALL_NAME>': str(row['hall']),
            '<HALL_LOCATION>': str(row['hall_city']), '<WORKPLACE>': str(row['school']),
            '<CITY>': str(row['city'])
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

with tab_upload:
    st.header("🔄 المزامنة مع Google Sheets")
    st.write("استخدم هذا الزر لسحب أحدث البيانات من ملف جوجل شيت الخاص بك.")
    
    if st.button("📥 سحب البيانات الجديدة", key="main_sync"):
        sync_data()

    st.divider() 
    st.subheader("📄 رفع نموذج كتاب التكليف")
    f_template = st.file_uploader("ارفع ملف الوورد (template.docx)", type="docx", key="u_docx")
    if f_template and st.button("تثبيت النموذج الجديد"):
        with open("template.docx", "wb") as f:
            f.write(f_template.getbuffer())
        st.success("✅ تم تحديث نموذج كتاب التكليف بنجاح!")

with tab_manage:
    st.header("⚙️ أدوات الإدارة والاستخراج")
    
    # 1. جلب البيانات من القاعدة
    df_all = pd.read_sql("SELECT * FROM teachers WHERE hall IS NOT NULL AND hall != ''", conn)
    
    st.subheader("📊 استخراج كشوفات القاعات")
    
    if not df_all.empty:
        # الحصول على قائمة القاعات
        halls_with_assignments = sorted(df_all['hall'].unique())
        
        # تعريف المتغير هنا أولاً قبل استخدامه
        selected_h_export = st.selectbox("اختر القاعة لتصدير كشف العاملين بها:", [""] + halls_with_assignments)
        
        # التأكد أن المستخدم اختار قاعة فعلاً
        if selected_h_export and selected_h_export != "":
            # تصفية البيانات
            df_hall_export = df_all[df_all['hall'] == selected_h_export][['id', 'name', 'role', 'school', 'city', 'phone']]
            df_hall_export.columns = ['الرقم الوطني/المنشأة', 'الاسم', 'المهمة', 'المدرسة الأصلية', 'السكن', 'رقم الهاتف']
            
            # إعداد ملف الإكسل المنسق
            # 1. إنشاء الـ Buffer والكاتب
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # كتابة البيانات
                df_hall_export.to_excel(writer, index=False, sheet_name='العاملين')
                
                # تعريف الكائنات
                workbook  = writer.book
                worksheet = writer.sheets['العاملين']

                # 2. إعدادات الصفحة (أفقي، يمين لليسار، احتواء صفحة)
                worksheet.right_to_left()
                worksheet.set_landscape()
                worksheet.fit_to_pages(1, 0)

                # 3. تعريف التنسيقات (الخطوط والحدود)
                header_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#D7E4BC', 'border': 1, 'font_size': 16
                })

                cell_format = workbook.add_format({
                    'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_size': 14
                })

                # 4. تطبيق التنسيق على الأعمدة
                for col_num, value in enumerate(df_hall_export.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 25, cell_format)
                
                worksheet.set_column(1, 1, 45, cell_format) # عمود الاسم أعرض
                # تنسيق العناوين (كبرنا الخط لـ 16)
                header_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'fg_color': '#D7E4BC', 'border': 1, 'font_size': 16
                })

                # تنسيق الخلايا (كبرنا الخط لـ 14 وخلينا المحاذاة يمين)
                cell_format = workbook.add_format({
                    'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_size': 14
                })

                # تطبيق التنسيق وتوسيع الأعمدة أكثر لتناسب الخط الكبير
                for col_num, value in enumerate(df_hall_export.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 25, cell_format) # زدنا العرض لـ 25
                
                worksheet.set_column(1, 1, 45, cell_format) # الاسم خليناه عريض جداً (45)

                for col_num, value in enumerate(df_hall_export.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 20, cell_format)
                
                worksheet.set_column(1, 1, 35, cell_format) # توسيع عمود الاسم

            st.download_button(
                label=f"📥 تحميل كشف {selected_h_export} منسق (Excel)",
                data=buffer.getvalue(),
                file_name=f"كشف_{selected_h_export}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("ℹ️ لا توجد تكليفات حالياً لتصديرها.")

    st.divider()
    st.subheader("🗑️ إدارة وحذف تكليفات القاعات")
    
    # 1. جلب قائمة القاعات التي بها تكليفات حالياً
    df_active_halls = pd.read_sql("SELECT DISTINCT hall FROM teachers WHERE hall IS NOT NULL AND hall != ''", conn)
    
    if not df_active_halls.empty:
        hall_to_manage = st.selectbox("اختر قاعة لعرض المكلفين بها وحذفهم:", [""] + sorted(df_active_halls['hall'].tolist()))
        
        if hall_to_manage:
            # جلب الموظفين في هذه القاعة فقط
            df_members = pd.read_sql("SELECT id, name, role FROM teachers WHERE hall = ?", conn, params=(hall_to_manage,))
            
            if not df_members.empty:
                st.write(f"👥 عدد المكلفين في {hall_to_manage}: **{len(df_members)}**")
                
                # عرض جدول مع زر حذف لكل موظف
                for index, row in df_members.iterrows():
                    col1, col2, col3 = st.columns([3, 2, 1])
                    with col1:
                        st.write(f"👤 {row['name']}")
                    with col2:
                        st.write(f"🏷️ {row['role']}")
                    with col3:
                        # زر الحذف لكل موظف بشكل مستقل
                        if st.button("حذف", key=f"del_{row['id']}"):
                            c.execute("UPDATE teachers SET hall='', role='', hall_city='' WHERE id=?", (row['id'],))
                            conn.commit()
                            st.success(f"تم حذف تكليف {row['name']}")
                            st.rerun() # لإعادة تحديث القائمة فوراً
            else:
                st.info("لا يوجد موظفون مكلفون في هذه القاعة حالياً.")
    else:
        st.info("لا توجد قاعات تحتوي على تكليفات حالياً.")

    st.divider()
    st.subheader("⚠️ منطقة الخطر")
    if st.button("⚠️ مسح شامل للتكليفات"):
        c.execute("UPDATE teachers SET hall='', role='', hall_city=''")
        conn.commit()
        st.success("تم مسح التكليفات بنجاح.")
        st.rerun()
