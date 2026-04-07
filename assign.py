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
