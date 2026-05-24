import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="نظام إدارة عقارات المحلات المطور", layout="wide")

# ==================== دالة معالجة الإكسل لفتحها بأعمدة منفصلة ====================
def convert_df_to_excel_csv(df):
    csv_string = "sep=,\n" + df.to_csv(index=False)
    return csv_string.encode('utf-8-sig')

# ==================== دالة توليد تقرير ذكي للطباعة كـ PDF ====================
def convert_df_to_pdf_html(df, title="تقرير النظام"):
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; direction: rtl; text-align: right; padding: 20px; background-color: #f5f5f5; }}
            .report-card {{ background-color: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); max-width: 900px; margin: auto; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; direction: rtl; }}
            th, td {{ border: 1px solid #ddd; padding: 12px; text-align: center; }}
            th {{ background-color: #2E86C1; color: white; font-size: 14px; }}
            td {{ font-size: 13px; color: #333; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            h2 {{ color: #2E86C1; text-align: center; margin-bottom: 5px; }}
            .date {{ text-align: left; font-size: 12px; color: #777; }}
            @media print {{
                .no-print {{ display: none !important; }}
                body {{ background-color: white; padding: 0; }}
                .report-card {{ box-shadow: none; padding: 0; max-width: 100%; }}
            }}
        </style>
    </head>
    <body>
        <div class="report-card">
            <h2>🏢 {title}</h2>
            <p class="date">تاريخ إصدار التقرير: {datetime.now().strftime('%Y-%m-%d | %I:%M %p')}</p>
            <hr style='border: 1px solid #eee;'>
            <table>
                <thead>
                    <tr>{"".join(f"<th>{col}</th>" for col in df.columns)}</tr>
                </thead>
                <tbody>
                    {"".join(f"<tr>{''.join(f'<td>{str(val)}</td>' for val in row)}</tr>" for row in df.values)}
                </tbody>
            </table>
            <br><br>
            <button class="no-print" onclick="window.print()" style="padding: 14px; background-color: #2E86C1; color: white; border: none; border-radius: 6px; cursor: pointer; width: 100%; font-size: 16px; font-weight: bold;">📸 اضغط هنا لحفظ التقرير كـ PDF أو طباعته</button>
        </div>
    </body>
    </html>
    """
    return html.encode('utf-8')

# ==================== دالة توليد سند قبض منفرد جاهز للطباعة الفورية ====================
def convert_receipt_to_pdf_html(receipt_data):
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; direction: rtl; text-align: right; padding: 40px; background-color: white; }}
            .receipt-card {{ border: 2px dashed #4CAF50; padding: 30px; border-radius: 10px; max-width: 550px; margin: auto; background-color: #f9f9f9; color: #333; }}
            h2 {{ text-align: center; color: #2E86C1; margin-bottom: 0; }}
            h4 {{ text-align: center; color: #555; margin-top: 5px; }}
            hr {{ border: 1px solid #ddd; }}
            p {{ font-size: 16px; line-height: 1.8; }}
            @media print {{
                .no-print {{ display: none !important; }}
                body {{ background-color: white; padding: 0; }}
                .receipt-card {{ border: 2px dashed #4CAF50; box-shadow: none; max-width: 100%; }}
            }}
        </style>
    </head>
    <body>
        <div class="receipt-card">
            <h2>🧾 سند قبض مالي رسمي</h2>
            <h4>رقم السند الموحد: {receipt_data['رقم السند']}</h4>
            <hr>
            <p><strong>التاريخ م:</strong> {receipt_data['التاريخ']} م</p>
            <p><strong>وصلنا من السيد/ة:</strong> {receipt_data['المستأجر']} ( المستأجر لـ {receipt_data['رقم المحل']} )</p>
            <p><strong>مبلغ وقدره:</strong> <b style='color:#2E86C1; font-size:18px;'>{receipt_data['المبلغ']:,.2f} ريال سعودي</b></p>
            <p><strong>وذلك كدفعة عن طريق:</strong> {receipt_data['طريقة الدفع']}</p>
            <br>
            <p style='text-align: left; font-weight:bold;'>توقيع المسؤول المالي والمحصل: .....................</p>
            <br><br>
            <button class="no-print" onclick="window.print()" style="padding: 14px; background-color: #2E86C1; color: white; border: none; border-radius: 6px; cursor: pointer; width: 100%; font-size: 16px; font-weight: bold;">🖨️ اضغط هنا لطباعة السند فوراً أو حفظه كـ PDF</button>
        </div>
    </body>
    </html>
    """
    return html.encode('utf-8')

# ==================== تهيئة قواعد البيانات المؤقتة ====================
if 'shops_db' not in st.session_state:
    data = []
    for i in range(1, 167):
        data.append({
            "رقم المحل": f"محل {i}", "المساحة": 60, "الحالة": "شاغر", 
            "المستأجر": "-", "الإيجار السنوي": 15000, "بداية العقد": "-", "نهاية العقد": "-", "المحصل": 0
        })
    st.session_state.shops_db = pd.DataFrame(data)

if 'transactions_db' not in st.session_state:
    st.session_state.transactions_db = pd.DataFrame(columns=["التاريخ", "رقم السند", "رقم المحل", "المستأجر", "طريقة الدفع", "المبلغ"])

if 'historical_debts_db' not in st.session_state:
    st.session_state.historical_debts_db = pd.DataFrame(columns=["السنة المالية", "المستأجر السابق", "تفاصيل العقد", "المبلغ المتبقي"])

if 'expenses_db' not in st.session_state:
    st.session_state.expenses_db = pd.DataFrame(columns=["التاريخ", "بند الصرف", "المبلغ", "ملاحظات"])

# تهيئة ذاكرة السندات النشطة لمنع الاختفاء عند الطباعة
if 'latest_receipt' not in st.session_state:
    st.session_state.latest_receipt = None

st.title("🏢 نظام إدارة وتحصيل عقارات المحلات التجارية")
st.markdown("---")

# ==================== القائمة الرئيسية ====================
menu = st.sidebar.radio("قائمة النظام الأساسية", ["عمليات التحصيل وإدخال البيانات", "لوحة المؤشرات والتحليلات"])

if menu == "عمليات التحصيل وإدخال البيانات":
    tab1, tab2, tab3, tab4 = st.tabs(["📝 إدارة العقود والمحلات", "💰 التحصيل وسندات القبض", "📂 أرشيف ديون المغادرين", "🛠️ إدارة المصروفات"])
    
    # 1. إدارة العقود والمحلات
    with tab1:
        st.subheader("إدارة بيانات عقود الـ 166 محل")
        action_type = st.radio("اختر الإجراء المطلوب المُراد تنفيذه:", ["✍️ تسجيل عقد لمحل جديد (إدخال جديد)", "🔄 تعديل بيانات عقد قائم (تحديث)"], horizontal=True)
        
        if action_type == "✍️ تسجيل عقد لمحل جديد (إدخال جديد)":
            available_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] != "مؤجر"]["رقم المحل"].tolist()
            if available_shops:
                with st.form("new_contract_form"):
                    st.info("ملاحظة: هذه القائمة تعرض المحلات الشاغرة فقط. بمجرد الحفظ ستختفي من هنا.")
                    selected_shop = st.selectbox("اختر رقم المحل المُراد تأجيره:", available_shops)
                    col1, col2 = st.columns(2)
                    with col1:
                        tenant = st.text_input("اسم المستأجر الجديد:")
                        rent = st.number_input("الإيجار السنوي المتفق عليه:", min_value=0, value=15000)
                    with col2:
                        start_date = st.date_input("تاريخ بداية العقد جديد:")
                        end_date = st.date_input("تاريخ نهاية العقد جديد:")
                    
                    if st.form_submit_button("💾 حفظ العقد الجديد"):
                        if tenant.strip() == "" or tenant == "-":
                            st.error("الرجاء إدخال اسم مستأجر صحيح.")
                        else:
                            idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].index[0]
                            st.session_state.shops_db.at[idx, "الحالة"] = "مؤجر"
                            st.session_state.shops_db.at[idx, "المستأجر"] = tenant
                            st.session_state.shops_db.at[idx, "الإيجار السنوي"] = rent
                            st.session_state.shops_db.at[idx, "بداية العقد"] = start_date.strftime("%Y-%m-%d")
                            st.session_state.shops_db.at[idx, "نهاية العقد"] = end_date.strftime("%Y-%m-%d")
                            st.success(f"تم حفظ العقد بنجاح للمحل ({selected_shop}) وتم نقله لقائمة المؤجرة!")
                            st.rerun()
            else:
                st.success("🎉 جميع الـ 166 محل مؤجرة بالكامل حالياً!")
                
        else:
            rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
            if rented_shops:
                selected_shop = st.selectbox("اختر رقم المحل المؤجر المُراد تعديل بياناته:", rented_shops)
                current_data = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].iloc[0]
                
                with st.form("edit_contract_form"):
                    st.warning(f"أنت الآن تقوم بتعديل بيانات العقد للمحل: {selected_shop}")
                    col1, col2 = st.columns(2)
                    with col1:
                        edit_status = st.selectbox("تحديث الحالة (مثلاً في حال الإخلاء أو الصيانة):", ["مؤجر", "شاغر", "تحت الصيانة"])
                        edit_tenant = st.text_input("اسم المستأجر الحالي:", value=current_data["المستأجر"])
                        edit_rent = st.number_input("الإيجار السنوي الحالي:", min_value=0, value=int(current_data["الإيجار السنوي"]))
                    with col2:
                        try:
                            d1 = datetime.strptime(current_data["بداية العقد"], "%Y-%m-%d")
                            d2 = datetime.strptime(current_data["نهاية العقد"], "%Y-%m-%d")
                        except:
                            d1 = datetime.now()
                            d2 = datetime.now()
                        edit_start = st.date_input("تعديل تاريخ بداية العقد:", value=d1)
                        edit_end = st.date_input("تعديل تاريخ نهاية العقد:", value=d2)
                    
                    if st.form_submit_button("🔄 تحديث بيانات العقد"):
                        idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].index[0]
                        st.session_state.shops_db.at[idx, "الحالة"] = edit_status
                        st.session_state.shops_db.at[idx, "المستأجر"] = edit_tenant if edit_status == "مؤجر" else "-"
                        st.session_state.shops_db.at[idx, "الإيجار السنوي"] = edit_rent
                        st.session_state.shops_db.at[idx, "بداية العقد"] = edit_start.strftime("%Y-%m-%d")
                        st.session_state.shops_db.at[idx, "نهاية العقد"] = edit_end.strftime("%Y-%m-%d")
                        st.success(f"تم تحديث تعديلات المحل ({selected_shop}) بنجاح تام!")
                        st.rerun()
            else:
                st.info("لا توجد أي محلات مؤجرة حالياً لتعديلها.")
        
        st.markdown("---")
        st.subheader("📋 النظرة العامة: المحلات المؤجرة تفصيلياً")
        rented_display_df = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"][["رقم المحل", "المستأجر", "الإيجار السنوي", "بداية العقد", "نهاية العقد", "المحصل"]]
        if not rented_display_df.empty:
            st.dataframe(rented_display_df, use_container_width=True)
            ec1, ec2 = st.columns(2)
            with ec1:
                excel_data = convert_df_to_excel_csv(rented_display_df)
                st.download_button(label="📥 تحميل النظرة العامة (ملف Excel بأعمدة منفصلة)", data=excel_data, file_name='📊_المحلات_المؤجرة.csv', mime='text/csv')
            with ec2:
                pdf_html_data = convert_df_to_pdf_html(rented_display_df, "تقرير النظرة العامة للمحلات المؤجرة")
                st.download_button(label="📄 تصدير النظرة العامة كتقرير (PDF)", data=pdf_html_data, file_name='📑_تقرير_المحلات_المؤجرة.html', mime='text/html')
        else:
            st.info("لا توجد محلات مؤجرة لعرضها في النظرة العامة حالياً.")

    # 2. التحصيل وسندات القبض (محدث بإضافة ميزة الطباعة المستقرة)
    with tab2:
        st.subheader("تسجيل الدفعات وإصدار السندات")
        rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
        if rented_shops:
            with st.form("receipt_form"):
                r_shop = st.selectbox("المحل المُراد التحصيل منه:", rented_shops)
                tenant_name = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop]["المستأجر"].values[0]
                amount = st.number_input("المبلغ المحصل:", min_value=1)
                pay_method = st.selectbox("طريقة الدفع:", ["تحويل بنكي", "كاش", "شيك"])
                
                if st.form_submit_button("اعتماد وإصدار سند قبض"):
                    current_year = datetime.now().year
                    receipt_number = f"{current_year}-{len(st.session_state.transactions_db) + 1:04d}"
                    
                    idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop].index[0]
                    st.session_state.shops_db.at[idx, "المحصل"] += amount
                    
                    new_tx = pd.DataFrame([{
                        "التاريخ": datetime.now().strftime("%Y-%m-%d"), 
                        "رقم السند": receipt_number,
                        "رقم المحل": r_shop, 
                        "المستأجر": tenant_name, 
                        "طريقة الدفع": pay_method, 
                        "المبلغ": amount
                    }])
                    st.session_state.transactions_db = pd.concat([st.session_state.transactions_db, new_tx], ignore_index=True)
                    
                    # حفظ السند الجديد في الذاكرة لتفعيل الطباعة
                    st.session_state.latest_receipt = {
                        "التاريخ": datetime.now().strftime("%Y-%m-%d"),
                        "رقم السند": receipt_number,
                        "رقم المحل": r_shop,
                        "المستأجر": tenant_name,
                        "طريقة الدفع": pay_method,
                        "المبلغ": amount
                    }
                    st.rerun()

            # عرض السند الحالي الجاهز مع ميزة زر الطباعة
            if st.session_state.latest_receipt is not None:
                lr = st.session_state.latest_receipt
                st.success("تم تسجيل الدفعة بنجاح بنظام السندات المتسلسلة!")
                st.markdown(f"""
                <div style='border: 2px dashed #4CAF50; padding: 20px; border-radius: 10px; margin-top: 10px; background-color: #f9f9f9; color: #333;'>
                    <h2 style='text-align: center; color: #2E86C1; margin-bottom: 0;'>🧾 سند قبض مالي</h2>
                    <h4 style='text-align: center; color: #555; margin-top: 5px;'>رقم السند الموحد: {lr['رقم السند']}</h4>
                    <hr style='border: 1px solid #ddd;'>
                    <p><strong>التاريخ م:</strong> {lr['التاريخ']} م</p>
                    <p><strong>وصلنا من السيد/ة:</strong> {lr['المستأجر']} ( المستأجر لـ {lr['رقم المحل']} )</p>
                    <p><strong>مبلغ وقدره:</strong> <b style='color:#2E86C1; font-size:18px;'>{lr['المبلغ']:,.2f} ريال سعودي</b></p>
                    <p><strong>وذلك كدفعة عن طريق:</strong> {lr['طريقة الدفع']}</p>
                    <br>
                    <p style='text-align: left; font-weight:bold;'>توقيع المسؤول المالي والمحصل: .....................</p>
                </div>
                """, unsafe_allow_html=True)
                
                # توليد ملف الطباعة وإظهار الزر
                receipt_html_bytes = convert_receipt_to_pdf_html(lr)
                st.download_button(
                    label="🖨️ اضغط هنا لفتح السند والطباعة الفورية / الحفظ كـ PDF",
                    data=receipt_html_bytes,
                    file_name=f"سند_قبض_{lr['رقم السند']}.html",
                    mime="text/html"
                )
        else:
            st.warning("لا توجد محلات مؤجرة لإصدار سندات حالياً.")

    # 3. ديون المغادرين
    with tab3:
        st.subheader("جدولة وأرشيف ديون المستأجرين المغادرين")
        with st.form("historical_debt"):
            col1, col2 = st.columns(2)
            with col1:
                hist_year = st.text_input("السنة المالية (مثال: 2023):")
                hist_tenant = st.text_input("اسم المستأجر السابق المغادر:")
            with col2:
                hist_details = st.text_area("تفاصيل العقد والمديونية المتبقية:")
                hist_amount = st.number_input("إجمالي المبلغ المتبقي (ريال):", min_value=0)
                
            if st.form_submit_button("🎯 اعتماد وجدولة المديونية السابقة"):
                if hist_year and hist_tenant:
                    new_debt = pd.DataFrame([{
                        "السنة المالية": str(hist_year), 
                        "المستأجر السابق": str(hist_tenant), 
                        "تفاصيل العقد": str(hist_details), 
                        "المبلغ المتبقي": float(hist_amount)
                    }])
                    st.session_state.historical_debts_db = pd.concat([st.session_state.historical_debts_db, new_debt], ignore_index=True)
                    st.success("تم إدراج المديونية في الأرشيف المجدول بنجاح.")
        
        st.markdown("---")
        st.subheader("📊 جدول أرشيف الديون المجدولة")
        st.dataframe(st.session_state.historical_debts_db, use_container_width=True)
        if not st.session_state.historical_debts_db.empty:
            dc1, dc2 = st.columns(2)
            with dc1:
                debt_excel = convert_df_to_excel_csv(st.session_state.historical_debts_db)
                st.download_button(label="📥 تحميل أرشيف الديون (ملف Excel بأعمدة منفصلة)", data=debt_excel, file_name='📊_أرشيف_الديون.csv', mime='text/csv')
            with dc2:
                debt_pdf = convert_df_to_pdf_html(st.session_state.historical_debts_db, "تقرير أرشيف الديون المجدولة للمستأجرين المغادرين")
                st.download_button(label="📄 تصدير أرشيف الديون كتقرير (PDF)", data=debt_pdf, file_name='📑_تقرير_أرشيف_الديون.html', mime='text/html')

    # 4. إدارة المصروفات
    with tab4:
        st.subheader("إدارة وتسجيل المصروفات التشغيلية")
        with st.form("expenses_form"):
            col1, col2 = st.columns(2)
            with col1:
                exp_date = st.date_input("تاريخ الصرف:")
                exp_cat = st.text_input("بند ومجال الصرف (مثل: صيانة شبكة المياه، فواتير الكهرباء):")
            with col2:
                exp_amount = st.number_input("المبلغ المالي المصروف:", min_value=1)
                exp_notes = st.text_input("أي ملاحظات تود تدوينها:")
            
            if st.form_submit_button("🚨 تسجيل واعتماد المصروف"):
                new_exp = pd.DataFrame([{
                    "التاريخ": exp_date.strftime("%Y-%m-%d"), 
                    "بند الصرف": str(exp_cat), 
                    "المبلغ": float(exp_amount), 
                    "ملاحظات": str(exp_notes)
                }])
                st.session_state.expenses_db = pd.concat([st.session_state.expenses_db, new_exp], ignore_index=True)
                st.success("تم قفل وإدراج بند المصروف بالجدول بنجاح وتطابق كامل!")
                
        st.markdown("---")
        st.subheader("📋 سجل المصروفات الحالية المطابق")
        st.dataframe(st.session_state.expenses_db, use_container_width=True)

# ==================== لوحة المؤشرات الإحصائية ====================
elif menu == "لوحة المؤشرات والتحليلات":
    st.header("📊 لوحة المؤشرات والتحليلات الإستراتيجية للمجمع")
    df_shops = st.session_state.shops_db
    total_collected = df_shops["المحصل"].sum()
    total_historical_debt = st.session_state.historical_debts_db["المبلغ المتبقي"].sum() if not st.session_state.historical_debts_db.empty else 0
    total_expenses = st.session_state.expenses_db["المبلغ"].sum() if not st.session_state.expenses_db.empty else 0
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("إجمالي التحصيلات (العقود الحالية)", f"{total_collected:,} ريال")
    with col2:
        st.metric("إجمالي المصروفات التشغيلية", f"{total_expenses:,} ريال")
    with col3:
        st.metric("صافي الدخل الحالي للمشروع", f"{(total_collected - total_expenses):,} ريال")
    with col4:
        st.metric("إجمالي المديونيات المعلقة", f"{total_historical_debt:,} ريال")
        
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📊 توزيع حالة الـ 166 محل عقارياً")
        status_counts = df_shops["الحالة"].value_counts().reset_index()
        status_counts.columns = ["الحالة", "العدد"]
        fig_pie = px.pie(status_counts, values="العدد", names="الحالة", color_discrete_sequence=["#2ecc71", "#e74c3c", "#f1c40f"], hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
    with c2:
        st.subheader("⚖️ الإيرادات المحصلة مقابل المصروفات")
        fin_df = pd.DataFrame({"البند": ["الإيرادات", "المصروفات"], "المبلغ": [total_collected, total_expenses]})
        fig_bar = px.bar(fin_df, x="البند", y="المبلغ", color="البند", color_discrete_sequence=["#3498db", "#e67e22"])
        st.plotly_chart(fig_bar, use_container_width=True)
