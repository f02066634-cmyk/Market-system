import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="نظام إدارة عقارات المحلات", layout="wide")

# ==================== تهيئة قواعد البيانات ====================
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
    # هيكلة صارمة للأعمدة لتجنب مشاكل التطابق
    st.session_state.expenses_db = pd.DataFrame(columns=["التاريخ", "بند الصرف", "المبلغ", "ملاحظات"])

st.title("🏢 نظام إدارة وتحصيل عقارات المحلات التجارية")
st.markdown("---")

# ==================== القائمة الرئيسية ====================
menu = st.sidebar.radio("قائمة النظام الأساسية", ["عمليات التحصيل وإدخال البيانات", "لوحة المؤشرات والتحليلات"])

if menu == "عمليات التحصيل وإدخال البيانات":
    tab1, tab2, tab3, tab4 = st.tabs(["📝 تحديث المحلات والعقود", "💰 التحصيل وسندات القبض", "📂 أرشيف ديون المغادرين", "🛠️ إدارة المصروفات"])
    
    # 1. تحديث المحلات والعقود
    with tab1:
        st.subheader("تحديث بيانات العقود (نظام إيجار)")
        with st.form("shop_form"):
            shop_id = st.selectbox("اختر رقم المحل:", st.session_state.shops_db["رقم المحل"])
            col1, col2 = st.columns(2)
            with col1:
                status = st.selectbox("الحالة:", ["مؤجر", "شاغر", "تحت الصيانة"])
                tenant = st.text_input("اسم المستأجر:")
                rent = st.number_input("الإيجار السنوي:", min_value=0, value=15000)
            with col2:
                start_date = st.date_input("تاريخ بداية العقد:")
                end_date = st.date_input("تاريخ نهاية العقد:")
            
            if st.form_submit_button("حفظ وتحديث البيانات"):
                idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == shop_id].index[0]
                st.session_state.shops_db.at[idx, "الحالة"] = status
                st.session_state.shops_db.at[idx, "المستأجر"] = tenant if status == "مؤجر" else "-"
                st.session_state.shops_db.at[idx, "الإيجار السنوي"] = rent
                st.session_state.shops_db.at[idx, "بداية العقد"] = start_date.strftime("%Y-%m-%d")
                st.session_state.shops_db.at[idx, "نهاية العقد"] = end_date.strftime("%Y-%m-%d")
                st.success("تم تحديث بيانات العقد بنجاح!")
        
        st.markdown("---")
        st.subheader("📋 نظرة عامة: المحلات المؤجرة حالياً")
        rented_df = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]
        if not rented_df.empty:
            st.dataframe(rented_df[["رقم المحل", "المستأجر", "الإيجار السنوي", "بداية العقد", "نهاية العقد", "المحصل"]], use_container_width=True)
        else:
            st.info("لا توجد محلات مؤجرة حالياً.")

    # 2. التحصيل وسندات القبض
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
                    # توليد الرقم التسلسلي (السنة الحالية - رقم العملية)
                    current_year = datetime.now().year
                    receipt_number = f"{current_year}-{len(st.session_state.transactions_db) + 1:04d}"
                    
                    # تحديث البيانات
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
                    
                    # عرض السند للطباعة
                    st.success("تم تسجيل الدفعة بنجاح!")
                    st.markdown(f"""
                    <div style='border: 2px dashed #4CAF50; padding: 20px; border-radius: 10px; margin-top: 10px; background-color: #f9f9f9; color: #333;'>
                        <h2 style='text-align: center; color: #2E86C1; margin-bottom: 0;'>🧾 سند قبض</h2>
                        <h4 style='text-align: center; color: #555; margin-top: 5px;'>رقم السند: {receipt_number}</h4>
                        <hr style='border: 1px solid #ddd;'>
                        <p><strong>التاريخ:</strong> {datetime.now().strftime('%Y-%m-%d')} م</p>
                        <p><strong>استلمنا من السيد/ة:</strong> {tenant_name} ( {r_shop} )</p>
                        <p><strong>مبلغ وقدره:</strong> <b>{amount:,.2f}</b> ريال سعودي</p>
                        <p><strong>طريقة الدفع:</strong> {pay_method}</p>
                        <br>
                        <p style='text-align: left;'>التوقيع: .....................</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.warning("لا توجد محلات مؤجرة حالياً لإصدار سندات لها.")

    # 3. ديون المغادرين
    with tab3:
        st.subheader("أرشيف ديون المستأجرين المغادرين")
        with st.form("historical_debt"):
            col1, col2 = st.columns(2)
            with col1:
                hist_year = st.text_input("السنة المالية (مثال: 2023):")
                hist_tenant = st.text_input("اسم المستأجر السابق:")
            with col2:
                hist_details = st.text_area("تفاصيل العقد أو الملاحظات:")
                hist_amount = st.number_input("المبلغ المتبقي (المديونية):", min_value=0)
                
            if st.form_submit_button("إضافة للمديونيات السابقة"):
                if hist_year and hist_tenant:
                    new_debt = pd.DataFrame([{
                        "السنة المالية": str(hist_year), 
                        "المستأجر السابق": str(hist_tenant), 
                        "تفاصيل العقد": str(hist_details), 
                        "المبلغ المتبقي": float(hist_amount)
                    }])
                    st.session_state.historical_debts_db = pd.concat([st.session_state.historical_debts_db, new_debt], ignore_index=True)
                    st.success("تم أرشفة المديونية بنجاح.")
        
        st.dataframe(st.session_state.historical_debts_db, use_container_width=True)
        
        if not st.session_state.historical_debts_db.empty:
            # زر تصدير الإكسل (يدعم اللغة العربية)
            csv = st.session_state.historical_debts_db.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 تحميل أرشيف الديون (Excel)", data=csv, file_name='Historical_Debts.csv', mime='text/csv')

    # 4. المصروفات والصيانة
    with tab4:
        st.subheader("تسجيل المصروفات التشغيلية")
        with st.form("expenses_form"):
            col1, col2 = st.columns(2)
            with col1:
                exp_date = st.date_input("تاريخ الصرف:")
                exp_cat = st.text_input("بند الصرف (مثال: صيانة مصاعد، نظافة):")
            with col2:
                exp_amount = st.number_input("المبلغ المصروف:", min_value=1)
                exp_notes = st.text_input("ملاحظات إضافية:")
            
            if st.form_submit_button("تسجيل المصروف"):
                # فرض هيكلة الأعمدة بدقة
                new_exp = pd.DataFrame([{
                    "التاريخ": exp_date.strftime("%Y-%m-%d"), 
                    "بند الصرف": str(exp_cat), 
                    "المبلغ": float(exp_amount), 
                    "ملاحظات": str(exp_notes)
                }])
                st.session_state.expenses_db = pd.concat([st.session_state.expenses_db, new_exp], ignore_index=True)
                st.success("تم تسجيل المصروف بنجاح.")
                
        st.dataframe(st.session_state.expenses_db, use_container_width=True)

# ==================== لوحة المؤشرات ====================
elif menu == "لوحة المؤشرات والتحليلات":
    st.header("📊 لوحة المؤشرات والتحليلات الإستراتيجية")
    
    df_shops = st.session_state.shops_db
    total_collected = df_shops["المحصل"].sum()
    total_historical_debt = st.session_state.historical_debts_db["المبلغ المتبقي"].sum() if not st.session_state.historical_debts_db.empty else 0
    total_expenses = st.session_state.expenses_db["المبلغ"].sum() if not st.session_state.expenses_db.empty else 0
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("إجمالي الإيرادات المحصلة", f"{total_collected:,} ريال")
    with col2:
        st.metric("إجمالي المصروفات", f"{total_expenses:,} ريال")
    with col3:
        st.metric("صافي الدخل الحالي", f"{(total_collected - total_expenses):,} ريال")
    with col4:
        st.metric("إجمالي الديون للمغادرين", f"{total_historical_debt:,} ريال")
        
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("حالة المجمع التجاري")
        status_counts = df_shops["الحالة"].value_counts().reset_index()
        status_counts.columns = ["الحالة", "العدد"]
        fig_pie = px.pie(status_counts, values="العدد", names="الحالة", color_discrete_sequence=["#2ecc71", "#e74c3c", "#f1c40f"], hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
    with c2:
        st.subheader("الإيرادات مقابل المصروفات")
        fin_df = pd.DataFrame({"البند": ["الإيرادات", "المصروفات"], "المبلغ": [total_collected, total_expenses]})
        fig_bar = px.bar(fin_df, x="البند", y="المبلغ", color="البند", color_discrete_sequence=["#3498db", "#e67e22"])
        st.plotly_chart(fig_bar, use_container_width=True)
