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
    st.session_state.transactions_db = pd.DataFrame(columns=["التاريخ", "رقم المحل", "المستأجر", "طريقة الدفع", "المبلغ"])

if 'historical_debts_db' not in st.session_state:
    st.session_state.historical_debts_db = pd.DataFrame(columns=["السنة المالية", "المستأجر السابق", "تفاصيل العقد", "المبلغ المتبقي"])

if 'expenses_db' not in st.session_state:
    st.session_state.expenses_db = pd.DataFrame(columns=["التاريخ", "بند الصرف", "المبلغ", "ملاحظات"])

st.title("🏢 نظام إدارة وتحصيل عقارات المحلات التجارية")
st.markdown("---")

# ==================== القائمة الرئيسية (الفصل التام) ====================
menu = st.sidebar.radio("قائمة النظام الأساسية", ["عمليات التحصيل وإدخال البيانات", "لوحة المؤشرات والتحليلات"])

# ==================== القسم الأول: عمليات التحصيل وإدخال البيانات ====================
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
                st.file_uploader("إرفاق صورة العقد (PDF/JPG)", type=['pdf', 'jpg', 'png'])
            
            if st.form_submit_button("حفظ وتحديث البيانات"):
                idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == shop_id].index[0]
                st.session_state.shops_db.at[idx, "الحالة"] = status
                st.session_state.shops_db.at[idx, "المستأجر"] = tenant if status == "مؤجر" else "-"
                st.session_state.shops_db.at[idx, "الإيجار السنوي"] = rent
                st.session_state.shops_db.at[idx, "بداية العقد"] = start_date.strftime("%Y-%m-%d")
                st.session_state.shops_db.at[idx, "نهاية العقد"] = end_date.strftime("%Y-%m-%d")
                st.success("تم تحديث بيانات العقد بنجاح!")

    # 2. التحصيل وسندات القبض
    with tab2:
        st.subheader("تسجيل الدفعات وإصدار السندات")
        rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
        if rented_shops:
            with st.form("receipt_form"):
                r_shop = st.selectbox("المحل المُراد التحصيل منه:", rented_shops)
                tenant_name = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop]["المستأجر"].values[0]
                st.info(f"المستأجر: {tenant_name}")
                amount = st.number_input("المبلغ المحصل:", min_value=1)
                pay_method = st.selectbox("طريقة الدفع:", ["تحويل بنكي", "كاش", "شيك"])
                
                if st.form_submit_button("اعتماد وإصدار سند قبض"):
                    idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop].index[0]
                    st.session_state.shops_db.at[idx, "المحصل"] += amount
                    new_tx = pd.DataFrame([{"التاريخ": datetime.now().strftime("%Y-%m-%d"), "رقم المحل": r_shop, "المستأجر": tenant_name, "طريقة الدفع": pay_method, "المبلغ": amount}])
                    st.session_state.transactions_db = pd.concat([st.session_state.transactions_db, new_tx], ignore_index=True)
                    st.success(f"تم تسجيل الدفعة. يمكنك طباعة السند للمبلغ {amount} ريال ({pay_method}).")
                    st.button("🖨️ طباعة السند (PDF)")
        else:
            st.warning("لا توجد محلات مؤجرة حالياً لإصدار سندات لها.")

    # 3. ديون المغادرين (إدخال حر للسنة)
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
                    new_debt = pd.DataFrame([{"السنة المالية": hist_year, "المستأجر السابق": hist_tenant, "تفاصيل العقد": hist_details, "المبلغ المتبقي": hist_amount}])
                    st.session_state.historical_debts_db = pd.concat([st.session_state.historical_debts_db, new_debt], ignore_index=True)
                    st.success("تم أرشفة المديونية بنجاح.")
        st.dataframe(st.session_state.historical_debts_db, use_container_width=True)

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
                new_exp = pd.DataFrame([{"التاريخ": exp_date.strftime("%Y-%m-%d"), "بند الصرف": exp_cat, "المبلغ": exp_amount, "ملاحظات": exp_notes}])
                st.session_state.expenses_db = pd.concat([st.session_state.expenses_db, new_exp], ignore_index=True)
                st.success("تم تسجيل المصروف.")
        st.dataframe(st.session_state.expenses_db, use_container_width=True)
        if not st.session_state.expenses_db.empty:
            st.button("📥 تصدير جدول المصروفات (PDF/Excel)")

# ==================== القسم الثاني: لوحة المؤشرات (منفصلة تماماً) ====================
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
        st.metric("إجمالي المصروفات التشغيلية", f"{total_expenses:,} ريال")
    with col3:
        net_income = total_collected - total_expenses
        st.metric("صافي الدخل الحالي", f"{net_income:,} ريال")
    with col4:
        st.metric("إجمالي الديون المتأخرة (للمغادرين)", f"{total_historical_debt:,} ريال")
        
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
