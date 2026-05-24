import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# ضبط إعدادات الصفحة وجعلها متوافقة مع التصميم الاحترافي
st.set_page_config(page_title="نظام الإدارة العقارية المطور", layout="wide", initial_sidebar_state="expanded")

# ==================== نظام التنسيق والحقن الجمالي (Custom CSS) ====================
st.markdown("""
<style>
    /* استيراد خطوط احترافية وتطبيقها على كامل النظام */
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    
    html, body, [data-testid="stSidebar"], .stApp {
        font-family: 'Tajawal', sans-serif !important;
        direction: rtl;
        text-align: right;
    }
    
    /* تحسين شكل البطاقات المالية الإحصائية */
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border-right: 5px solid #2E86C1;
        margin-bottom: 15px;
        transition: transform 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-3px);
    }
    .metric-title { font-size: 14px; color: #7F8C8D; margin-bottom: 5px; font-weight: 500; }
    .metric-value { font-size: 22px; color: #2C3E50; font-weight: 700; }
    
    /* تصميم الأزرار الاحترافية */
    div.stButton > button {
        width: 100% !important;
        background-color: #2E86C1 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-size: 16px !important;
        font-weight: bold !important;
        box-shadow: 0 3px 6px rgba(46, 134, 193, 0.2) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:hover {
        background-color: #1B4F72 !important;
        box-shadow: 0 5px 12px rgba(46, 134, 193, 0.4) !important;
    }
    
    /* زر التنزيل المخصص */
    div.stDownloadButton > button {
        background-color: #27AE60 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 15px !important;
        font-size: 14px !important;
        box-shadow: 0 3px 6px rgba(39, 174, 96, 0.2) !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #1E8449 !important;
    }
    
    /* تحسين شكل النماذج والقوائم */
    .stForm {
        background-color: #FAFAFA !important;
        border-radius: 12px !important;
        border: 1px solid #E0E0E0 !important;
        padding: 25px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02) !important;
    }
    
    /* تصميم فخم للسند المالي قابل للطباعة */
    .premium-receipt {
        background: #ffffff;
        border: 1px solid #dcdde1;
        border-radius: 15px;
        padding: 30px;
        margin-top: 15px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        position: relative;
        overflow: hidden;
    }
    .premium-receipt::before {
        content: "";
        position: absolute;
        top: 0; right: 0; left: 0;
        height: 6px;
        background: linear-gradient(90deg, #2E86C1, #27AE60);
    }
</style>
""", unsafe_allow_html=True)

# ==================== دالة التصدير إلى إكسل حقيقي XLSX بنسبة 100% ====================
def convert_to_real_excel(df):
    output = io.BytesIO()
    # استخدام محرك openpyxl لإنشاء ملف إكسل رسمي معتمد تفهمه كل الحواسب والأجهزة
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير النظام')
    return output.getvalue()

# ==================== دالة توليد تقرير ذكي للطباعة كـ PDF متناسق ====================
def convert_df_to_pdf_html(df, title="تقرير النظام"):
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; direction: rtl; text-align: right; padding: 20px; background-color: #f5f5f5; }}
            .report-card {{ background-color: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); max-width: 950px; margin: auto; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; direction: rtl; }}
            th, td {{ border: 1px solid #e0e0e0; padding: 12px; text-align: center; }}
            th {{ background-color: #2E86C1; color: white; font-size: 14px; font-weight: bold; }}
            td {{ font-size: 13px; color: #2C3E50; }}
            tr:nth-child(even) {{ background-color: #f8f9fa; }}
            h2 {{ color: #2E86C1; text-align: center; margin-bottom: 5px; }}
            .date {{ text-align: left; font-size: 12px; color: #7F8C8D; }}
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
            <hr style='border: 1px solid #f1f2f6;'>
            <table>
                <thead>
                    <tr>{"".join(f"<th>{col}</th>" for col in df.columns)}</tr>
                </thead>
                <tbody>
                    {"".join(f"<tr>{''.join(f'<td>{str(val)}</td>' for val in row)}</tr>" for row in df.values)}
                </tbody>
            </table>
            <br><br>
            <button class="no-print" onclick="window.print()" style="padding: 14px; background-color: #2E86C1; color: white; border: none; border-radius: 8px; cursor: pointer; width: 100%; font-size: 16px; font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">📸 اضغط هنا لحفظ التقرير كـ PDF أو طباعته فوراً</button>
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

# عنوان النظام الرئيسي بتصميم عصري
st.markdown("<h1 style='text-align: center; color: #2C3E50; margin-bottom:10px;'>🏢 نظام إدارة وعقارات الـ 166 محل التجاري</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #7F8C8D; font-size:16px;'>منصة الإدارة الذكية والتحصيل وإصدار السندات والأرشفة</p>", unsafe_allow_html=True)
st.markdown("<hr style='margin-top:5px; margin-bottom:25px;'>", unsafe_allow_html=True)

# ==================== القائمة الجانبية المحدثة ====================
st.sidebar.markdown("<h3 style='text-align:center; color:#2E86C1;'>🧭 لوحة التحكم</h3>", unsafe_allow_html=True)
menu = st.sidebar.radio("انتقل بين الأقسام بحرية:", ["عمليات التحصيل وإدخل البيانات", "لوحة المؤشرات والتحليلات"])
st.sidebar.markdown("---")
st.sidebar.caption("💡 هذا النظام مصمم ليعمل بكفاءة عالية على شاشات الجوال والحواسب الشخصية.")

if menu == "عمليات التحصيل وإدخل البيانات":
    tab1, tab2, tab3, tab4 = st.tabs(["📝 إدارة العقود والمحلات", "💰 التحصيل وسندات القبض", "📂 أرشيف ديون المغادرين", "🛠️ إدارة المصروفات"])
    
    # 1. إدارة العقود والمحلات (إدخال وتعديل منفصلين تماماً)
    with tab1:
        st.markdown("<h3 style='color:#2C3E50;'>📝 إدارة عقود المحلات وتحديثها</h3>", unsafe_allow_html=True)
        action_type = st.radio("اختر الإجراء الإداري المطلوب تفعيله الآن:", ["✍️ تسجيل عقد لمحل جديد (إدخال جديد)", "🔄 تعديل بيانات عقد قائم (تحديث بيانات)"], horizontal=True)
        
        # --- الحالة الأولى: تسجيل عقد جديد ---
        if action_type == "✍️ تسجيل عقد لمحل جديد (إدخال جديد)":
            available_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] != "مؤجر"]["رقم المحل"].tolist()
            
            if available_shops:
                with st.form("new_contract_form"):
                    st.markdown("<p style='color:#3498db; font-weight:bold;'>💡 تظهر هنا المحلات الشاغرة فقط (تختفي تلقائياً فور حفظها كمؤجرة):</p>", unsafe_allow_html=True)
                    selected_shop = st.selectbox("اختر رقم المحل الشاغر لربطه بعقد جديد:", available_shops)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        tenant = st.text_input("اسم المستأجر الجديد بالكامل:")
                        rent = st.number_input("قيمة الإيجار السنوي (ريال):", min_value=0, value=15000)
                    with col2:
                        start_date = st.date_input("تاريخ سريان بداية العقد:")
                        end_date = st.date_input("تاريخ نهاية العقد المحدد:")
                    
                    if st.form_submit_button("💾 حفظ وإدراج العقد الجديد في النظام"):
                        if tenant.strip() == "" or tenant == "-":
                            st.error("خطأ: يرجى كتابة اسم مستأجر صحيح ومعتمد.")
                        else:
                            idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].index[0]
                            st.session_state.shops_db.at[idx, "الحالة"] = "مؤجر"
                            st.session_state.shops_db.at[idx, "المستأجر"] = tenant
                            st.session_state.shops_db.at[idx, "الإيجار السنوي"] = rent
                            st.session_state.shops_db.at[idx, "بداية العقد"] = start_date.strftime("%Y-%m-%d")
                            st.session_state.shops_db.at[idx, "نهاية العقد"] = end_date.strftime("%Y-%m-%d")
                            st.success(f"ممتاز! تم إدخال وحفظ العقد بنجاح للمحل ({selected_shop}) وتم نقله آلياً لقسم المحلات المؤجرة.")
                            st.rerun()
            else:
                st.success("🎉 إنجاز مذهل! جميع المحلات الـ 166 ممتلئة ومؤجرة بالكامل حالياً.")
                
        # --- الحالة الثانية: تعديل عقد قائم ---
        else:
            rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
            
            if rented_shops:
                selected_shop = st.selectbox("اختر رقم المحل المؤجر لغرض تعديل بياناته الحالية:", rented_shops)
                current_data = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].iloc[0]
                
                with st.form("edit_contract_form"):
                    st.warning(f"⚠️ وضع التعديل النشط للمحل: {selected_shop}")
                    col1, col2 = st.columns(2)
                    with col1:
                        edit_status = st.selectbox("تحديث حالة المحل الحالية:", ["مؤجر", "شاغر", "تحت الصيانة"])
                        edit_tenant = st.text_input("اسم المستأجر القائم:", value=current_data["المستأجر"])
                        edit_rent = st.number_input("تعديل قيمة الإيجار السنوي الحالية:", min_value=0, value=int(current_data["الإيجار السنوي"]))
                    with col2:
                        try:
                            d1 = datetime.strptime(current_data["بداية العقد"], "%Y-%m-%d")
                            d2 = datetime.strptime(current_data["نهاية العقد"], "%Y-%m-%d")
                        except:
                            d1 = datetime.now()
                            d2 = datetime.now()
                            
                        edit_start = st.date_input("تغيير بداية العقد:", value=d1)
                        edit_end = st.date_input("تغيير نهاية العقد:", value=d2)
                    
                    if st.form_submit_button("🔄 تحديث وحفظ بيانات العقد المعدلة"):
                        idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop].index[0]
                        st.session_state.shops_db.at[idx, "الحالة"] = edit_status
                        st.session_state.shops_db.at[idx, "المستأجر"] = edit_tenant if edit_status == "مؤجر" else "-"
                        st.session_state.shops_db.at[idx, "الإيجار السنوي"] = edit_rent
                        st.session_state.shops_db.at[idx, "بداية العقد"] = edit_start.strftime("%Y-%m-%d")
                        st.session_state.shops_db.at[idx, "نهاية العقد"] = edit_end.strftime("%Y-%m-%d")
                        st.success(f"تم اعتماد وتحديث تعديلات العقد للمحل ({selected_shop}) بنجاح تام.")
                        st.rerun()
            else:
                st.info("لا يوجد أي محلات مؤجرة حالياً لتعديلها في النظام.")
        
        # --- النظرة العامة على العقود والتصدير المطور ---
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<h4 style='color:#2C3E50;'>📋 النظرة العامة والتقارير: العقود المؤجرة الحالية</h4>", unsafe_allow_html=True)
        rented_display_df = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"][["رقم المحل", "المستأجر", "الإيجار السنوي", "بداية العقد", "نهاية العقد", "المحصل"]]
        
        if not rented_display_df.empty:
            st.dataframe(rented_display_df, use_container_width=True)
            
            # أزرار تصدير منسقة واحترافية وبدون أي مشاكل ترميز لغوية
            ec1, ec2 = st.columns(2)
            with ec1:
                excel_bytes = convert_to_real_excel(rented_display_df)
                st.download_button(label="📥 تحميل الجدول كملف Excel (مباشر ومنظم بأعمدة)", data=excel_bytes, file_name='📊_المحلات_المؤجرة.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            with ec2:
                pdf_html_data = convert_df_to_pdf_html(rented_display_df, "تقرير النظرة العامة للمحلات المؤجرة")
                st.download_button(label="📄 تصدير وتجهيز التقرير للطباعة كـ PDF", data=pdf_html_data, file_name='📑_تقرير_المحلات_المؤجرة.html', mime='text/html')
        else:
            st.info("النظام جاهز بانتظار إدخال عقود لعرضها في جدول النظرة العامة.")

    # 2. التحصيل وسندات القبض الفخمة
    with tab2:
        st.markdown("<h3 style='color:#2C3E50;'>💰 تسجيل عمليات التحصيل والتدفقات المالية</h3>", unsafe_allow_html=True)
        rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
        if rented_shops:
            with st.form("receipt_form"):
                r_shop = st.selectbox("اختر رقم المحل المُراد إيداع مبلغه وتحصيله:", rented_shops)
                tenant_name = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop]["المستأجر"].values[0]
                amount = st.number_input("المبلغ المالي المستلم حالياً (ريال):", min_value=1)
                pay_method = st.selectbox("طريقة دفع واستلام المبلغ:", ["تحويل بنكي مباشر", "نقداً (كاش)", "شيك مصرفي معتمد"])
                
                if st.form_submit_button("اعتماد الحركة وإصدار سند القبض"):
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
                    
                    st.success("تم إدراج الحركة بنجاح في السجل المالي للنظام!")
                    # تصميم فخم وراقي جداً للسند المالي يظهر مباشرة للمستخدم للاحتفاظ به أو تصويره الشاشة
                    st.markdown(f"""
                    <div class="premium-receipt">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                            <div style="text-align: right;">
                                <h3 style="margin: 0; color: #2C3E50; font-size:22px;">🧾 سند قبض مالي رسمي</h3>
                                <p style="margin: 2px 0 0 0; color: #7F8C8D; font-size:13px;">رقم السند الموحد: <b style="color:#2E86C1;">{receipt_number}</b></p>
                            </div>
                            <div style="text-align: left; color:#7F8C8D; font-size:13px;">
                                <p style="margin: 0;"><b>التاريخ:</b> {datetime.now().strftime('%Y-%m-%d')} م</p>
                                <p style="margin: 2px 0 0 0;"><b>الوقت:</b> {datetime.now().strftime('%I:%M %p')}</p>
                            </div>
                        </div>
                        <hr style="border:0; border-top: 1px solid #e0e0e0; margin-bottom: 20px;">
                        <div style="font-size: 15px; color: #34495E; line-height: 1.8;">
                            <p>استلمنا من السيد / السيدة: <span style="color:#2C3E50; font-weight:bold; background:#f1f2f6; padding:3px 8px; border-radius:4px;">{tenant_name}</span></p>
                            <p>المستأجر للموقع العقاري: <span style="font-weight:bold; color:#2E86C1;">{r_shop}</span></p>
                            <p>مبلغاً وقدره الإجمالي: <span style="font-size:18px; color:#27AE60; font-weight:bold; background:#E8F8F5; padding:4px 10px; border-radius:6px;">{amount:,.2f} ريال سعودي</span></p>
                            <p>وذلك لقاء الإيجار والدفع المتفق عليه بواسطة: <b>{pay_method}</b></p>
                        </div>
                        <br>
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 25px; font-size:13px; color:#7F8C8D;">
                            <p>توقيع الإدارة المالية والتحصيل: ...........................</p>
                            <p style="font-style: italic;">نظام الإدارة الرقمي الذكي للمجمع</p>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.warning("يتطلب النظام وجود محلات مؤجرة نشطة أولاً لإصدار السندات لها.")

    # 3. ديون المغادرين (المجدولة والأرشفة مع تصدير مطور حقيقي)
    with tab3:
        st.markdown("<h3 style='color:#2C3E50;'>📂 أرشيف وجدولة ديون المستأجرين المغادرين</h3>", unsafe_allow_html=True)
        with st.form("historical_debt"):
            col1, col2 = st.columns(2)
            with col1:
                hist_year = st.text_input("السنة المالية المعلقة (مثال: 2023):")
                hist_tenant = st.text_input("اسم المستأجر السابق المغادر:")
            with col2:
                hist_details = st.text_area("أسباب المديونية وتفاصيل العقد السابق:")
                hist_amount = st.number_input("إجمالي المديونية المتبقية عليه بذمته (ريال):", min_value=0)
                
            if st.form_submit_button("🎯 اعتماد وجدولة المديونية التاريخية بالأرشيف"):
                if hist_year and hist_tenant:
                    new_debt = pd.DataFrame([{
                        "السنة المالية": str(hist_year), 
                        "المستأجر السابق": str(hist_tenant), 
                        "تفاصيل العقد": str(hist_details), 
                        "المبلغ المتبقي": float(hist_amount)
                    }])
                    st.session_state.historical_debts_db = pd.concat([st.session_state.historical_debts_db, new_debt], ignore_index=True)
                    st.success("تم ترحيل وجدولة المديونية بنجاح في أرشيف الديون الشامل.")
        
        st.markdown("<br><h5>📊 جدول أرشيف الديون التاريخية المجدولة</h5>", unsafe_allow_html=True)
        st.dataframe(st.session_state.historical_debts_db, use_container_width=True)
        
        if not st.session_state.historical_debts_db.empty:
            dc1, dc2 = st.columns(2)
            with dc1:
                debt_excel_bytes = convert_to_real_excel(st.session_state.historical_debts_db)
                st.download_button(label="📥 تحميل أرشيف الديون كملف Excel (أعمدة سليمة)", data=debt_excel_bytes, file_name='📊_أرشيف_الديون.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            with dc2:
                debt_pdf = convert_df_to_pdf_html(st.session_state.historical_debts_db, "تقرير أرشيف الديون المجدولة للمستأجرين المغادرين")
                st.download_button(label="📄 تصدير أرشيف الديون كتقرير للطباعة (PDF)", data=debt_pdf, file_name='📑_تقرير_أرشيف_الديون.html', mime='text/html')

    # 4. إدارة المصروفات المطابقة تماماً
    with tab4:
        st.markdown("<h3 style='color:#2C3E50;'>🛠️ إدارة وتسجيل المصروفات التشغيلية والصيانة</h3>", unsafe_allow_html=True)
        with st.form("expenses_form"):
            col1, col2 = st.columns(2)
            with col1:
                exp_date = st.date_input("تاريخ الصرف والاتفاق:")
                exp_cat = st.text_input("بند ومجال الصرف (مثال: صيانة بوابات، فواتير إنارة):")
            with col2:
                exp_amount = st.number_input("إجمالي المبلغ المالي المصروف:", min_value=1)
                exp_notes = st.text_input("ملاحظات إدارية ملحقة:")
            
            if st.form_submit_button("🚨 تسجيل واعتماد المصروف بالدفاتر"):
                new_exp = pd.DataFrame([{
                    "التاريخ": exp_date.strftime("%Y-%m-%d"), 
                    "بند الصرف": str(exp_cat), 
                    "المبلغ": float(exp_amount), 
                    "ملاحظات": str(exp_notes)
                }])
                st.session_state.expenses_db = pd.concat([st.session_state.expenses_db, new_exp], ignore_index=True)
                st.success("تم قفل وإدراج قيد بند المصروف بالجدول بامتثال كامل وتطابق تام في الهيكلية والأعمدة!")
                
        st.markdown("<br><h5>📋 سجل بنود المصروفات التشغيلية المطابق</h5>", unsafe_allow_html=True)
        st.dataframe(st.session_state.expenses_db, use_container_width=True)

# ==================== لوحة المؤشرات التنفيذية (Executive Dashboard) ====================
elif menu == "لوحة المؤشرات والتحليلات":
    st.markdown("<h2 style='color:#2C3E50; margin-bottom:5px;'>📊 لوحة المؤشرات المالية والتحليلات الإستراتيجية</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#7F8C8D;'>ملخص تنفيذي فوري لأداء الـ 166 محل عقارياً ومتابعة الأرباح والديون</p>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    df_shops = st.session_state.shops_db
    total_collected = df_shops["المحصل"].sum()
    total_historical_debt = st.session_state.historical_debts_db["المبلغ المتبقي"].sum() if not st.session_state.historical_debts_db.empty else 0
    total_expenses = st.session_state.expenses_db["المبلغ"].sum() if not st.session_state.expenses_db.empty else 0
    net_income = total_collected - total_expenses
    
    # بناء وتصميم بطاقات المؤشرات الاحترافية الفخمة بدلاً من المؤشرات البدائية الافتراضية
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-card" style="border-right-color: #27AE60;">
            <div class="metric-title">إجمالي التدفقات والتحصيلات</div>
            <div class="metric-value">{total_collected:,.2f} ريال</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card" style="border-right-color: #E67E22;">
            <div class="metric-title">إجمالي المصروفات التشغيلية</div>
            <div class="metric-value">{total_expenses:,.2f} ريال</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        color = "#2E86C1" if net_income >= 0 else "#C0392B"
        st.markdown(f"""
        <div class="metric-card" style="border-right-color: {color};">
            <div class="metric-title">صافي الأرباح والدخل الحالي</div>
            <div class="metric-value">{net_income:,.2f} ريال</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card" style="border-right-color: #8E44AD;">
            <div class="metric-title">ديون معلقة (المغادرين)</div>
            <div class="metric-value">{total_historical_debt:,.2f} ريال</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    
    # عرض الرسوم البيانية التفاعلية الحديثة
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<h4 style='text-align: center; color: #34495E;'>📊 الهيكل العقاري والإشغال للـ 166 محل</h4>", unsafe_allow_html=True)
        status_counts = df_shops["الحالة"].value_counts().reset_index()
        status_counts.columns = ["الحالة", "العدد"]
        fig_pie = px.pie(status_counts, values="العدد", names="الحالة", color_discrete_sequence=["#2ecc71", "#e74c3c", "#f1c40f"], hole=0.45)
        fig_pie.update_layout(font_family="Tajawal", margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_pie, use_container_width=True)
    with c2:
        st.markdown("<h4 style='text-align: center; color: #34495E;'>⚖️ الميزان المالي الافتراضي (الإيراد مقابل الصرف)</h4>", unsafe_allow_html=True)
        fin_df = pd.DataFrame({"البند المالي": ["الإيرادات المحصلة", "المصروفات العامة"], "القيمة المالية (ريال)": [total_collected, total_expenses]})
        fig_bar = px.bar(fin_df, x="البند المالي", y="القيمة المالية (ريال)", color="البند المالي", color_discrete_sequence=["#3498db", "#e67e22"])
        fig_bar.update_layout(font_family="Tajawal", showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_bar, use_container_width=True)
