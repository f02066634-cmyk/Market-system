import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# إعداد الصفحة وتفعيل التصميم الأساسي
st.set_page_config(page_title="نظام إدارة وتحصيل أسواق الشبرمي", layout="wide")

# ==================== حقن ثيم التصميم الجذاب والمودرن (CSS) ====================
st.markdown("""
    <style>
        .stApp { background: linear-gradient(135deg, #e0f2fe 0%, #f8fafc 100%) !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important; }
        [data-testid="stSidebar"] { display: none !important; }
        div[data-testid="stForm"], .stTabs, [data-testid="stMetricValue"] {
            background: rgba(255, 255, 255, 0.7) !important; backdrop-filter: blur(12px) !important; border-radius: 24px !important;
            padding: 25px !important; box-shadow: 0 8px 32px 0 rgba(31, 135, 211, 0.05) !important; border: 1px solid rgba(255, 255, 255, 0.6) !important; margin-bottom: 20px !important;
        }
        .stTabs [data-baseweb="tab-list"] { gap: 10px !important; background-color: rgba(241, 245, 249, 0.8) !important; padding: 8px !important; border-radius: 16px !important; }
        .stTabs [data-baseweb="tab"] { border-radius: 12px !important; padding: 10px 20px !important; background-color: transparent !important; transition: all 0.3s ease !important; font-weight: 600 !important; }
        .stTabs [aria-selected="true"] { background-color: #2E86C1 !important; color: white !important; box-shadow: 0 4px 12px rgba(46, 134, 193, 0.2) !important; }
        .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea {
            border-radius: 14px !important; border: 1px solid rgba(203, 213, 225, 0.8) !important; background-color: rgba(255, 255, 255, 0.9) !important; padding: 10px !important; color: #334155 !important;
        }
        .stButton button {
            border-radius: 14px !important; background: linear-gradient(90deg, #2E86C1 0%, #2471A3 100%) !important; color: white !important;
            border: none !important; padding: 12px 24px !important; font-weight: bold !important; box-shadow: 0 4px 15px rgba(46, 134, 193, 0.3) !important; transition: all 0.3s ease !important; width: 100% !important;
        }
        .stButton button:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 20px rgba(46, 134, 193, 0.4) !important; }
        h1, h2, h3 { color: #1e3a8a !important; font-weight: 700 !important; text-align: right !important; }
        .stDataFrame { border-radius: 16px !important; overflow: hidden !important; box-shadow: 0 4px 15px rgba(0,0,0,0.02) !important; }
    </style>
""", unsafe_allow_html=True)

# ==================== دوال المعالجة والطباعة ====================
def convert_df_to_excel_csv(df):
    csv_string = "sep=,\n" + df.to_csv(index=False)
    return csv_string.encode('utf-8-sig')

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
            @media print {{ .no-print {{ display: none !important; }} body {{ background-color: white; padding: 0; }} .report-card {{ box-shadow: none; padding: 0; max-width: 100%; }} }}
        </style>
    </head>
    <body>
        <div class="report-card">
            <h2>🏢 {title}</h2>
            <p class="date">تاريخ إصدار التقرير: {datetime.now().strftime('%Y-%m-%d | %I:%M %p')}</p>
            <hr style='border: 1px solid #eee;'>
            <table>
                <thead><tr>{"".join(f"<th>{col}</th>" for col in df.columns)}</tr></thead>
                <tbody>{"".join(f"<tr>{''.join(f'<td>{str(val)}</td>' for val in row)}</tr>" for row in df.values)}</tbody>
            </table>
            <br><br><button class="no-print" onclick="window.print()" style="padding: 14px; background-color: #2E86C1; color: white; border: none; border-radius: 6px; cursor: pointer; width: 100%; font-size: 16px; font-weight: bold;">📸 اضغط هنا لحفظ التقرير كـ PDF أو طباعته</button>
        </div>
    </body>
    </html>
    """
    return html.encode('utf-8')

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
            @media print {{ .no-print {{ display: none !important; }} body {{ background-color: white; padding: 0; }} .receipt-card {{ border: 2px dashed #4CAF50; box-shadow: none; max-width: 100%; }} }}
        </style>
    </head>
    <body>
        <div class="receipt-card">
            <h2>🧾 سند قبض مالي رسمي - أسواق الشبرمي</h2>
            <h4>رقم السند الموحد: {receipt_data['رقم السند']}</h4>
            <hr>
            <p><strong>تاريخ الإغلاق والاعتماد:</strong> {receipt_data['التاريخ']} م</p>
            <p><strong>وصلنا من السيد/ة:</strong> {receipt_data['المستأجر']} ( المستأجر لـ {receipt_data['رقم المحل']} )</p>
            <p><strong>إجمالي مبلغ الدفعة المكتملة:</strong> <b style='color:#2E86C1; font-size:18px;'>{receipt_data['إجمالي المتفق عليه']:,.2f} ريال سعودي</b></p>
            <p><strong>طريقة الدفع والاستلام:</strong> {receipt_data['طريقة الدفع']}</p>
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
    st.session_state.transactions_db = pd.DataFrame(columns=[
        "رقم السند", "تاريخ البدء", "تاريخ التحديث", "رقم المحل", 
        "المستأجر", "إجمالي المتفق عليه", "إجمالي المدفوع حتى الآن", "المبلغ المتبقي", "طريقة الدفع", "الحالة"
    ])
else:
    # تحديث قاعدة البيانات القديمة إن وجدت لإضافة عمود طريقة الدفع
    if "طريقة الدفع" not in st.session_state.transactions_db.columns:
        st.session_state.transactions_db["طريقة الدفع"] = "-"

if 'historical_debts_db' not in st.session_state:
    st.session_state.historical_debts_db = pd.DataFrame(columns=["السنة المالية", "المستأجر السابق", "تفاصيل العقد", "المبلغ المتبقي"])

if 'expenses_db' not in st.session_state:
    st.session_state.expenses_db = pd.DataFrame(columns=["التاريخ", "بند الصرف", "المبلغ", "ملاحظات"])

# ==================== واجهة النظام ====================
st.title("🏢 نظام إدارة وتحصيل أسواق الشبرمي")
st.markdown("---")

main_menu_tab1, main_menu_tab2 = st.tabs(["📥 عمليات التحصيل وإدخال البيانات", "📊 لوحة المؤشرات والتحليلات"])

with main_menu_tab1:
    tab1, tab2, tab3, tab4 = st.tabs(["📝 إدارة العقود والمحلات", "💰 التحصيل وسندات القبض", "📂 أرشيف ديون المغادرين", "🛠️ إدارة المصروفات"])
    
    # 1. إدارة العقود والمحلات
    with tab1:
        st.subheader("إدارة بيانات عقود الـ 166 محل")
        sub_tab1, sub_tab2 = st.tabs(["✍️ تسجيل عقد لمحل جديد", "🔄 تعديل بيانات عقد قائم"])
        
        with sub_tab1:
            available_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] != "مؤجر"]["رقم المحل"].tolist()
            if available_shops:
                with st.form("new_contract_form", clear_on_submit=True):
                    st.info("ملاحظة: هذه القائمة تعرض المحلات الشاغرة فقط.")
                    selected_shop = st.selectbox("اختر رقم المحل المُراد تأجيره:", available_shops)
                    col1, col2 = st.columns(2)
                    with col1:
                        tenant = st.text_input("اسم المستأجر الجديد:")
                        rent = st.number_input("الإيجار السنوي المتفق عليه:", min_value=0, value=15000)
                    with col2:
                        start_date = st.date_input("تاريخ بداية العقد:")
                        end_date = st.date_input("تاريخ نهاية العقد:")
                    
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
                            st.success(f"تم حفظ العقد بنجاح للمحل ({selected_shop})!")
                            st.rerun()
            else:
                st.success("🎉 جميع المحلات مؤجرة بالكامل حالياً!")
                
        with sub_tab2:
            rented_shops = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
            if rented_shops:
                selected_shop_edit = st.selectbox("اختر رقم المحل المؤجر لتعديله:", rented_shops, key="edit_shop_select")
                current_data = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop_edit].iloc[0]
                
                with st.form("edit_contract_form"):
                    st.warning(f"تعديل بيانات العقد للمحل: {selected_shop_edit}")
                    col1, col2 = st.columns(2)
                    with col1:
                        edit_status = st.selectbox("الحالة:", ["مؤجر", "شاغر", "تحت الصيانة"])
                        edit_tenant = st.text_input("اسم المستأجر الحالي:", value=current_data["المستأجر"])
                        edit_rent = st.number_input("الإيجار السنوي الحالي:", min_value=0, value=int(current_data["الإيجار السنوي"]))
                    with col2:
                        try: d1, d2 = datetime.strptime(current_data["بداية العقد"], "%Y-%m-%d"), datetime.strptime(current_data["نهاية العقد"], "%Y-%m-%d")
                        except: d1, d2 = datetime.now(), datetime.now()
                        edit_start = st.date_input("تعديل بداية العقد:", value=d1)
                        edit_end = st.date_input("تعديل نهاية العقد:", value=d2)
                    
                    if st.form_submit_button("🔄 تحديث بيانات العقد"):
                        idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == selected_shop_edit].index[0]
                        st.session_state.shops_db.at[idx, "الحالة"] = edit_status
                        st.session_state.shops_db.at[idx, "المستأجر"] = edit_tenant if edit_status == "مؤجر" else "-"
                        st.session_state.shops_db.at[idx, "الإيجار السنوي"] = edit_rent
                        st.session_state.shops_db.at[idx, "بداية العقد"] = edit_start.strftime("%Y-%m-%d")
                        st.session_state.shops_db.at[idx, "نهاية العقد"] = edit_end.strftime("%Y-%m-%d")
                        st.success(f"تم تحديث بيانات المحل ({selected_shop_edit}) بنجاح!")
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
                st.download_button(label="📥 تحميل (ملف Excel)", data=convert_df_to_excel_csv(rented_display_df), file_name='المحلات_المؤجرة.csv', mime='text/csv')
            with ec2:
                st.download_button(label="📄 تصدير كتقرير (PDF)", data=convert_df_to_pdf_html(rented_display_df, "المحلات المؤجرة"), file_name='تقرير_المحلات.html', mime='text/html')

    # 2. التحصيل وسندات القبض الذكية
    with tab2:
        st.subheader("💰 نظام تحصيل الدفعات وإدارة السندات")
        pay_sub_tab1, pay_sub_tab2 = st.tabs(["🆕 إنشاء سند دفعة جديد (مفتوح)", "🔄 تحديث وإغلاق السندات المفتوحة حالياً"])
        
        # الفرع أ: إنشاء دفعة جديدة (ديناميكي وبتفريغ تلقائي)
        with pay_sub_tab1:
            rented_shops_list = st.session_state.shops_db[st.session_state.shops_db["الحالة"] == "مؤجر"]["رقم المحل"].tolist()
            if rented_shops_list:
                # اختيار المحل خارج النموذج ليتم عرض اسم المستأجر ديناميكياً فور الاختيار
                r_shop = st.selectbox("اختر المحل المُراد تسجيل الدفعة عليه:", rented_shops_list, key="new_receipt_shop")
                tenant_name = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop]["المستأجر"].values[0]
                
                st.markdown(f"**👤 المستأجر المرتبط بالمحل المحدد:** <span style='color:#2E86C1; font-weight:bold; font-size:18px;'>{tenant_name}</span>", unsafe_allow_html=True)
                
                # إدخال المبالغ وطريقة الدفع داخل نموذج ليتم تفريغها بعد الحفظ
                with st.form("new_receipt_split_form", clear_on_submit=True):
                    pay_method_new = st.selectbox("طريقة الدفع والاستلام:", ["نقد", "إيداع بنكي"])
                    col1, col2 = st.columns(2)
                    with col1:
                        target_amount = st.number_input("المبلغ المتفق عليه للدفعة كاملة:", min_value=1, value=1000)
                    with col2:
                        paid_now = st.number_input("المبلغ المدفوع (الآن):", min_value=1, value=500)
                        
                    if st.form_submit_button("➕ حفظ الدفعة وفتح السند"):
                        if paid_now > target_amount:
                            st.error("خطأ: لا يمكن أن يكون المبلغ المدفوع أكبر من إجمالي الدفعة!")
                        else:
                            current_year = datetime.now().year
                            receipt_number = f"SH-{current_year}-{len(st.session_state.transactions_db) + 1:04d}"
                            remaining = target_amount - paid_now
                            status = "مغلق (مكتمل)" if remaining == 0 else "مفتوح (قيد التحصيل)"
                            
                            new_tx = pd.DataFrame([{
                                "رقم السند": receipt_number,
                                "تاريخ البدء": datetime.now().strftime("%Y-%m-%d"),
                                "تاريخ التحديث": datetime.now().strftime("%Y-%m-%d"),
                                "رقم المحل": r_shop,
                                "المستأجر": tenant_name,
                                "إجمالي المتفق عليه": float(target_amount),
                                "إجمالي المدفوع حتى الآن": float(paid_now),
                                "المبلغ المتبقي": float(remaining),
                                "طريقة الدفع": pay_method_new,
                                "الحالة": status
                            }])
                            st.session_state.transactions_db = pd.concat([st.session_state.transactions_db, new_tx], ignore_index=True)
                            
                            idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == r_shop].index[0]
                            st.session_state.shops_db.at[idx, "المحصل"] += paid_now
                            
                            if status == "مغلق (مكتمل)":
                                st.success(f"تم اكتمال الدفعة وإغلاق السند ({receipt_number}) جاهز للطباعة من الأرشيف بالأسفل!")
                            else:
                                st.warning(f"تم حفظ الدفعة! السند ({receipt_number}) 'مفتوح' ومتبقي عليه {remaining} ريال.")
            else:
                st.warning("لا توجد محلات مؤجرة لإصدار دفعات لها حالياً.")
                
        # الفرع ب: إدارة وتحديث السندات المفتوحة
        with pay_sub_tab2:
            tx_df = st.session_state.transactions_db
            open_tx = tx_df[tx_df["الحالة"] == "مفتوح (قيد التحصيل)"]
            
            if not open_tx.empty:
                st.info("💡 السندات المفتوحة التي لم تكتمل مبالغها بعد:")
                
                # عرض البيانات ديناميكياً خارج النموذج
                selected_open_id = st.selectbox("اختر رقم السند لتحديثه وإغلاقه:", open_tx["رقم السند"].tolist())
                target_tx_row = tx_df[tx_df["رقم السند"] == selected_open_id].iloc[0]
                
                st.markdown(f"""
                * **المحل والمستأجر:** {target_tx_row['رقم المحل']} - ({target_tx_row['المستأجر']})
                * **المدفوع سابقاً:** {target_tx_row['إجمالي المدفوع حتى الآن']:,} ريال | **المتبقي المطلوب:** <span style='color:red; font-weight:bold;'>{target_tx_row['المبلغ المتبقي']:,} ريال</span>
                * **طريقة الدفع السابقة:** {target_tx_row['طريقة الدفع']}
                """, unsafe_allow_html=True)
                
                # التحديث والتفريغ
                with st.form("update_split_receipt", clear_on_submit=True):
                    pay_method_update = st.selectbox("طريقة الدفع والاستلام للمبلغ المتبقي:", ["نقد", "إيداع بنكي"])
                    new_pay = st.number_input("أدخل المبلغ المدفوع الجديد حالياً:", min_value=1, value=int(target_tx_row['المبلغ المتبقي']))
                    
                    if st.form_submit_button("🔄 اعتماد التحديث"):
                        if new_pay > target_tx_row['المبلغ المتبقي']:
                            st.error("خطأ: المبلغ المدفوع أكبر من المتبقي على هذا السند!")
                        else:
                            idx = tx_df[tx_df["رقم السند"] == selected_open_id].index[0]
                            
                            updated_paid = target_tx_row['إجمالي المدفوع حتى الآن'] + new_pay
                            updated_remaining = target_tx_row['إجمالي المتفق عليه'] - updated_paid
                            updated_status = "مغلق (مكتمل)" if updated_remaining == 0 else "مفتوح (قيد التحصيل)"
                            
                            # دمج طريقة الدفع إذا اختلفت لتوثيقها بشكل كامل
                            old_method = target_tx_row['طريقة الدفع']
                            if pay_method_update not in old_method:
                                final_pay_method = f"{old_method} و {pay_method_update}"
                            else:
                                final_pay_method = old_method
                                
                            st.session_state.transactions_db.at[idx, "إجمالي المدفوع حتى الآن"] = updated_paid
                            st.session_state.transactions_db.at[idx, "المبلغ المتبقي"] = updated_remaining
                            st.session_state.transactions_db.at[idx, "طريقة الدفع"] = final_pay_method
                            st.session_state.transactions_db.at[idx, "الحالة"] = updated_status
                            st.session_state.transactions_db.at[idx, "تاريخ التحديث"] = datetime.now().strftime("%Y-%m-%d")
                            
                            shop_name = target_tx_row['رقم المحل']
                            s_idx = st.session_state.shops_db[st.session_state.shops_db["رقم المحل"] == shop_name].index[0]
                            st.session_state.shops_db.at[s_idx, "المحصل"] += new_pay
                            
                            st.success(f"تم تحديث الدفعة بنجاح! السند الآن {'مغلق وجاهز للطباعة' if updated_remaining == 0 else 'ما زال مفتوحاً'}")
                            st.rerun()
            else:
                st.success("🎉 ممتاز! لا توجد أي سندات مفتوحة أو معلقة حالياً.")

        # عرض أرشيف الدفعات وطباعة السندات
        st.markdown("---")
        st.subheader("🖨️ طباعة السندات المكتملة والأرشيف")
        if not st.session_state.transactions_db.empty:
            closed_tx = st.session_state.transactions_db[st.session_state.transactions_db["الحالة"] == "مغلق (مكتمل)"]
            if not closed_tx.empty:
                print_col1, print_col2 = st.columns([3, 1])
                with print_col1:
                    print_id = st.selectbox("اختر السند المكتمل لإصدار وطباعة الفاتورة:", closed_tx["رقم السند"].tolist())
                with print_col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    p_data = st.session_state.transactions_db[st.session_state.transactions_db["رقم السند"] == print_id].iloc[0]
                    receipt_dict = {
                        "رقم السند": p_data["رقم السند"],
                        "التاريخ": p_data["تاريخ التحديث"],
                        "المستأجر": p_data["المستأجر"],
                        "رقم المحل": p_data["رقم المحل"],
                        "إجمالي المتفق عليه": p_data["إجمالي المتفق عليه"],
                        "طريقة الدفع": p_data["طريقة الدفع"]
                    }
                    receipt_html_bytes = convert_receipt_to_pdf_html(receipt_dict)
                    st.download_button(label="🖨️ طباعة السند", data=receipt_html_bytes, file_name=f"سند_{print_id}.html", mime="text/html")
            
            st.dataframe(st.session_state.transactions_db, use_container_width=True)
        else:
            st.info("لا توجد سجلات دفعات مسجلة بالنظام حتى الآن.")

    # 3. ديون المغادرين
    with tab3:
        st.subheader("جدولة ديون المستأجرين المغادرين")
        with st.form("historical_debt", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                hist_year = st.text_input("السنة المالية (مثال: 2024):")
                hist_tenant = st.text_input("اسم المستأجر المغادر:")
            with col2:
                hist_details = st.text_area("تفاصيل المديونية:")
                hist_amount = st.number_input("المبلغ المتبقي (ريال):", min_value=0)
                
            if st.form_submit_button("🎯 اعتماد وجدولة المديونية"):
                if hist_year and hist_tenant:
                    new_debt = pd.DataFrame([{"السنة المالية": str(hist_year), "المستأجر السابق": str(hist_tenant), "تفاصيل العقد": str(hist_details), "المبلغ المتبقي": float(hist_amount)}])
                    st.session_state.historical_debts_db = pd.concat([st.session_state.historical_debts_db, new_debt], ignore_index=True)
                    st.success("تم إدراج المديونية بنجاح.")
        st.dataframe(st.session_state.historical_debts_db, use_container_width=True)

    # 4. إدارة المصروفات
    with tab4:
        st.subheader("تسجيل المصروفات التشغيلية")
        with st.form("expenses_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                exp_date = st.date_input("تاريخ الصرف:")
                exp_cat = st.text_input("بند ومجال الصرف:")
            with col2:
                exp_amount = st.number_input("المبلغ المصروف:", min_value=1)
                exp_notes = st.text_input("ملاحظات:")
            
            if st.form_submit_button("🚨 اعتماد المصروف"):
                new_exp = pd.DataFrame([{"التاريخ": exp_date.strftime("%Y-%m-%d"), "بند الصرف": str(exp_cat), "المبلغ": float(exp_amount), "ملاحظات": str(exp_notes)}])
                st.session_state.expenses_db = pd.concat([st.session_state.expenses_db, new_exp], ignore_index=True)
                st.success("تم إدراج المصروف بنجاح!")
        st.dataframe(st.session_state.expenses_db, use_container_width=True)

# القسم الثاني: لوحة المؤشرات والتحليلات
with main_menu_tab2:
    st.header("📊 لوحة المؤشرات الإستراتيجية")
    df_shops = st.session_state.shops_db
    total_collected = df_shops["المحصل"].sum()
    total_historical_debt = st.session_state.historical_debts_db["المبلغ المتبقي"].sum() if not st.session_state.historical_debts_db.empty else 0
    total_expenses = st.session_state.expenses_db["المبلغ"].sum() if not st.session_state.expenses_db.empty else 0
    
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("التحصيلات", f"{total_collected:,} ريال")
    with col2: st.metric("المصروفات", f"{total_expenses:,} - ريال")
    with col3: st.metric("صافي الدخل", f"{(total_collected - total_expenses):,} ريال")
    with col4: st.metric("الديون المعلقة", f"{total_historical_debt:,} ريال")
        
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        status_counts = df_shops["الحالة"].value_counts().reset_index(); status_counts.columns = ["الحالة", "العدد"]
        st.plotly_chart(px.pie(status_counts, values="العدد", names="الحالة", color_discrete_sequence=["#2ecc71", "#e74c3c", "#f1c40f"], hole=0.4), use_container_width=True)
    with c2:
        fin_df = pd.DataFrame({"البند": ["الإيرادات", "المصروفات"], "المبلغ": [total_collected, total_expenses]})
        st.plotly_chart(px.bar(fin_df, x="البند", y="المبلغ", color="البند", color_discrete_sequence=["#3498db", "#e67e22"]), use_container_width=True)
