import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime, timedelta
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- 1. تنظیمات اولیه صفحه (باید اولین دستور باشد) ---
st.set_page_config(
    page_title="گزارشگر مالی مدرن",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed" # سایدبار را پیش‌فرض می‌بندیم
)

# --- 2. تزریق CSS مدرن و انیمیشن‌ها ---
st.markdown("""
    <style>
        /* وارد کردن فونت وزیرمتن از گوگل فونت */
        @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;700;900&display=swap');

        /* تنظیمات پایه و RTL */
        html, body, [class*="css"] {
            font-family: 'Vazirmatn', sans-serif !important;
            direction: rtl;
            text-align: right;
        }
        
        /* پس‌زمینه کلی تمیز و مینیمال */
        .stApp {
            background-color: #F9FAFB; /* رنگ پس‌زمینه خیلی روشن شبیه iOS */
        }

        /* --- Hero Section Animation --- */
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .hero-container {
            text-align: center;
            padding: 2rem 0;
            animation: fadeUp 0.8s ease-out;
        }
        .hero-title {
            font-size: 2.5rem !important;
            font-weight: 900 !important;
            background: linear-gradient(45deg, #111827, #4B5563); /* گرادینت ملایم تیره */
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem !important;
        }
        .hero-subtitle {
            font-size: 1.1rem !important;
            color: #6B7280; /* خاکستری ملایم */
            font-weight: 300 !important;
        }

        /* --- Modern Cards (Metric & Info) --- */
        div[data-testid="stMetric"], .modern-card {
            background-color: #FFFFFF !important;
            border-radius: 16px !important; /* گوشه‌های گرد */
            padding: 20px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03) !important; /* سایه نرم */
            border: 1px solid #F3F4F6 !important;
            transition: all 0.3s ease;
        }
        div[data-testid="stMetric"]:hover, .modern-card:hover {
             box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04) !important;
             transform: translateY(-2px);
        }
        /* رنگ و سایز اعداد در کارت‌ها */
        div[data-testid="stMetricValue"] {
            font-size: 1.8rem !important;
            font-weight: 700 !important;
            color: #1F2937 !important;
        }
        div[data-testid="stMetricLabel"] {
            font-size: 0.9rem !important;
            color: #9CA3AF !important;
        }

        /* --- Tabs Styling --- */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: transparent;
            padding-bottom: 10px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: nowrap;
            background-color: #FFFFFF;
            border-radius: 12px !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            color: #4B5563;
            font-weight: 400;
            transition: all 0.2s;
            border: none !important;
        }
        .stTabs [aria-selected="true"] {
            background-color: #EFF6FF !important; /* آبی خیلی روشن برای تب فعال */
            color: #1D4ED8 !important; /* آبی تیره */
            font-weight: 700 !important;
        }

        /* --- Buttons --- */
        .stButton > button {
            border-radius: 12px !important;
            font-weight: 600 !important;
            padding-top: 0.6rem !important;
            padding-bottom: 0.6rem !important;
        }
        /* دکمه‌های اصلی */
        div[data-testid="stFileUploader"] button {
             background-color: #1D4ED8 !important;
             color: white !important;
             border: none !important;
        }

        /* --- Subtle Helper Text --- */
        .subtle-text {
            font-size: 0.85rem;
            color: #9CA3AF;
            margin-top: 5px;
            font-weight: 300;
        }
        
        /* مخفی کردن منوی همبرگری استریم‌لیت برای ظاهر تمیزتر */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
    </style>
""", unsafe_allow_html=True)

# --- 3. توابع کمکی و داده ساز ---

@st.cache_data
def generate_dummy_data():
    """تولید داده‌های تصادفی برای نمایش قابلیت‌ها"""
    dates = pd.date_range(end=datetime.today(), periods=30)
    data = []
    descriptions = [
        "خرید کارتخوان", "انتقال از کارت 6037...", "کارمزد تراکنش", 
        "واریز مدرن سامانه غذارسان اطلس (اسنپ)", "انتقال وجه پایا", "خرید شارژ"
    ]
    balance = 50000000
    for date in dates:
        num_transactions = np.random.randint(3, 8)
        for i in range(num_transactions):
            desc = np.random.choice(descriptions)
            amount = np.random.randint(100000, 5000000)
            
            deposit = 0
            withdrawal = 0
            
            if "کارمزد" in desc or "انتقال وجه" in desc or "خرید" in desc:
                withdrawal = amount / 10  # مبالغ برداشتی کوچکتر
                balance -= withdrawal
            else:
                deposit = amount
                balance += deposit
                
            data.append({
                'Date': date,
                'Time': datetime.now().time(),
                'Description': desc,
                'Withdrawal': withdrawal,
                'Deposit': deposit,
                'Balance': balance
            })
    return pd.DataFrame(data)

@st.cache_data(show_spinner=False)
def process_data(df, tax_rate_percent, keywords):
    """پردازش هسته مرکزی داده‌ها"""
    if 'Date' not in df.columns: return pd.DataFrame()
    df = df.dropna(subset=['Date']).copy()
    unique_dates = df['Date'].sort_values().unique()

    for col in ['Deposit', 'Withdrawal', 'Balance']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Description'] = df['Description'].astype(str)
    
    # اعمال فیلترها بر اساس ورودی
    card_mask = df['Description'].str.contains(keywords.get('card', ''), na=False)
    fee_mask = df['Description'].str.contains(keywords.get('fee', ''), na=False)
    snap_mask = df['Description'].str.contains(keywords.get('snap', ''), na=False)

    # تجمیع
    grouped = df.groupby('Date')
    report = pd.DataFrame(index=unique_dates)
    report['کارتخوان'] = df[card_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    report['اسنپ'] = df[snap_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    report['کارمزدها'] = df[fee_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
    report['مانده نهایی'] = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last().reindex(unique_dates, fill_value=0)

    # محاسبات
    tax_factor = 1 + (tax_rate_percent / 100)
    # فرض: مالیات فقط به درآمد کارتخوان تعلق می‌گیرد
    report['فروش خالص (کارتخوان)'] = report['کارتخوان'] / tax_factor
    report['مالیات برآوردی'] = report['کارتخوان'] - report['فروش خالص (کارتخوان)']
    report['کل درآمد ناخالص'] = report['کارتخوان'] + report['اسنپ']
    report['سود عملیاتی'] = report['فروش خالص (کارتخوان)'] + report['اسنپ'] - report['کارمزدها']

    return report.reset_index().rename(columns={'index': 'تاریخ'})

def create_excel_report(df):
    """تولید فایل اکسل خروجی"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش مالی"
    ws.sheet_view.rightToLeft = True

    # استایل‌ها
    header_font = Font(bold=True, size=11, name='Tahoma', color='FFFFFF')
    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid") # رنگ تیره مدرن
    regular_font = Font(size=10, name='Tahoma')
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin", color="E5E7EB"), right=Side(style="thin", color="E5E7EB"), top=Side(style="thin", color="E5E7EB"), bottom=Side(style="thin", color="E5E7EB"))

    cols_to_export = ['تاریخ', 'کارتخوان', 'اسنپ', 'کل درآمد ناخالص', 'مالیات برآوردی', 'کارمزدها', 'سود عملیاتی', 'مانده نهایی']
    ws.append(cols_to_export)

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for row in df[cols_to_export].itertuples(index=False, name=None):
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = regular_font
            cell.alignment = center_align
            cell.border = thin_border
            if cell.column_letter != 'A': cell.number_format = '#,##0'

    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18

    if ws.max_row >= 2:
        tab = Table(displayName="FinancialTable", ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- 4. رابط کاربری اصلی ---

def main():
    # --- Hero Section (Animated Title) ---
    st.markdown("""
        <div class="hero-container">
            <h1 class="hero-title">داشبورد هوشمند مالی</h1>
            <p class="hero-subtitle">تحلیل خودکار تراکنش‌ها، تفکیک درآمدها و محاسبه مالیات در یک نگاه</p>
        </div>
    """, unsafe_allow_html=True)

    # --- State Management (برای نگهداری داده بین تب‌ها) ---
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = pd.DataFrame()
    if 'is_demo_data' not in st.session_state:
        st.session_state.is_demo_data = True # پیش‌فرض حالت دمو است

    # --- Tabs Navigation ---
    tab_dashboard, tab_details, tab_settings = st.tabs(["💎 پیشخوان", "📄 جزئیات تراکنش‌ها", "⚙️ تنظیمات و ورودی"])

    # ==================== Tab 3: تنظیمات و ورودی (اول پردازش می‌شود) ====================
    with tab_settings:
        st.markdown("### 📥 ورودی داده‌ها و پیکربندی")
        
        col_upload, col_params = st.columns([1, 1.5], gap="large")
        
        with col_upload:
            st.markdown("""
                <div class="modern-card">
                    <h4>آپلود فایل اکسل</h4>
                    <p class="subtle-text" style="margin-bottom: 15px;">فایل استاندارد گردش حساب بانکی خود را اینجا رها کنید.</p>
                """, unsafe_allow_html=True)
            uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if uploaded_file:
                 st.session_state.is_demo_data = False
            else:
                 st.session_state.is_demo_data = True
                 st.info("💡 در حال نمایش داده‌های نمونه (Demo). برای تحلیل واقعی، فایل خود را آپلود کنید.")

        with col_params:
             st.markdown("""
                <div class="modern-card">
                    <h4>پارامترهای محاسباتی</h4>
                """, unsafe_allow_html=True)
             tax_rate = st.slider("نرخ مالیات بر ارزش افزوده (%)", 0, 15, 10)
             st.markdown('<p class="subtle-text">این نرخ برای جداسازی مالیات از فروش کارتخوان استفاده می‌شود.</p>', unsafe_allow_html=True)
             
             with st.expander("تنظیمات پیشرفته کلمات کلیدی"):
                 k_card = st.text_input("کلیدواژه کارتخوان", "انتقال از")
                 k_snap = st.text_input("کلیدواژه اسنپ", "مدرن سامانه")
                 k_fee = st.text_input("کلیدواژه کارمزد", "کارمزد")
                 keywords = {'card': k_card, 'snap': k_snap, 'fee': k_fee}
             st.markdown("</div>", unsafe_allow_html=True)

        # --- منطق پردازش داده (واقعی یا دمو) ---
        try:
            if uploaded_file:
                df_raw = pd.read_excel(uploaded_file, skiprows=2)
                # استانداردسازی ساده ستون‌ها
                expected_cols = ['Index', 'Branch', 'Code', 'Date', 'Time', 'Doc', 'Receipt', 'Check', 'Description', 'Withdrawal', 'Deposit', 'Balance', 'Notes']
                if len(df_raw.columns) >= len(expected_cols):
                     df_raw.columns = expected_cols + list(df_raw.columns[len(expected_cols):])
                     st.session_state.processed_data = process_data(df_raw, tax_rate, keywords)
                else:
                    st.error("ساختار فایل اکسل استاندارد نیست.")
            elif st.session_state.is_demo_data:
                # استفاده از دادهساز برای نمایش قابلیت‌ها
                df_dummy = generate_dummy_data()
                st.session_state.processed_data = process_data(df_dummy, tax_rate, keywords)

        except Exception as e:
             st.error(f"خطا در پردازش: {e}")

    # ==================== Tab 1: پیشخوان (Dashboard) ====================
    with tab_dashboard:
        report = st.session_state.processed_data
        if not report.empty:
            # --- KPI Section ---
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            with kpi1:
                st.metric("کل درآمد دوره", f"{report['کل درآمد ناخالص'].sum():,.0f}", delta="ناخالص")
                st.markdown('<p class="subtle-text">مجموع واریزی‌های شناسایی شده (کارتخوان + اسنپ)</p>', unsafe_allow_html=True)
            with kpi2:
                st.metric("سود عملیاتی", f"{report['سود عملیاتی'].sum():,.0f}", delta_color="normal")
                st.markdown('<p class="subtle-text">درآمد پس از کسر مالیات و کارمزدها</p>', unsafe_allow_html=True)
            with kpi3:
                st.metric("مالیات برآوردی", f"{report['مالیات برآوردی'].sum():,.0f}", delta="-کسورات", delta_color="inverse")
                st.markdown(f'<p class="subtle-text">محاسبه شده با نرخ {tax_rate}% از فروش کارتخوان</p>', unsafe_allow_html=True)
            with kpi4:
                 # آخرین مانده حساب
                 last_balance = report.iloc[-1]['مانده نهایی'] if not report.empty else 0
                 st.metric("آخرین موجودی حساب", f"{last_balance:,.0f}")
                 st.markdown('<p class="subtle-text">مانده نهایی در آخرین روز گزارش</p>', unsafe_allow_html=True)

            st.markdown("---")

            # --- Charts Section (Modern Grid) ---
            chart_col1, chart_col2 = st.columns([2, 1], gap="large")
            
            with chart_col1:
                st.markdown("##### 📈 روند نقدینگی و درآمد")
                chart_data = report[['تاریخ', 'کل درآمد ناخالص', 'مانده نهایی']].set_index('تاریخ')
                st.line_chart(chart_data, color=["#3B82F6", "#10B981"], height=320) # رنگ‌های مدرن آبی و سبز
                st.markdown('<p class="subtle-text">مقایسه جریان ورودی روزانه با مانده کل حساب</p>', unsafe_allow_html=True)

            with chart_col2:
                 st.markdown("##### 🍩 ترکیب منابع درآمد")
                 source_df = pd.DataFrame({
                     'منبع': ['کارتخوان', 'اسنپ'],
                     'مبلغ': [report['کارتخوان'].sum(), report['اسنپ'].sum()]
                 }).set_index('منبع')
                 st.bar_chart(source_df, color="#F59E0B", height=320) # رنگ زرد/نارنجی مدرن
                 st.markdown('<p class="subtle-text">سهم هر درگاه در کل درآمد دوره</p>', unsafe_allow_html=True)
        else:
            st.warning("داده‌ای برای نمایش وجود ندارد. لطفاً به تب تنظیمات بروید.")

    # ==================== Tab 2: جزئیات تراکنش‌ها ====================
    with tab_details:
        report = st.session_state.processed_data
        if not report.empty:
            st.markdown("### 📄 جدول ریز محاسبات روزانه")
            
            # دانلود باکس مدرن
            dl_col1, dl_col2 = st.columns([3, 1])
            with dl_col2:
                excel_data = create_excel_report(report)
                st.download_button(
                    label="📥 دانلود گزارش اکسل",
                    data=excel_data,
                    file_name=f"Financial_Report_Modern_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # نمایش جدول با کانفیگ مدرن
            st.dataframe(
                report,
                use_container_width=True,
                height=500,
                hide_index=True,
                column_config={
                    "تاریخ": st.column_config.DatetimeColumn("تاریخ", format="YYYY-MM-DD"),
                    "کارتخوان": st.column_config.NumberColumn(format="%.0f ﷼"),
                    "اسنپ": st.column_config.NumberColumn(format="%.0f ﷼"),
                    "کل درآمد ناخالص": st.column_config.NumberColumn(format="%.0f ﷼"),
                    "مالیات برآوردی": st.column_config.NumberColumn(format="%.0f ﷼"),
                    "کارمزدها": st.column_config.NumberColumn(format="%.0f ﷼"),
                    "سود عملیاتی": st.column_config.NumberColumn(format="%.0f ﷼", help="درآمد خالص نهایی"),
                    "مانده نهایی": st.column_config.NumberColumn(format="%.0f ﷼"),
                }
            )
        else:
             st.info("برای مشاهده جزئیات، ابتدا داده‌ها را در تب تنظیمات بارگذاری یا ایجاد کنید.")

if __name__ == "__main__":
    main()
