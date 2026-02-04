import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
import jdatetime
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- تنظیمات صفحه (باید اولین خط باشد) ---
st.set_page_config(
    page_title="مدیریت مالی",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- توابع مبدل (فارسی‌ساز) ---

def to_persian_num(number):
    """تبدیل اعداد انگلیسی به فارسی"""
    if pd.isna(number): return ""
    number = str(number)
    mapping = {
        '0': '۰', '1': '۱', '2': '۲', '3': '۳', '4': '۴',
        '5': '۵', '6': '۶', '7': '۷', '8': '۸', '9': '۹',
        '.': '/', ',': '،'
    }
    for k, v in mapping.items():
        number = number.replace(k, v)
    return number

def convert_to_jalali(date_obj):
    """تبدیل تاریخ میلادی به شمسی (۱۴۰۲/۰۱/۰۱)"""
    if pd.isna(date_obj): return ""
    try:
        jalali_date = jdatetime.date.fromgregorian(date=date_obj.date())
        return jalali_date.strftime("%Y/%m/%d")
    except:
        return str(date_obj)

# --- استایل‌دهی CSS (تم تیره اپلیکیشنی) ---
st.markdown("""
    <style>
        /* ایمپورت فونت وزیرمتن */
        @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');

        /* تنظیمات کلی بدنه */
        .stApp {
            background-color: #0F172A; /* رنگ پس‌زمینه تیره */
            font-family: 'Vazirmatn', sans-serif !important;
        }
        
        * {
            font-family: 'Vazirmatn', sans-serif !important;
            direction: rtl;
        }

        /* مخفی کردن منوی پیش‌فرض استریم‌لیت */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* --- طراحی کارت‌ها (شبیه تصویر ارسالی) --- */
        .kpi-card {
            background: linear-gradient(145deg, #1E293B, #0F172A);
            border-radius: 20px;
            padding: 20px;
            margin-bottom: 15px;
            border: 1px solid #334155;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.5);
            text-align: center;
            transition: transform 0.2s;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .kpi-card:hover {
            transform: translateY(-5px);
            border-color: #3B82F6;
        }
        .kpi-title {
            color: #94A3B8;
            font-size: 0.9rem;
            margin-bottom: 10px;
            font-weight: 300;
        }
        .kpi-value {
            color: #F8FAFC;
            font-size: 1.6rem;
            font-weight: 800;
            margin: 0;
            direction: ltr; /* برای نمایش صحیح اعداد منفی */
        }
        .kpi-icon {
            font-size: 1.5rem;
            margin-bottom: 10px;
        }

        /* --- هدر مدرن --- */
        .app-header {
            text-align: center;
            padding: 2rem 0 1rem 0;
            margin-bottom: 2rem;
        }
        .app-title {
            background: linear-gradient(to right, #4ADE80, #3B82F6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.2rem;
            font-weight: 900;
            margin-bottom: 0.5rem;
        }
        .app-subtitle {
            color: #64748B;
            font-size: 1rem;
        }

        /* --- استایل جداول --- */
        div[data-testid="stDataFrame"] {
            direction: rtl;
            background-color: #1E293B;
            border-radius: 15px;
            padding: 10px;
        }
        
        /* دکمه‌ها */
        .stButton > button {
            border-radius: 12px;
            background-color: #3B82F6;
            color: white;
            border: none;
            width: 100%;
            padding: 10px;
            font-weight: bold;
        }
        .stButton > button:hover {
            background-color: #2563EB;
        }

        /* تب‌ها */
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            background-color: transparent;
        }
        .stTabs [data-baseweb="tab"] {
            height: 45px;
            background-color: #1E293B;
            border-radius: 10px;
            color: #94A3B8;
            border: 1px solid #334155;
            flex: 1; /* هم‌اندازه کردن تب‌ها */
        }
        .stTabs [aria-selected="true"] {
            background-color: #3B82F6 !important;
            color: white !important;
            border-color: #3B82F6 !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- توابع منطقی ---

@st.cache_data
def generate_dummy_data():
    """تولید داده‌های نمونه برای نمایش اولیه"""
    dates = pd.date_range(end=datetime.today(), periods=15)
    data = []
    balance = 50000000
    descriptions = ["خرید اینترنتی", "واریز اسنپ فود", "انتقال وجه پایا", "کارمزد بانکی", "کارتخوان فروشگاه"]
    
    for date in dates:
        desc = np.random.choice(descriptions)
        amount = np.random.randint(100000, 20000000)
        dep = 0
        wit = 0
        if "واریز" in desc or "کارتخوان" in desc:
            dep = amount
            balance += amount
        else:
            wit = amount
            balance -= amount
            
        data.append({
            'Date': date,
            'Time': datetime.now().time(),
            'Description': desc,
            'Withdrawal': wit,
            'Deposit': dep,
            'Balance': balance
        })
    return pd.DataFrame(data)

@st.cache_data(show_spinner=False)
def process_data(df, tax_rate, keywords):
    """پردازش هسته داده‌ها"""
    if 'Date' not in df.columns: return pd.DataFrame()
    
    # کپی و حذف سطرهای بدون تاریخ
    df = df.dropna(subset=['Date']).copy()
    
    # تبدیل تاریخ به شمسی (ستون جدید)
    df['Jalali_Date'] = df['Date'].apply(convert_to_jalali)
    
    unique_dates = df['Date'].sort_values().unique()

    # تبدیل اعداد
    for col in ['Deposit', 'Withdrawal', 'Balance']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Description'] = df['Description'].astype(str)

    # فیلترها
    card_mask = df['Description'].str.contains(keywords.get('card', ''), na=False)
    snap_mask = df['Description'].str.contains(keywords.get('snap', ''), na=False)
    fee_mask = df['Description'].str.contains(keywords.get('fee', ''), na=False)

    # تجمیع بر اساس تاریخ میلادی (برای سورت درست) اما نمایش شمسی
    grouped = df.groupby('Date')
    
    report = pd.DataFrame(index=unique_dates)
    report['کارتخوان'] = df[card_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    report['اسنپ'] = df[snap_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    report['کارمزد'] = df[fee_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
    report['مانده'] = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last().reindex(unique_dates, fill_value=0)
    
    # اضافه کردن ستون تاریخ شمسی به گزارش نهایی
    # مپ کردن تاریخ میلادی ایندکس به تاریخ شمسی
    date_mapping = df.set_index('Date')['Jalali_Date'].to_dict()
    report['تاریخ'] = report.index.map(date_mapping)

    # محاسبات
    tax_factor = 1 + (tax_rate / 100)
    report['درآمد کل'] = report['کارتخوان'] + report['اسنپ']
    report['فروش خالص'] = report['کارتخوان'] / tax_factor
    report['مالیات'] = report['کارتخوان'] - report['فروش خالص']
    
    # بازنشانی ایندکس
    report = report.reset_index(drop=True)
    
    # ترتیب ستون‌ها
    final_cols = ['تاریخ', 'درآمد کل', 'کارتخوان', 'اسنپ', 'مالیات', 'کارمزد', 'مانده']
    return report[final_cols]

def create_excel(df):
    """ایجاد اکسل خروجی"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    ws.title = "گزارش مالی"
    
    # نوشتن هدر و داده‌ها (کد اکسل مشابه قبل ولی خلاصه شده)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(row)
        
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- بدنه اصلی اپلیکیشن ---

def main():
    # هدر انیمیشنی
    st.markdown("""
        <div class="app-header">
            <div class="app-title">داشبورد مالی کیمیا</div>
            <div class="app-subtitle">مدیریت هوشمند تراکنش‌های بانکی</div>
        </div>
    """, unsafe_allow_html=True)

    # --- مدیریت وضعیت (State) ---
    if 'data' not in st.session_state:
        st.session_state.data = pd.DataFrame()
    if 'is_demo' not in st.session_state:
        st.session_state.is_demo = True

    # --- تب‌بندی ---
    tab1, tab2, tab3 = st.tabs(["📊 نمای کلی", "📋 ریز تراکنش‌ها", "⚙️ تنظیمات"])

    # === تب تنظیمات (ابتدا پردازش می‌شود) ===
    with tab3:
        st.markdown("### بارگذاری و تنظیمات")
        col_up, col_set = st.columns(2)
        
        with col_up:
            uploaded_file = st.file_uploader("انتخاب فایل اکسل", type=['xlsx'])
            if uploaded_file:
                st.session_state.is_demo = False
            else:
                st.session_state.is_demo = True
        
        with col_set:
            tax_rate = st.slider("نرخ مالیات (%)", 0, 20, 10)
            with st.expander("تعریف کلمات کلیدی"):
                k_card = st.text_input("کارتخوان", "انتقال از")
                k_snap = st.text_input("اسنپ", "مدرن سامانه")
                k_fee = st.text_input("کارمزد", "کارمزد")
                keywords = {'card': k_card, 'snap': k_snap, 'fee': k_fee}

        # پردازش
        try:
            if uploaded_file:
                raw = pd.read_excel(uploaded_file, skiprows=2)
                # استانداردسازی ستون‌ها
                expected = ['Index', 'Branch', 'Code', 'Date', 'Time', 'Doc', 'Rec', 'Chk', 'Description', 'Withdrawal', 'Deposit', 'Balance', 'Note']
                if len(raw.columns) >= len(expected):
                    raw.columns = expected + list(raw.columns[len(expected):])
                    st.session_state.data = process_data(raw, tax_rate, keywords)
            elif st.session_state.is_demo:
                raw = generate_dummy_data()
                st.session_state.data = process_data(raw, tax_rate, keywords)
                
        except Exception as e:
            st.error(f"خطا: {e}")

    # === تب نمای کلی (داشبورد) ===
    with tab1:
        df = st.session_state.data
        if not df.empty:
            # محاسبه سرجمع‌ها
            total_income = df['درآمد کل'].sum()
            total_tax = df['مالیات'].sum()
            total_fee = df['کارمزد'].sum()
            last_bal = df.iloc[-1]['مانده']

            # ردیف اول: کارت‌های وضعیت (KPI)
            # استفاده از HTML سفارشی برای زیبایی و اعداد فارسی
            c1, c2, c3, c4 = st.columns(4)
            
            cards = [
                (c1, "💰", "درآمد کل", total_income, "#10B981"),
                (c2, "🏦", "موجودی حساب", last_bal, "#3B82F6"),
                (c3, "📉", "مالیات برآوردی", total_tax, "#EF4444"),
                (c4, "💸", "هزینه کارمزد", total_fee, "#F59E0B"),
            ]

            for col, icon, title, val, color in cards:
                val_persian = to_persian_num(f"{val:,.0f}")
                with col:
                    st.markdown(f"""
                        <div class="kpi-card">
                            <div class="kpi-icon">{icon}</div>
                            <div class="kpi-title">{title}</div>
                            <div class="kpi-value" style="color: {color};">{val_persian}</div>
                        </div>
                    """, unsafe_allow_html=True)

            st.markdown("---")

            # ردیف دوم: نمودارها
            g1, g2 = st.columns([2, 1])
            
            with g1:
                st.markdown("##### 📈 روند نقدینگی (ماهانه)")
                # برای نمودار باید ایندکس تاریخ باشد
                chart_data = df.set_index('تاریخ')[['درآمد کل', 'مانده']]
                st.line_chart(chart_data, color=["#10B981", "#3B82F6"], height=300)

            with g2:
                st.markdown("##### 🍰 ترکیب درآمد")
                src_data = pd.DataFrame({
                    'مبلغ': [df['کارتخوان'].sum(), df['اسنپ'].sum()],
                    'منبع': ['کارتخوان', 'اسنپ']
                }).set_index('منبع')
                st.bar_chart(src_data, color="#8B5CF6", height=300)

        else:
            st.info("داده‌ای موجود نیست.")

    # === تب ریز تراکنش‌ها ===
    with tab2:
        df = st.session_state.data
        if not df.empty:
            
            # آماده‌سازی دیتافریم برای نمایش (تبدیل همه اعداد به رشته فارسی)
            display_df = df.copy()
            numeric_cols = ['درآمد کل', 'کارتخوان', 'اسنپ', 'مالیات', 'کارمزد', 'مانده']
            for col in numeric_cols:
                display_df[col] = display_df[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))
            
            st.dataframe(display_df, use_container_width=True, height=500)
            
            # دانلود
            xlsx = create_excel(df)
            st.download_button("📥 دانلود فایل اکسل", xlsx, "Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
