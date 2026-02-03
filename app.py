import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- 1. تنظیمات صفحه و استایل ---
st.set_page_config(page_title="داشبورد مالی هوشمند", page_icon="💰", layout="wide")

# استایل CSS برای راست‌چین کردن و زیباسازی فونت‌ها
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;700&display=swap');
    
    * {
        font-family: 'Tahoma', 'Vazirmatn', sans-serif !important;
    }
    .stApp {
        background-color: #f8f9fa;
    }
    .metric-card {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        text-align: center;
        border-top: 3px solid #0068c9;
    }
    h1, h2, h3 {
        color: #2c3e50;
        text-align: right;
    }
    .stDataFrame {
        direction: rtl;
    }
    /* تنظیمات جدول برای نمایش بهتر */
    div[data-testid="stMetricValue"] {
        font-size: 1.2rem !important;
        color: #0068c9;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. توابع کمکی ---

def create_excel_report(df):
    """تولید فایل اکسل فرمت‌دهی شده برای دانلود"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش مالی"
    ws.sheet_view.rightToLeft = True

    # استایل‌ها
    header_font = Font(bold=True, size=11, name='Tahoma', color='FFFFFF')
    header_fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
    regular_font = Font(size=10, name='Tahoma')
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # ستون‌ها
    cols = ['تاریخ', 'کارت به کارت', 'فروش خالص', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده نهایی', 'واریزی اسنپ', 'کل درآمد']
    
    # اطمینان از وجود ستون‌ها
    for c in cols:
        if c not in df.columns: df[c] = 0
            
    ws.append(cols)

    # استایل دهی هدر
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # نوشتن داده‌ها
    for row in df[cols].itertuples(index=False, name=None):
        ws.append(row)

    # فرمت‌دهی سلول‌ها
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = regular_font
            cell.alignment = center_align
            cell.border = thin_border
            if cell.column_letter != 'A': # ستون تاریخ فرمت عدد نگیرد
                cell.number_format = '#,##0'

    # تنظیم عرض ستون‌ها
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 16

    # تبدیل به جدول اکسل
    if ws.max_row >= 2:
        tab = Table(displayName="FinancialTable", ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tab)

    wb.save(output)
    output.seek(0)
    return output.getvalue()

@st.cache_data(show_spinner=False)
def process_data(df, tax_rate_percent, keywords):
    """پردازش داده‌ها با پارامترهای پویا"""
    if 'Date' not in df.columns: return pd.DataFrame()

    df = df.dropna(subset=['Date'])
    unique_dates = df['Date'].sort_values().unique()

    # تبدیل به عدد
    for col in ['Deposit', 'Withdrawal', 'Balance']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # فیلترینگ هوشمند
    df['Description'] = df['Description'].astype(str)
    
    # استفاده از تنظیمات کاربر
    card_mask = df['Description'].str.contains(keywords['card'], na=False)
    fee_mask = df['Description'].str.contains(keywords['fee'], na=False)
    withdraw_mask = df['Description'].str.contains(keywords['withdraw'], na=False)
    snap_mask = df['Description'].str.contains(keywords['snap'], na=False)

    # تجمیع داده‌ها
    agg_funcs = {
        'کارت به کارت': df[card_mask].groupby('Date')['Deposit'].sum(),
        'کارمزد': df[fee_mask].groupby('Date')['Withdrawal'].sum(),
        'برداشت روز': df[withdraw_mask].groupby('Date')['Withdrawal'].sum(),
        'واریزی اسنپ': df[snap_mask].groupby('Date')['Deposit'].sum(),
        'مانده نهایی': df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last()
    }

    report = pd.DataFrame(index=unique_dates)
    for name, series in agg_funcs.items():
        report[name] = series.reindex(unique_dates, fill_value=0)

    # محاسبات پیشرفته
    tax_factor = 1 + (tax_rate_percent / 100)
    report['فروش خالص'] = report['کارت به کارت'] / tax_factor
    report['مالیات'] = report['کارت به کارت'] - report['فروش خالص']
    
    # ستون جدید: کل درآمد (شامل اسنپ و کارتخوان)
    report['کل درآمد'] = report['کارت به کارت'] + report['واریزی اسنپ']

    return report.reset_index().rename(columns={'index': 'تاریخ'})

# --- 3. رابط کاربری اصلی ---

def main():
    # سایدبار: تنظیمات و ورودی
    with st.sidebar:
        st.header("⚙️ تنظیمات پردازش")
        
        uploaded_file = st.file_uploader("آپلود فایل اکسل", type=["xlsx"])
        
        st.markdown("---")
        st.subheader("📊 تنظیمات مالی")
        tax_rate = st.slider("نرخ مالیات (درصد)", 0, 20, 10, help="برای محاسبه فروش خالص از کل واریزی کارتخوان")
        
        with st.expander("🔍 تنظیمات فیلترها (پیشرفته)"):
            k_card = st.text_input("کلیدواژه کارتخوان", "انتقال از")
            k_snap = st.text_input("کلیدواژه اسنپ", "مدرن سامانه")
            k_fee = st.text_input("کلیدواژه کارمزد", "کارمزد")
            k_withdraw = st.text_input("کلیدواژه برداشت", "انتقال وجه")
            
            keywords = {'card': k_card, 'snap': k_snap, 'fee': k_fee, 'withdraw': k_withdraw}

        st.info(f"نسخه: 4.0 (Intelligent UI)\nتاریخ: {datetime.now().strftime('%Y-%m-%d')}")

    # بدنه اصلی
    st.title("📊 داشبورد مدیریت مالی")
    
    if not uploaded_file:
        st.markdown("""
        <div style='text-align: right; color: #555;'>
        خوش آمدید. برای شروع، فایل اکسل تراکنش‌های بانکی خود را از منوی سمت راست آپلود کنید.
        <br>این سیستم به طور خودکار درآمدهای اسنپ، فروش کارتخوان و مالیات را تفکیک می‌کند.
        </div>
        """, unsafe_allow_html=True)
        return

    try:
        # خواندن فایل
        df_raw = pd.read_excel(uploaded_file, skiprows=2)
        
        # استانداردسازی ستون‌ها
        expected_cols = ['Index', 'Branch Code', 'Branch', 'Date', 'Time', 'Doc Num', 'Receipt', 'Check', 'Description', 'Withdrawal', 'Deposit', 'Balance', 'Notes']
        if len(df_raw.columns) >= len(expected_cols):
            df_raw.columns = expected_cols + list(df_raw.columns[len(expected_cols):])
        else:
            st.error("فرمت فایل اکسل استاندارد نیست.")
            return

        # پردازش
        with st.spinner('در حال تحلیل هوشمند داده‌ها...'):
            report = process_data(df_raw, tax_rate, keywords)

        if report.empty:
            st.warning("داده‌ای یافت نشد.")
            return

        # --- بخش 1: کارت‌های شاخص (KPI) ---
        total_sales = report['کل درآمد'].sum()
        total_tax = report['مالیات'].sum()
        total_fee = report['کارمزد'].sum()
        net_profit = report['فروش خالص'].sum() + report['واریزی اسنپ'].sum() - total_fee

        st.markdown("### 📈 نمای کلی دوره")
        col1, col2, col3, col4 = st.columns(4)
        
        col1.metric("کل درآمد (ناخالص)", f"{total_sales:,.0f}", help="جمع کارتخوان + اسنپ")
        col2.metric("مالیات برآوردی", f"{total_tax:,.0f}", delta="-کسورات", delta_color="inverse")
        col3.metric("کارمزد بانکی", f"{total_fee:,.0f}", delta="-هزینه", delta_color="inverse")
        col4.metric("درآمد خالص", f"{net_profit:,.0f}", help="درآمد منهای مالیات و کارمزد", delta_color="normal")

        st.markdown("---")

        # --- بخش 2: نمودارها ---
        c1, c2 = st.columns([2, 1])
        
        with c1:
            st.markdown("##### 📅 روند درآمد و موجودی")
            chart_data = report[['تاریخ', 'کل درآمد', 'مانده نهایی']].set_index('تاریخ')
            st.line_chart(chart_data, color=["#2ecc71", "#3498db"], height=300)
            
        with c2:
            st.markdown("##### 🍰 ترکیب درآمد")
            # آماده‌سازی داده برای نمودار دایره‌ای
            source_data = pd.DataFrame({
                'منبع': ['کارتخوان', 'اسنپ'],
                'مبلغ': [report['کارت به کارت'].sum(), report['واریزی اسنپ'].sum()]
            }).set_index('منبع')
            st.bar_chart(source_data, color="#f1c40f", height=300)

        # --- بخش 3: جدول داده‌ها ---
        st.markdown("### 📋 ریز تراکنش‌های روزانه")
        
        tab_view, tab_download = st.tabs(["مشاهده جدول", "دانلود فایل"])
        
        with tab_view:
            st.dataframe(
                report,
                use_container_width=True,
                column_config={
                    "تاریخ": st.column_config.TextColumn("تاریخ"),
                    "کل درآمد": st.column_config.NumberColumn(format="%.0f"),
                    "کارت به کارت": st.column_config.NumberColumn(format="%.0f"),
                    "واریزی اسنپ": st.column_config.NumberColumn(format="%.0f"),
                    "فروش خالص": st.column_config.NumberColumn(format="%.0f"),
                    "مالیات": st.column_config.NumberColumn(format="%.0f"),
                    "کارمزد": st.column_config.NumberColumn(format="%.0f"),
                    "مانده نهایی": st.column_config.NumberColumn(format="%.0f"),
                },
                hide_index=True
            )
            
        with tab_download:
            excel_data = create_excel_report(report)
            st.download_button(
                label="📥 دانلود گزارش کامل اکسل",
                data=excel_data,
                file_name=f"Gozaresh_Mali_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="این فایل شامل تمام محاسبات و جدول‌های مرتب شده است."
            )

    except Exception as e:
        st.error(f"خطای غیرمنتظره: {e}")

if __name__ == "__main__":
    main()