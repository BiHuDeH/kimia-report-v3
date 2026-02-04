import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
import jdatetime
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- تنظیمات صفحه ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- توابع مبدل فارسی ---
def to_persian_num(number):
    """تبدیل اعداد انگلیسی به فارسی"""
    if pd.isna(number): return ""
    number = str(number)
    mapping = {
        '0': '۰', '1': '۱', '2': '۲', '3': '۳', '4': '۴',
        '5': '۵', '6': '۶', '7': '۷', '8': '۸', '9': '۹',
        '.': '/', ',': '،', '-': '−'
    }
    for k, v in mapping.items():
        number = number.replace(k, v)
    return number

def convert_to_jalali(date_obj):
    try:
        jalali_date = jdatetime.date.fromgregorian(date=date_obj.date())
        return jalali_date.strftime("%Y/%m/%d")
    except:
        return str(date_obj)

# --- CSS پیشرفته و انیمیشن‌ها ---
st.markdown("""
    <style>
        /* فونت و تنظیمات کلی */
        @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700&display=swap');
        
        * {
            font-family: 'Vazirmatn', sans-serif !important;
            direction: rtl;
        }
        
        .stApp {
            background-color: #050511; /* مشکی خیلی عمیق */
            background-image: radial-gradient(circle at 50% 0%, #1a1a40 0%, #050511 70%);
        }

        /* --- Glassmorphism Cards --- */
        .glass-card {
            background: rgba(255, 255, 255, 0.03);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 16px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        .glass-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 0 20px rgba(59, 130, 246, 0.3); /* درخشش آبی */
            border-color: rgba(59, 130, 246, 0.5);
        }

        /* --- Glowing Border Effect --- */
        @keyframes border-flow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .glowing-btn {
            position: relative;
            background: rgba(255,255,255,0.05);
            border-radius: 12px;
            z-index: 1;
            overflow: hidden;
            padding: 2px; /* ضخامت خط دور */
            cursor: pointer;
        }
        .glowing-btn::before {
            content: '';
            position: absolute;
            top: -50%; left: -50%; width: 200%; height: 200%;
            background: linear-gradient(60deg, transparent, #3B82F6, #8B5CF6, transparent);
            background-size: 200% 200%;
            animation: border-flow 3s linear infinite;
            z-index: -1;
        }
        .glowing-content {
            background: #0F172A;
            border-radius: 10px;
            height: 100%;
            width: 100%;
            display: flex;
            align-items: center;
            justify_content: center;
            padding: 10px;
        }

        /* --- Custom Navbar --- */
        .nav-container {
            display: flex;
            justify-content: center;
            gap: 15px;
            margin-bottom: 30px;
            padding: 10px;
            background: rgba(255,255,255,0.02);
            border-radius: 20px;
        }

        /* --- رفع باگ آپلودر و فارسی سازی متن --- */
        [data-testid='stFileUploader'] {
            border: 1px dashed rgba(255,255,255,0.2);
            border-radius: 15px;
            padding: 20px;
            text-align: center;
            background: rgba(255,255,255,0.02);
        }
        [data-testid='stFileUploader'] section > input + div {
            display: none; /* مخفی کردن متن انگلیسی */
        }
        [data-testid='stFileUploader'] section::after {
            content: "📂 فایل اکسل را اینجا رها کنید یا کلیک نمایید";
            color: #94A3B8;
            font-size: 0.9rem;
            display: block;
            margin-top: -30px;
            pointer-events: none;
        }
        [data-testid='stFileUploader'] button {
            display: none;
        }

        /* --- رفع باگ متن قاطی شده (Input Label) --- */
        .stTextInput label, .stNumberInput label {
            text-align: right !important;
            font-size: 0.8rem !important;
            color: #94A3B8 !important;
            margin-bottom: 5px;
        }

        /* --- تایپوگرافی ظریف --- */
        h1, h2, h3 {
            font-weight: 200 !important; /* لایت */
            color: #F8FAFC;
        }
        p, span, div {
            font-weight: 300;
        }
        
        /* مخفی کردن المان‌های مزاحم */
        #MainMenu, footer, header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- توابع منطقی و گرافیک ---

@st.cache_data
def generate_dummy_data():
    dates = pd.date_range(end=datetime.today(), periods=10)
    data = []
    balance = 50000000
    for date in dates:
        deposit = np.random.randint(1000000, 50000000)
        withdrawal = np.random.randint(100000, 5000000)
        balance += (deposit - withdrawal)
        data.append({
            'Date': date, 'Description': "تراکنش نمونه",
            'Deposit': deposit, 'Withdrawal': withdrawal, 'Balance': balance
        })
    return pd.DataFrame(data)

@st.cache_data(show_spinner=False)
def process_data(df, tax_rate, keywords):
    if 'Date' not in df.columns: return pd.DataFrame()
    df = df.dropna(subset=['Date']).copy()
    df['Jalali_Date'] = df['Date'].apply(convert_to_jalali)
    
    for col in ['Deposit', 'Withdrawal', 'Balance']:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['Description'] = df['Description'].astype(str)
    
    # فیلترها
    card = df[df['Description'].str.contains(keywords.get('card', ''), na=False)].groupby('Date')['Deposit'].sum()
    snap = df[df['Description'].str.contains(keywords.get('snap', ''), na=False)].groupby('Date')['Deposit'].sum()
    fee = df[df['Description'].str.contains(keywords.get('fee', ''), na=False)].groupby('Date')['Withdrawal'].sum()
    bal = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last()
    
    unique_dates = df['Date'].sort_values().unique()
    report = pd.DataFrame(index=unique_dates)
    report['کارتخوان'] = card.reindex(unique_dates, fill_value=0)
    report['اسنپ'] = snap.reindex(unique_dates, fill_value=0)
    report['کارمزد'] = fee.reindex(unique_dates, fill_value=0)
    report['مانده'] = bal.reindex(unique_dates, fill_value=0)
    
    date_map = df.set_index('Date')['Jalali_Date'].to_dict()
    report['تاریخ'] = report.index.map(date_map)
    
    report['درآمد کل'] = report['کارتخوان'] + report['اسنپ']
    report['مالیات'] = report['کارتخوان'] - (report['کارتخوان'] / (1 + tax_rate/100))
    report['سود خالص'] = report['درآمد کل'] - report['مالیات'] - report['کارمزد']
    
    return report.reset_index(drop=True)

def plot_persian_chart(df, x_col, y_cols, title, type='line'):
    """رسم نمودار با اعداد و متن کاملاً فارسی با Plotly"""
    fig = go.Figure()
    
    colors = ['#10B981', '#3B82F6', '#F59E0B', '#EF4444']
    
    for idx, col in enumerate(y_cols):
        y_data = df[col]
        # فرمت تولتیپ فارسی
        hover_text = [f"{to_persian_num('{:,.0f}'.format(v))} ریال" for v in y_data]
        
        if type == 'line':
            fig.add_trace(go.Scatter(
                x=df[x_col], y=y_data, mode='lines+markers', name=col,
                line=dict(width=3, color=colors[idx % len(colors)]),
                marker=dict(size=6, line=dict(width=2, color='#050511')),
                text=hover_text, hoverinfo='text+name'
            ))
        else:
            fig.add_trace(go.Bar(
                x=df[x_col], y=y_data, name=col,
                marker_color=colors[idx % len(colors)],
                text=hover_text, hoverinfo='text+name'
            ))

    # تنظیمات محورها برای نمایش فارسی
    fig.update_layout(
        title=dict(text=title, font=dict(family="Vazirmatn", size=18, color="#F8FAFC")),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Vazirmatn", color="#94A3B8"),
        hovermode="x unified",
        margin=dict(l=20, r=20, t=50, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(showgrid=False, color="#64748B"),
        yaxis=dict(showgrid=True, gridcolor="rgba(255,255,255,0.05)", color="#64748B")
    )
    return fig

# --- مدیریت وضعیت (Session State) ---
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = 'home'
if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame()
if 'is_demo' not in st.session_state:
    st.session_state.is_demo = True

# --- هدر و ناوبری ---
def render_header():
    col1, col2 = st.columns([1, 4])
    with col2:
        st.markdown("""
            <div style="text-align: right; padding-right: 20px;">
                <h1 style="font-size: 2.5rem; margin-bottom: 0; background: linear-gradient(to left, #22d3ee, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
                    داشبورد مالی کیمیا
                </h1>
                <p style="color: #64748B; font-size: 0.9rem;">سیستم هوشمند تحلیل تراکنش و محاسبه سود و زیان</p>
            </div>
        """, unsafe_allow_html=True)
    with col1:
        # آیکون یا لوگوی متحرک
        st.markdown("""
            <div style="display: flex; justify-content: center; align-items: center; height: 100%;">
                <div style="width: 60px; height: 60px; background: linear-gradient(135deg, #3B82F6, #EC4899); border-radius: 50%; display: flex; align-items: center; justify-content: center; box-shadow: 0 0 20px rgba(59, 130, 246, 0.5);">
                    <span style="font-size: 30px;">💎</span>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")

def change_tab(tab_name):
    st.session_state.active_tab = tab_name

# --- محتوای اصلی ---

def main():
    render_header()

    # --- بخش معرفی قابلیت‌ها (Hero Features) ---
    if st.session_state.active_tab == 'home':
        st.markdown("### 🚀 امکانات و تحلیل‌های موجود")
        f1, f2, f3 = st.columns(3)
        
        features = [
            (f1, "📊 تحلیل روند فروش", "بررسی نموداری صعود یا نزول درآمد در طول ماه", "مشاهده نمودار", "charts"),
            (f2, "🧾 محاسبه دقیق مالیات", "تفکیک خودکار مالیات و کارمزد از درآمد خالص", "تنظیمات و محاسبه", "settings"),
            (f3, "🍰 مقایسه منابع درآمد", "سهم کارتخوان در برابر اسنپ و سایر واریزی‌ها", "ریز تراکنش‌ها", "details"),
        ]
        
        for col, title, desc, btn_text, target in features:
            with col:
                st.markdown(f"""
                <div class="glass-card" style="text-align: center; height: 200px;">
                    <h3 style="color: #38bdf8; margin-bottom: 10px;">{title}</h3>
                    <p style="font-size: 0.85rem; color: #cbd5e1; height: 60px;">{desc}</p>
                </div>
                """, unsafe_allow_html=True)
                if st.button(f"{btn_text} 👈", key=f"btn_{target}", use_container_width=True):
                    change_tab(target)
                    st.rerun()

    # --- نوار دسترسی سریع (Custom Tabs) ---
    st.markdown('<div class="nav-container">', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("🏠 خانه", use_container_width=True): change_tab('home'); st.rerun()
    with c2:
        if st.button("⚙️ تنظیمات و فایل", use_container_width=True): change_tab('settings'); st.rerun()
    with c3:
        if st.button("📈 نمودارها", use_container_width=True): change_tab('charts'); st.rerun()
    with c4:
        if st.button("📋 جدول جزئیات", use_container_width=True): change_tab('details'); st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # ==================== تب تنظیمات ====================
    if st.session_state.active_tab == 'settings':
        col_main, col_side = st.columns([2, 1])
        
        with col_main:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 📂 بارگذاری فایل")
            uploaded_file = st.file_uploader("", type=['xlsx'])
            if uploaded_file:
                st.session_state.is_demo = False
                st.success("فایل با موفقیت دریافت شد.")
            else:
                st.session_state.is_demo = True
                st.info("حالت دمو: نمایش داده‌های تصادفی")
            st.markdown('</div>', unsafe_allow_html=True)

        with col_side:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 🎛 پارامترها")
            tax_rate = st.slider("نرخ مالیات (%)", 0, 20, 10)
            
            with st.expander("🔍 کلمات کلیدی (فیلترها)"):
                k_card = st.text_input("کارتخوان", "انتقال از")
                k_snap = st.text_input("اسنپ", "مدرن سامانه")
                k_fee = st.text_input("کارمزد", "کارمزد")
            st.markdown('</div>', unsafe_allow_html=True)
            
        # پردازش
        try:
            keywords = {'card': k_card, 'snap': k_snap, 'fee': k_fee}
            if uploaded_file:
                raw = pd.read_excel(uploaded_file, skiprows=2)
                # Simple standardization
                if len(raw.columns) > 10:
                    expected = ['Index', 'Branch', 'Code', 'Date', 'Time', 'Doc', 'Rec', 'Chk', 'Description', 'Withdrawal', 'Deposit', 'Balance', 'Note']
                    raw.columns = expected + list(raw.columns[len(expected):])
                    st.session_state.data = process_data(raw, tax_rate, keywords)
            elif st.session_state.is_demo:
                raw = generate_dummy_data()
                st.session_state.data = process_data(raw, tax_rate, keywords)
        except Exception as e:
            st.error(f"خطا: {e}")

    # ==================== تب نمودارها (Dashboard) ====================
    elif st.session_state.active_tab == 'charts':
        df = st.session_state.data
        if not df.empty:
            # KPI Cards
            k1, k2, k3, k4 = st.columns(4)
            metrics = [
                ("درآمد کل", df['درآمد کل'].sum(), "#10B981"),
                ("سود خالص", df['سود خالص'].sum(), "#3B82F6"),
                ("مالیات", df['مالیات'].sum(), "#EF4444"),
                ("میانگین روزانه", df['درآمد کل'].mean(), "#F59E0B"),
            ]
            
            for col, (title, val, color) in zip([k1,k2,k3,k4], metrics):
                with col:
                    st.markdown(f"""
                    <div class="glass-card" style="border-top: 3px solid {color}; text-align: center; padding: 15px;">
                        <span style="color: #94A3B8; font-size: 0.8rem;">{title}</span>
                        <h2 style="color: {color}; margin: 5px 0; font-size: 1.5rem;">{to_persian_num(f"{val:,.0f}")}</h2>
                        <span style="font-size: 0.7rem; color: #64748B;">ریال</span>
                    </div>
                    """, unsafe_allow_html=True)

            # Charts Row 1
            c1, c2 = st.columns([2, 1])
            with c1:
                st.markdown('<div class="glass-card">', unsafe_allow_html=True)
                fig_line = plot_persian_chart(df, 'تاریخ', ['درآمد کل', 'مانده'], "روند نقدینگی و درآمد")
                st.plotly_chart(fig_line, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with c2:
                st.markdown('<div class="glass-card">', unsafe_allow_html=True)
                fig_bar = plot_persian_chart(df, 'تاریخ', ['کارتخوان', 'اسنپ'], "مقایسه منابع", type='bar')
                st.plotly_chart(fig_bar, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
        else:
            st.warning("لطفاً ابتدا به بخش تنظیمات بروید و فایل را بارگذاری کنید.")

    # ==================== تب جزئیات ====================
    elif st.session_state.active_tab == 'details':
        df = st.session_state.data
        if not df.empty:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown("#### 📋 ریز تراکنش‌های محاسبه شده")
            
            # نمایش جدول با اعداد فارسی
            display_df = df.copy()
            for col in display_df.select_dtypes(include=np.number).columns:
                display_df[col] = display_df[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))
            
            st.dataframe(display_df, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        else:
             st.info("داده‌ای موجود نیست.")

if __name__ == "__main__":
    main()
