import streamlit as st
import pandas as pd
import sqlite3
import yfinance as yf
from datetime import datetime
from io import BytesIO

# =====================================================
# 鵬鵬的退休計畫系統 v10.4
# 重點：密碼保護、資料庫備份、KPI 整數顯示、配息行事曆、損益排行、Excel 匯出
# =====================================================

DB_NAME = "etf_portfolio.db"

# 1. 初始化資料庫與專用資料表
conn = sqlite3.connect(DB_NAME, check_same_thread=False)
c = conn.cursor()

c.execute('''
    CREATE TABLE IF NOT EXISTS tx_v74 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        symbol TEXT,
        amount REAL,
        shares REAL,
        fee REAL
    )
''')

c.execute('''
    CREATE TABLE IF NOT EXISTS dividend_v74 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        symbol TEXT,
        dividend_amount REAL
    )
''')

c.execute('''
    CREATE TABLE IF NOT EXISTS est_dividend_v90 (
        symbol TEXT PRIMARY KEY,
        frequency TEXT,
        cash_per_time REAL,
        stock_per_time REAL
    )
''')

c.execute('''
    CREATE TABLE IF NOT EXISTS dividend_calendar_v101 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        symbol TEXT,
        ex_dividend_date TEXT,
        payment_date TEXT,
        dividend_per_share REAL,
        note TEXT
    )
''')

c.execute('''
    CREATE TABLE IF NOT EXISTS sell_tx_v104 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        symbol TEXT,
        shares REAL,
        sell_price REAL,
        fee REAL,
        tax REAL
    )
''')

c.execute('''
    CREATE TABLE IF NOT EXISTS cash_account_v104 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        amount REAL,
        note TEXT
    )
''')

conn.commit()

# 2. 網頁佈局設定
st.set_page_config(page_title="鵬鵬的退休計畫系統 v10.4", layout="wide")

# ===== 網站密碼保護 =====
def check_password():
    """使用 Streamlit Secrets 的簡易密碼保護。"""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("🔒 鵬鵬的退休計畫系統")
    st.caption("請輸入密碼後再進入系統。")

    try:
        app_password = st.secrets.get("APP_PASSWORD", None)
    except Exception:
        app_password = None

    if not app_password:
        st.error("尚未設定 APP_PASSWORD。請到 Streamlit Cloud → App settings → Secrets 新增 APP_PASSWORD。")
        st.code('APP_PASSWORD = "請換成你的密碼"', language="toml")
        return False

    password = st.text_input("請輸入存取密碼", type="password")

    if st.button("登入"):
        if password == app_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("密碼錯誤，請再試一次。")

    return False


if not check_password():
    st.stop()

# KPI 卡片美化
st.markdown("""
<style>
[data-testid="metric-container"]{
    background-color:#ffffff;
    border:1px solid #e5e7eb;
    padding:18px;
    border-radius:18px;
    box-shadow:0 2px 8px rgba(0,0,0,0.06);
}
[data-testid="metric-container"] label{
    font-size:15px;
}
[data-testid="metric-container"] div{
    font-weight:700;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
.main .block-container{
    padding-top:1rem;
    max-width:1400px;
}
h1,h2,h3{
    font-weight:800 !important;
}
div[data-testid="stDataFrame"]{
    border-radius:16px;
    overflow:hidden;
    border:1px solid #e5e7eb;
}
[data-testid="metric-container"]{
    background:linear-gradient(180deg,#ffffff,#f8fafc);
    border:1px solid #e5e7eb;
    border-radius:18px;
    padding:18px;
    box-shadow:0 4px 12px rgba(0,0,0,.05);
}
[data-testid="metric-container"] label{
    font-weight:600;
}
.stTabs [data-baseweb="tab"]{
    font-size:16px;
    font-weight:700;
}
</style>
""", unsafe_allow_html=True)



st.markdown("""
<style>
.stApp{
    background: #f4f7fb;
}
.main .block-container{
    padding-top:0.8rem;
    max-width:1600px;
}
h1{
    font-size:2.4rem !important;
    font-weight:800 !important;
    color:#1f2937;
}
h2,h3{
    font-weight:700 !important;
}
[data-testid="metric-container"]{
    background:white !important;
    border:none !important;
    border-radius:22px !important;
    padding:22px !important;
    box-shadow:0 8px 24px rgba(0,0,0,.08) !important;
}
[data-testid="stSidebar"]{
    background:#ffffff;
    border-right:1px solid #e5e7eb;
}
div[data-testid="stDataFrame"]{
    background:white;
    border-radius:20px;
    padding:8px;
    box-shadow:0 6px 18px rgba(0,0,0,.06);
}
.stTabs [data-baseweb="tab-list"]{
    gap:10px;
}
.stTabs [data-baseweb="tab"]{
    background:white;
    border-radius:12px;
    padding:10px 18px;
    font-weight:700;
}
.stButton>button{
    border-radius:12px;
    font-weight:700;
}
</style>
""", unsafe_allow_html=True)


st.title("📊 鵬鵬的退休計畫系統")
st.caption("v10.2 備份保護版 · KPI 整數顯示 · Yahoo 批次報價快取 · 不需 Plotly · 損益排行 · Excel 匯出")

# --- 共用函式 ---
def clean_symbol(s, market_type="台灣股市 (自動補 .TW)"):
    if not s:
        return ""
    s = str(s).strip().upper()
    if s.isdigit():
        return f"{s}.TW"
    if market_type == "海外美股 / 主動型基金 (保持原樣)":
        return s
    return s if s.endswith(".TW") else f"{s}.TW"


def fmt_money0(v):
    try:
        return f"${round(float(v)):,.0f}"
    except Exception:
        return "$0"


def fmt_money2(v):
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"


def fmt_pct(v):
    try:
        return f"{float(v):+.2f}%"
    except Exception:
        return "+0.00%"


@st.cache_data(ttl=300, show_spinner=False)
def get_prices(symbols):
    """用 Yahoo Finance 一次抓取多檔報價，5 分鐘快取。"""
    symbols = [s for s in symbols if s]
    prices = {s: None for s in symbols}

    if not symbols:
        return prices

    try:
        data = yf.download(
            tickers=symbols,
            period="5d",
            interval="1d",
            progress=False,
            auto_adjust=True,
            group_by="ticker",
            threads=True,
        )

        if len(symbols) == 1:
            s = symbols[0]
            if "Close" in data.columns:
                close_series = data["Close"].dropna()
                if not close_series.empty:
                    prices[s] = float(close_series.iloc[-1])
        else:
            for s in symbols:
                try:
                    close_series = data[s]["Close"].dropna()
                    if not close_series.empty:
                        prices[s] = float(close_series.iloc[-1])
                except Exception:
                    prices[s] = None
    except Exception:
        pass

    return prices


def make_excel_bytes(display_df, raw_df, tx_df, div_df, est_df, calendar_df=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="綜合資產明細")
        raw_df.to_excel(writer, index=False, sheet_name="原始計算資料")
        tx_df.to_excel(writer, index=False, sheet_name="買入明細")
        div_df.to_excel(writer, index=False, sheet_name="配息明細")
        est_df.to_excel(writer, index=False, sheet_name="預估參數")
        if calendar_df is not None:
            calendar_df.to_excel(writer, index=False, sheet_name="配息行事曆")
    return output.getvalue()


# --- 側邊欄：資料庫備份 / 還原 ---
st.sidebar.markdown("---")
st.sidebar.subheader("🛡️ 資料備份 / 還原")

try:
    with open(DB_NAME, "rb") as db_file:
        st.sidebar.download_button(
            label="📥 下載資料庫備份",
            data=db_file.read(),
            file_name=f"pengpeng_retirement_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.db",
            mime="application/octet-stream",
            help="建議每次大量更新資料後下載一次，存到 Google Drive 或電腦備份資料夾。"
        )
except FileNotFoundError:
    st.sidebar.warning("目前尚未建立資料庫，請先輸入資料。")

st.sidebar.caption("備份檔可保存你的買入紀錄、配息紀錄、預估參數與配息行事曆。")

with st.sidebar.expander("📤 還原資料庫備份", expanded=False):
    st.warning("還原會覆蓋目前雲端資料。建議先下載一次目前備份，再執行還原。")

    uploaded_db = st.file_uploader(
        "選擇 .db 備份檔",
        type=["db"],
        key="restore_db_file"
    )

    confirm_restore = st.checkbox(
        "我確認要用上傳的備份覆蓋目前資料",
        key="confirm_restore_db"
    )

    if st.button("📤 執行還原", key="restore_db_button"):
        if uploaded_db is None:
            st.error("請先選擇一個 .db 備份檔。")
        elif not confirm_restore:
            st.error("請先勾選確認覆蓋目前資料。")
        else:
            uploaded_bytes = uploaded_db.getvalue()
            temp_restore_path = f"{DB_NAME}.restore_check"

            try:
                with open(temp_restore_path, "wb") as f:
                    f.write(uploaded_bytes)

                # 簡單驗證是否為可讀 SQLite 檔案，且至少包含系統主要資料表。
                test_conn = sqlite3.connect(temp_restore_path)
                table_names = pd.read_sql_query(
                    "SELECT name FROM sqlite_master WHERE type='table'",
                    test_conn
                )["name"].tolist()
                test_conn.close()

                required_any_tables = {"tx_v74", "dividend_v74", "est_dividend_v90"}
                if not required_any_tables.intersection(set(table_names)):
                    st.error("這個檔案不像是本系統的備份資料庫，已取消還原。")
                else:
                    conn.close()
                    with open(DB_NAME, "wb") as f:
                        f.write(uploaded_bytes)
                    st.success("資料庫已成功還原，系統即將重新整理。")
                    st.rerun()

            except Exception as e:
                st.error(f"還原失敗：{e}")
            finally:
                try:
                    import os
                    if os.path.exists(temp_restore_path):
                        os.remove(temp_restore_path)
                except Exception:
                    pass

# --- 側邊欄：後台輸入 ---
st.sidebar.header("⚙️ 系統資料輸入後台")
if st.sidebar.button("🚪 登出"):
    st.session_state["authenticated"] = False
    st.rerun()

action_type = st.sidebar.selectbox(
    "請選擇輸入類型",
    ["➕ 新增買入庫存", "💰 記錄收到配息", "🔮 設定未來預估配息率", "📅 設定未來配息公告"]
)

in_date = st.sidebar.date_input("入帳/申購日期", datetime.now())
raw_sym = st.sidebar.text_input("股票/ETF 代碼", placeholder="例如: 00712、00918、2834").strip()

if action_type == "➕ 新增買入庫存":
    mkt = st.sidebar.radio(
        "市場類型",
        ["台灣股市 (自動補 .TW)", "海外美股 / 主動型基金 (保持原樣)"],
        index=0
    )
    in_amt = st.sidebar.number_input("投入總金額 (不含手續費的純本金)", min_value=0.0, step=1000.0)
    in_shares = st.sidebar.number_input("獲得總單位數 (股)", min_value=0.0, step=10.0)
    in_fee = st.sidebar.number_input("交易手續費", min_value=0.0, step=1.0, value=0.0)

    if st.sidebar.button("確認寫入資產庫"):
        final_sym = clean_symbol(raw_sym, mkt)
        if final_sym and in_shares > 0:
            c.execute(
                "INSERT INTO tx_v74 (date, symbol, amount, shares, fee) VALUES (?, ?, ?, ?, ?)",
                (in_date.strftime('%Y-%m-%d'), final_sym, in_amt, in_shares, in_fee)
            )
            conn.commit()
            st.sidebar.success(f"🎉 庫存 {final_sym} 已成功寫入資料庫！")
            st.rerun()
        else:
            st.sidebar.error("❌ 請檢查：代碼不可為空，且股數必須大於 0！")

elif action_type == "💰 記錄收到配息":
    in_div = st.sidebar.number_input("本次收到配息總金額 (台幣)", min_value=0.0, step=100.0)

    if st.sidebar.button("確認寫入配息紀錄"):
        final_sym = clean_symbol(raw_sym)
        if final_sym and in_div > 0:
            c.execute(
                "INSERT INTO dividend_v74 (date, symbol, dividend_amount) VALUES (?, ?, ?)",
                (in_date.strftime('%Y-%m-%d'), final_sym, in_div)
            )
            conn.commit()
            st.sidebar.success(f"🎉 {final_sym} 配息紀錄已成功寫入！")
            st.rerun()
        else:
            st.sidebar.error("❌ 請檢查：代碼不可為空，且配息金額必須大於 0！")

elif action_type == "🔮 設定未來預估配息率":
    st.sidebar.markdown("**💡 說明：** 請依照該標的「單次」發放的狀況填寫。")
    freq = st.sidebar.selectbox(
        "配息頻率",
        ["月配 (一年12次)", "季配 (一年4次)", "半年配 (一年2次)", "年配 (一年1次)"]
    )
    cash_val = st.sidebar.number_input("單次現金股利 (元/股)", min_value=0.0, step=0.01, format="%.3f")
    stock_val = st.sidebar.number_input("單次股票股利 (元/股)", min_value=0.0, step=0.1, format="%.3f")

    if st.sidebar.button("確認儲存預估參數"):
        final_sym = clean_symbol(raw_sym)
        if final_sym:
            c.execute(
                "INSERT OR REPLACE INTO est_dividend_v90 (symbol, frequency, cash_per_time, stock_per_time) VALUES (?, ?, ?, ?)",
                (final_sym, freq, cash_val, stock_val)
            )
            conn.commit()
            st.sidebar.success(f"🔮 {final_sym} 預估參數已成功儲存！")
            st.rerun()
        else:
            st.sidebar.error("❌ 請檢查：代碼不可為空！")

else:  # 📅 設定未來配息公告
    st.sidebar.markdown("**💡 說明：** 輸入已公告的除息日、發放日與每股配息，系統會依持股股數自動估算入帳金額。")
    ex_date = st.sidebar.date_input("除息日", datetime.now())
    pay_date = st.sidebar.date_input("發放日", datetime.now())
    dividend_per_share = st.sidebar.number_input("本次每股配息 (元/股)", min_value=0.0, step=0.01, format="%.4f")
    note = st.sidebar.text_input("備註（可空白）", placeholder="例如：官方公告、估算、暫定")

    if st.sidebar.button("確認寫入配息行事曆"):
        final_sym = clean_symbol(raw_sym)
        if final_sym and dividend_per_share > 0:
            c.execute(
                "INSERT INTO dividend_calendar_v101 (symbol, ex_dividend_date, payment_date, dividend_per_share, note) VALUES (?, ?, ?, ?, ?)",
                (final_sym, ex_date.strftime('%Y-%m-%d'), pay_date.strftime('%Y-%m-%d'), dividend_per_share, note)
            )
            conn.commit()
            st.sidebar.success(f"📅 {final_sym} 配息公告已寫入行事曆！")
            st.rerun()
        else:
            st.sidebar.error("❌ 請檢查：代碼不可為空，且每股配息必須大於 0！")

# --- 主面板：資料庫核心運算 ---
df_tx = pd.read_sql_query("SELECT * FROM tx_v74", conn)
df_div = pd.read_sql_query("SELECT * FROM dividend_v74", conn)
df_est = pd.read_sql_query("SELECT * FROM est_dividend_v90", conn)
df_cal = pd.read_sql_query("SELECT * FROM dividend_calendar_v101", conn)

if df_tx.empty:
    st.info("💡 目前系統中沒有庫存資料。請先在左側後台輸入您的第一筆資產庫存！")
else:
    df_tx['symbol'] = df_tx['symbol'].apply(lambda x: clean_symbol(x) if str(x).isdigit() else str(x).strip().upper())
    if not df_div.empty:
        df_div['symbol'] = df_div['symbol'].apply(lambda x: clean_symbol(x) if str(x).isdigit() else str(x).strip().upper())
    if not df_est.empty:
        df_est['symbol'] = df_est['symbol'].apply(lambda x: clean_symbol(x) if str(x).isdigit() else str(x).strip().upper())
    if not df_cal.empty:
        df_cal['symbol'] = df_cal['symbol'].apply(lambda x: clean_symbol(x) if str(x).isdigit() else str(x).strip().upper())

    # 彙整庫存數值
    p = df_tx.groupby('symbol').agg(
        投入金額=('amount', 'sum'),
        總手續費=('fee', 'sum'),
        總股數=('shares', 'sum')
    ).reset_index()

    p['總投入本金'] = p['投入金額'] + p['總手續費']
    p['平均成本'] = p['總投入本金'] / p['總股數'].replace(0, pd.NA)
    p['平均成本'] = p['平均成本'].fillna(0.0)

    # 彙整歷史配息
    if not df_div.empty:
        div_sum = df_div.groupby('symbol').agg(累積配息=('dividend_amount', 'sum')).reset_index()
        p = pd.merge(p, div_sum, on='symbol', how='left')
        p['累積配息'] = p['累積配息'].fillna(0.0)
    else:
        p['累積配息'] = 0.0

    # 整合未來預估配息參數
    if not df_est.empty:
        p = pd.merge(p, df_est, on='symbol', how='left')
        p['frequency'] = p['frequency'].fillna("未設定")
        p['cash_per_time'] = p['cash_per_time'].fillna(0.0)
        p['stock_per_time'] = p['stock_per_time'].fillna(0.0)
    else:
        p['frequency'] = "未設定"
        p['cash_per_time'] = 0.0
        p['stock_per_time'] = 0.0

    # 頻率換算年乘數
    freq_map = {
        "月配 (一年12次)": 12,
        "季配 (一年4次)": 4,
        "半年配 (一年2次)": 2,
        "年配 (一年1次)": 1
    }
    p['annual_multiplier'] = p['frequency'].map(freq_map).fillna(0)

    # 核心配息流計算：總股數與預估股利完全連動
    p['未來預估年現金'] = p['總股數'] * p['cash_per_time'] * p['annual_multiplier']
    p['未來預估月均現金'] = p['未來預估年現金'] / 12.0
    p['未來預估年配股_張'] = (p['總股數'] * p['stock_per_time'] / 10.0) / 1000.0

    # 全球即時報價同步：批次下載 + 快取
    with st.spinner('🔄 正在同步 Yahoo Finance 全球即時報價中...'):
        prices = get_prices(p['symbol'].tolist())

    for s in p['symbol']:
        if prices.get(s) is None:
            prices[s] = float(p.loc[p['symbol'] == s, '平均成本'].iloc[0])

    # 核心損益計算
    p['目前現價'] = p['symbol'].map(prices).astype(float)
    p['目前現值'] = p['總股數'] * p['目前現價']
    p['未實現損益'] = p['目前現值'] - p['總投入本金']
    p['總報酬率 (%)'] = ((p['目前現值'] + p['累積配息'] - p['總投入本金']) / p['總投入本金'].replace(0, pd.NA)) * 100
    p['總報酬率 (%)'] = p['總報酬率 (%)'].fillna(0.0)
    p['預估年化現金殖利率 (%)'] = (p['未來預估年現金'] / p['目前現值'].replace(0, pd.NA)) * 100
    p['預估年化現金殖利率 (%)'] = p['預估年化現金殖利率 (%)'].fillna(0.0)

    # 頂部大看板計算：全部整數顯示，不出現小數點
    t_amt = round(p['總投入本金'].sum())
    t_val = round(p['目前現值'].sum())
    t_div = round(p['累積配息'].sum())
    t_est_monthly_cash = round(p['未來預估月均現金'].sum())
    t_unrealized = round(p['未實現損益'].sum())
    total_return = 0 if t_amt == 0 else ((t_val + t_div - t_amt) / t_amt) * 100

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 總投入本金 (含費用)", f"${t_amt:,} 元")
    col2.metric("📈 總資產現值 (即時)", f"${t_val:,} 元")
    col3.metric("🎁 帳戶歷史累積總配息", f"${t_div:,} 元")
    col4.metric("🔮 預估總被動收入", f"${t_est_monthly_cash:,} 元/月", "所有資產連動後的月配息加總")

    col5, col6, col7, col8 = st.columns(4)
    col5.metric("📊 未實現損益", f"${t_unrealized:+,} 元")
    col6.metric("🚀 總報酬率", f"{total_return:+.2f}%")
    col7.metric("📅 預估年度現金流", f"${round(p['未來預估年現金'].sum()):,} 元/年")
    col8.metric("🌱 預估年度配股", f"{p['未來預估年配股_張'].sum():,.3f} 張/年")

    st.markdown("---")

    # 資產配置與損益排行
    left_chart, right_rank = st.columns([1.2, 1])

    with left_chart:
        st.subheader("📊 資產配置分析")
        pie_df = p[p['目前現值'] > 0][['symbol', '目前現值']].copy()
        if pie_df.empty:
            st.info("目前沒有可繪製的資產配置資料。")
        else:
            chart_df = pie_df.set_index('symbol')
            st.bar_chart(chart_df, use_container_width=True)
            pie_df['配置比例'] = pie_df['目前現值'] / pie_df['目前現值'].sum()
            pie_df['目前現值'] = pie_df['目前現值'].map(fmt_money0)
            pie_df['配置比例'] = pie_df['配置比例'].map(lambda x: f"{x:.2%}")
            st.dataframe(
                pie_df.rename(columns={'symbol': '資產代碼', '目前現值': '目前現值', '配置比例': '配置比例'}),
                use_container_width=True,
                hide_index=True
            )

    with right_rank:
        st.subheader("🏆 損益排行")
        gain_df = p.sort_values(by="未實現損益", ascending=False)[['symbol', '未實現損益', '總報酬率 (%)']].head(5).copy()
        loss_df = p.sort_values(by="未實現損益", ascending=True)[['symbol', '未實現損益', '總報酬率 (%)']].head(5).copy()

        gain_df['未實現損益'] = gain_df['未實現損益'].map(lambda x: f"${x:+,.0f}")
        gain_df['總報酬率 (%)'] = gain_df['總報酬率 (%)'].map(fmt_pct)
        loss_df['未實現損益'] = loss_df['未實現損益'].map(lambda x: f"${x:+,.0f}")
        loss_df['總報酬率 (%)'] = loss_df['總報酬率 (%)'].map(fmt_pct)

        st.caption("獲利 TOP 5")
        st.dataframe(gain_df.rename(columns={'symbol': '資產代碼'}), use_container_width=True, hide_index=True)
        st.caption("虧損 TOP 5")
        st.dataframe(loss_df.rename(columns={'symbol': '資產代碼'}), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("📋 綜合資產明細表 (核心現金流前置 · 整數不切邊)")

    # 資料表複製與格式化顯示
    v = p.copy()
    v['總投入本金顯示'] = v['總投入本金'].map(fmt_money0)
    v['總股數顯示'] = v['總股數'].map(lambda x: f"{x:,.0f} 股")
    v['平均成本顯示'] = v['平均成本'].map(fmt_money2)
    v['目前現價顯示'] = v['目前現價'].map(fmt_money2)
    v['目前現值顯示'] = v['目前現值'].map(fmt_money0)
    v['累積配息顯示'] = v['累積配息'].map(fmt_money0)
    v['未實現損益顯示'] = v['未實現損益'].map(lambda x: f"${x:+,.0f}")
    v['總報酬率顯示'] = v['總報酬率 (%)'].map(fmt_pct)
    v['預估年化現金殖利率顯示'] = v['預估年化現金殖利率 (%)'].map(lambda x: f"{x:.2f}%")

    # 格式化連動後預估指標
    v['單次現金'] = v['cash_per_time'].map(lambda x: f"${x:,.3f}")
    v['單次配股'] = v['stock_per_time'].map(lambda x: f"{x:,.3f} 元")
    v['預估每月現金'] = v['未來預估月均現金'].map(fmt_money0)
    v['預估年度現金'] = v['未來預估年現金'].map(fmt_money0)
    v['預估年度配股'] = v['未來預估年配股_張'].map(lambda x: f"{x:,.3f} 張")

    show_cols = [
        'symbol',
        '總股數顯示',
        '預估年化現金殖利率顯示',
        '平均成本顯示',
        '目前現價顯示',
        '目前現值顯示',
        '累積配息顯示',
        '未實現損益顯示',
        '總報酬率顯示'
    ]

    rename_dict = {
        'symbol': '資產代碼',
        '總股數顯示': '總股數',
        '總投入本金顯示': '總投入本金',
        '平均成本顯示': '平均成本',
        '目前現價顯示': '目前現價',
        '目前現值顯示': '目前現值',
        '累積配息顯示': '累積配息',
        '未實現損益顯示': '未實現損益',
        '總報酬率顯示': '總報酬率 (%)',
        '預估年化現金殖利率顯示': '預估現金殖利率',
        'frequency': '配息週期'
    }

    display_df = v[show_cols].rename(columns=rename_dict)
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # 未來配息行事曆：依已輸入公告 + 持股股數自動估算
    st.markdown("---")
    st.subheader("📅 未來配息行事曆")

    calendar_display_df = pd.DataFrame()

    if df_cal.empty:
        st.info("目前還沒有配息公告資料。請到左側後台選擇「📅 設定未來配息公告」新增除息日、發放日與每股配息。")
    else:
        calendar_work = df_cal.copy()
        calendar_work['ex_dividend_date'] = pd.to_datetime(calendar_work['ex_dividend_date'], errors='coerce')
        calendar_work['payment_date'] = pd.to_datetime(calendar_work['payment_date'], errors='coerce')

        calendar_work = pd.merge(
            calendar_work,
            p[['symbol', '總股數']],
            on='symbol',
            how='left'
        )
        calendar_work['總股數'] = calendar_work['總股數'].fillna(0.0)
        calendar_work['預估入帳金額'] = calendar_work['總股數'] * calendar_work['dividend_per_share']

        today = pd.Timestamp(datetime.now().date())
        upcoming = calendar_work[calendar_work['payment_date'].notna()].copy()
        upcoming = upcoming[upcoming['payment_date'] >= today].sort_values(by='payment_date')

        if upcoming.empty:
            st.info("目前沒有未來發放日的配息資料；舊資料仍可在下方管理分頁刪除。")
        else:
            this_month = today.strftime('%Y-%m')
            upcoming['發放月份'] = upcoming['payment_date'].dt.strftime('%Y-%m')

            total_announced = round(upcoming['預估入帳金額'].sum())
            month_income = round(upcoming.loc[upcoming['發放月份'] == this_month, '預估入帳金額'].sum())
            next_row = upcoming.iloc[0]

            ca, cb, cc = st.columns(3)
            ca.metric("💰 未來已公告預估入帳", f"${total_announced:,} 元")
            cb.metric("📆 本月預估入帳", f"${month_income:,} 元")
            cc.metric("⏭️ 下一筆入帳", f"{next_row['symbol']} / ${round(next_row['預估入帳金額']):,} 元", next_row['payment_date'].strftime('%Y-%m-%d'))

            calendar_display_df = upcoming[[
                'symbol', '總股數', 'ex_dividend_date', 'payment_date',
                'dividend_per_share', '預估入帳金額', 'note'
            ]].copy()

            calendar_display_df['總股數'] = calendar_display_df['總股數'].map(lambda x: f"{x:,.0f} 股")
            calendar_display_df['ex_dividend_date'] = calendar_display_df['ex_dividend_date'].dt.strftime('%Y-%m-%d')
            calendar_display_df['payment_date'] = calendar_display_df['payment_date'].dt.strftime('%Y-%m-%d')
            calendar_display_df['dividend_per_share'] = calendar_display_df['dividend_per_share'].map(lambda x: f"${x:,.4f}")
            calendar_display_df['預估入帳金額'] = calendar_display_df['預估入帳金額'].map(fmt_money0)
            calendar_display_df['note'] = calendar_display_df['note'].fillna('')

            calendar_display_df = calendar_display_df.rename(columns={
                'symbol': '資產代碼',
                '總股數': '持有股數',
                'ex_dividend_date': '除息日',
                'payment_date': '發放日',
                'dividend_per_share': '每股配息',
                '預估入帳金額': '預估入帳',
                'note': '備註'
            })

            st.dataframe(calendar_display_df, use_container_width=True, hide_index=True)

            monthly_cash = upcoming.groupby('發放月份')['預估入帳金額'].sum().reset_index()
            monthly_cash = monthly_cash.set_index('發放月份')
            st.caption("📊 依發放月份統計的未來現金流")
            st.bar_chart(monthly_cash, use_container_width=True)

    # 匯出 Excel
    st.download_button(
        label="📥 匯出 Excel 報表",
        data=make_excel_bytes(display_df, p, df_tx, df_div, df_est, calendar_display_df),
        file_name=f"鵬鵬退休計畫報表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- 下方明細分頁 ---
    st.markdown("---")
    tab1, tab2, tab3, tab4 = st.tabs(["📦 庫存買入明細管理", "💰 已收配息明細管理", "🔮 預估配息參數體檢", "📅 配息公告管理"])

    with tab1:
        if df_tx.empty:
            st.info("💡 目前還沒有買入明細。")
        else:
            for idx, row in df_tx.sort_values(by='date', ascending=False).iterrows():
                rid, rdate, rsym, ramt, rsh, rfe = row['id'], row['date'], row['symbol'], row['amount'], row['shares'], row['fee']
                c_1, c_2, c_3, c_4, c_btn = st.columns([2, 2, 3, 3, 2])
                c_1.write(f"📅 {rdate}")
                c_2.warning(f"🔤 {rsym}")
                c_3.write(f"本金: ${ramt:,.0f} (費: ${rfe:,.0f})")
                c_4.write(f"股數: {rsh:,.0f} 股")
                
                with st.expander(f"✏️ 修改 #{rid}"):
                    new_date = st.text_input("日期", value=str(rdate), key=f"edit_date_{rid}")
                    new_amt = st.number_input("本金", value=float(ramt), key=f"edit_amt_{rid}")
                    new_sh = st.number_input("股數", value=float(rsh), key=f"edit_sh_{rid}")
                    new_fee = st.number_input("手續費", value=float(rfe), key=f"edit_fee_{rid}")
                    if st.button("💾 儲存修改", key=f"save_tx_{rid}"):
                        c.execute(
                            "UPDATE tx_v74 SET date=?, amount=?, shares=?, fee=? WHERE id=?",
                            (new_date, new_amt, new_sh, new_fee, rid)
                        )
                        conn.commit()
                        st.rerun()

                if c_btn.button("🗑️ 刪除", key=f"del_tx_{rid}"):

                    c.execute("DELETE FROM tx_v74 WHERE id = ?", (rid,))
                    conn.commit()
                    st.rerun()

    with tab2:
        if df_div.empty:
            st.info("💡 目前還沒有任何實際配息紀錄。")
        else:
            for idx, row in df_div.sort_values(by='date', ascending=False).iterrows():
                rid, rdate, rsym, rdiv = row['id'], row['date'], row['symbol'], row['dividend_amount']
                c_1, c_2, c_3, c_btn = st.columns([3, 3, 4, 2])
                c_1.write(f"📅 {rdate}")
                c_2.success(f"🔤 {rsym}")
                c_3.write(f"實收配息: ${rdiv:,.0f} 元")
                if c_btn.button("🗑️ 刪除", key=f"del_div_{rid}"):
                    c.execute("DELETE FROM dividend_v74 WHERE id = ?", (rid,))
                    conn.commit()
                    st.rerun()

    with tab3:
        if df_est.empty:
            st.info("💡 目前尚未設定任何 ETF 的預估配息參數。請至左側後台切換設定。")
        else:
            st.caption("以下為您目前設定的未來多軌配息率參數：")
            for idx, row in df_est.iterrows():
                rsym, rfreq, rcash, rstock = row['symbol'], row['frequency'], row['cash_per_time'], row['stock_per_time']
                c_1, c_2, c_3, c_4, c_btn = st.columns([2, 3, 3, 3, 1])
                c_1.warning(f"🔤 {rsym}")
                c_2.write(f"週期: {rfreq}")
                c_3.write(f"單次現金: ${rcash:.3f}")
                c_4.write(f"單次股票: {rstock:.3f} 元")
                if c_btn.button("🗑️ 清除", key=f"del_est_{rsym}"):
                    c.execute("DELETE FROM est_dividend_v90 WHERE symbol = ?", (rsym,))
                    conn.commit()
                    st.rerun()

    with tab4:
        if df_cal.empty:
            st.info("💡 目前尚未建立任何未來配息公告。請至左側後台切換到「📅 設定未來配息公告」新增。")
        else:
            st.caption("以下為您目前手動建立的配息公告資料，可刪除錯誤或過期項目。")
            df_cal_show = df_cal.copy()
            df_cal_show = df_cal_show.sort_values(by=['payment_date', 'ex_dividend_date'], ascending=True)
            for idx, row in df_cal_show.iterrows():
                rid = row['id']
                rsym = row['symbol']
                rex = row['ex_dividend_date']
                rpay = row['payment_date']
                rdiv = row['dividend_per_share']
                rnote = row.get('note', '') or ''
                c_1, c_2, c_3, c_4, c_5, c_btn = st.columns([2, 2, 2, 2, 3, 1])
                c_1.warning(f"🔤 {rsym}")
                c_2.write(f"除息: {rex}")
                c_3.write(f"發放: {rpay}")
                c_4.write(f"每股: ${rdiv:.4f}")
                c_5.write(f"備註: {rnote}")
                if c_btn.button("🗑️", key=f"del_cal_{rid}"):
                    c.execute("DELETE FROM dividend_calendar_v101 WHERE id = ?", (rid,))
                    conn.commit()
                    st.rerun()



# =========================
# v10.5 退休儀表板
# =========================
try:
    st.markdown("---")
    st.header("🎯 退休規劃中心")

    col_a, col_b, col_c, col_d = st.columns(4)

    current_age = col_a.number_input("目前年齡", 20, 80, 43)
    retire_age = col_b.number_input("退休年齡", 40, 80, 55)
    monthly_invest = col_c.number_input("每月投入金額", 0, 500000, 50000, step=1000)
    target_asset = col_d.number_input("退休目標資產", 1000000, 100000000, 30000000, step=1000000)

    years_left = max(retire_age - current_age, 0)

    current_assets = float(t_val) if 't_val' in globals() else 0
    annual_dividend = float(p['未來預估年現金'].sum()) if 'p' in globals() else 0

    future_assets = current_assets
    for _ in range(years_left):
        future_assets = future_assets * 1.06 + monthly_invest * 12

    retirement_gap = max(target_asset - current_assets, 0)
    progress = min(current_assets / target_asset, 1.0)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("目前總資產", f"${current_assets:,.0f}")
    k2.metric("退休缺口", f"${retirement_gap:,.0f}")
    k3.metric("55歲預估資產", f"${future_assets:,.0f}")
    k4.metric("退休達成率", f"{progress*100:.1f}%")

    st.progress(progress)

    st.subheader("💰 配息覆蓋率")

    monthly_need = st.number_input("退休每月生活費", 10000, 300000, 60000, step=5000)
    passive_income = annual_dividend / 12
    cover_rate = (passive_income / monthly_need * 100) if monthly_need else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("目前被動收入", f"${passive_income:,.0f}/月")
    c2.metric("生活費需求", f"${monthly_need:,.0f}/月")
    c3.metric("覆蓋率", f"{cover_rate:.1f}%")

except Exception as e:
    st.warning(f"退休模組載入失敗: {e}")
