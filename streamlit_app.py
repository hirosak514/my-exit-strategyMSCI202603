import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime
import google.generativeai as genai
from PIL import Image
import json
import re
import os
import copy
import gspread
from google.oauth2.service_account import Credentials

# --- 0. データの保存・読み込み ---
DB_FILE = "portfolio.json"
EVENT_FILE = "events.json"
REMINDER_FILE = "reminder.json"
CONFIG_FILE = "config.json"

# 【改修箇所】対象のスプレッドシートURLをここに記述してください
FIXED_SHEET_URL = "https://docs.google.com/spreadsheets/d/17kAFl14q8EaaQ6kvezlAe1Yzr71Yo673T61--_cyESQ/edit"

def load_json(file_path, default_value):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return default_value

def save_json(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# --- Google Spreadsheet 連携関数 ---
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Streamlit secrets または環境変数から認証情報を取得
        creds_dict = st.secrets["gcp_service_account"]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Google認証に失敗しました: {e}")
        return None

def export_to_spreadsheet(data):
    gc = get_gspread_client()
    if not gc: return
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ws.clear()
        ws.update('A1', [[json.dumps(data, ensure_ascii=False)]])
        st.success("スプレッドシートへのエクスポートが完了しました")
    except Exception as e:
        st.error(f"エクスポート失敗: {e}")

def import_from_spreadsheet():
    gc = get_gspread_client()
    if not gc: return None
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        content = ws.acell('A1').value
        return json.loads(content) if content else None
    except Exception as e:
        st.error(f"インポート失敗: {e}")
        return None

# --- 1. セッション状態の初期化 ---
if 'portfolio' not in st.session_state:
    st.session_state.portfolio = load_json(DB_FILE, {})
if 'prev_portfolio' not in st.session_state:
    st.session_state.prev_portfolio = None
if 'events' not in st.session_state:
    st.session_state.events = load_json(EVENT_FILE, [])
if 'reminder_text' not in st.session_state:
    st.session_state.reminder_text = load_json(REMINDER_FILE, "- ターゲット日程を入力してください")
if 'api_key' not in st.session_state:
    st.session_state.api_key = load_json(CONFIG_FILE, {"gemini_key": ""}).get("gemini_key", "")

# 復元用バックアップ
def backup_portfolio():
    st.session_state.prev_portfolio = copy.deepcopy(st.session_state.portfolio)

# --- 2. API設定 ---
current_api_key = st.session_state.api_key or st.secrets.get("GEMINI_API_KEY", "")
if current_api_key:
    genai.configure(api_key=current_api_key)

# --- 3. 解析・価格取得関数 ---
def get_live_prices(portfolio_keys):
    prices = {}
    for key in portfolio_keys:
        symbol = key.split('_')[0]
        is_japan = symbol.isdigit() and len(symbol) == 4
        ticker = f"{symbol}.T" if is_japan else ("7013.T" if symbol == "IHI" else symbol)
        
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period="5d")
            if not hist.empty:
                prices[key] = {
                    "current": hist['Close'].iloc[-1],
                    "prev_close": hist['Close'].iloc[-2] if len(hist) >= 2 else None
                }
            else:
                prices[key] = None
        except:
            prices[key] = None
            
    try:
        usdjpy = yf.Ticker("JPY=X").history(period="5d")
        prices["USDJPY"] = usdjpy['Close'].iloc[-1] if not usdjpy.empty else 159.2
    except:
        prices["USDJPY"] = 159.2
    return prices

# 【オリジナルを完全踏襲】
def analyze_multiple_images(uploaded_files):
    if not current_api_key:
        raise ValueError("APIキーが設定されていません。サイドバーで設定してください。")
    
    available_models = [m.name for m in genai.list_models() if "generateContent" in m.supported_generation_methods]
    target_model = next((m for m in available_models if "flash" in m), available_models[0])
    model = genai.GenerativeModel(target_model)

    prompt = """
    証券口座のスクリーンショット（複数可）から、保有銘柄の情報を抽出して、以下のJSON形式のみで回答してください。
    余計な説明や装飾（```json など）は一切不要です。

    【抽出ルール】
    1. キーは「銘柄コード_区分」としてください。
        - 現物株の場合：コードのみ（例：8136_現物, NVDA_現物）
        - 信用買い（制度・無期限）の場合：末尾に _MARGIN_LONG（例：8136_MARGIN_LONG）
        - 信用売りの場合：末尾に _SHORT（例：8136_SHORT）
    2. 銘柄コードが不明な場合は、銘柄名をアルファベット表記にして代用してください。
    3. 数値（数量、取得単価）からカンマや円、ドル記号を除去して数値のみにしてください。
    4. 通貨は、日本株なら "JPY"、米国株なら "USD" としてください。

    【出力フォーマット】
    {
      "銘柄コード_区分": {"name": "銘柄名", "shares": 数量, "cost": 取得単価, "currency": "通貨"},
      ...
    }
    """
    
    images = []
    for uploaded_file in uploaded_files:
        images.append(Image.open(uploaded_file))
    
    response = model.generate_content([prompt] + images)
    json_match = re.search(r'\{.*\}', response.text, re.DOTALL)
    if json_match:
        return json.loads(json_match.group())
    else:
        raise ValueError("AI解析に失敗しました。画像の形式や内容を確認してください。")

# --- 4. UI ---
st.set_page_config(page_title="Strategist Dashboard", layout="wide")

st.markdown("""
<style>
    [data-testid="stMetricDelta"] > div { color: white !important; }
    div[data-testid="column"]:nth-child(3) button {
        background-color: #ff4b4b !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("🔑 Settings")
    new_api_key = st.text_input("Gemini API Key", value=st.session_state.api_key, type="password")
    if st.button("APIキーを保存"):
        st.session_state.api_key = new_api_key
        save_json(CONFIG_FILE, {"gemini_key": new_api_key})
        st.success("APIキーを保存しました")
        st.rerun()

    st.divider()

    st.header("✏️ 銘柄情報の直接入力")
    portfolio_items = list(st.session_state.portfolio.keys())
    selected_no = None
    if portfolio_items:
        no_options = [i + 1 for i in range(len(portfolio_items))]
        selected_no = st.selectbox("銘柄No.を選択", options=no_options)
        target_key = portfolio_items[selected_no - 1]
        target_info = st.session_state.portfolio[target_key]
        new_shares = st.number_input(f"数量 ({target_key})", value=float(target_info.get('shares', 0)))
        new_cost = st.number_input(f"取得単価 ({target_key})", value=float(target_info.get('cost', 0)))
    else:
        st.info("編集する銘柄がありません")
        new_shares, new_cost = 0.0, 0.0

    btn_col1, btn_col2, btn_col3 = st.columns(3)
    mod_ready = btn_col1.button("修正")
    rev_ready = btn_col2.button("復元", type="primary")
    del_ready = btn_col3.button("削除")

    if selected_no:
        if mod_ready:
            backup_portfolio()
            if new_shares == 0:
                st.session_state.portfolio[target_key]['shares'] = 0
                st.session_state.portfolio[target_key]['cost'] = 0
            else:
                st.session_state.portfolio[target_key]['shares'] = new_shares
                st.session_state.portfolio[target_key]['cost'] = new_cost
            save_json(DB_FILE, st.session_state.portfolio)
            st.rerun()

        if rev_ready:
            if st.session_state.prev_portfolio is not None:
                st.session_state.portfolio = copy.deepcopy(st.session_state.prev_portfolio)
                st.session_state.prev_portfolio = None
                save_json(DB_FILE, st.session_state.portfolio)
                st.rerun()
            else:
                st.error("復元できる履歴がありません")

        if del_ready:
            backup_portfolio()
            del st.session_state.portfolio[target_key]
            save_json(DB_FILE, st.session_state.portfolio)
            st.rerun()

    st.divider()
    
    st.header("📌 Event Manager")
    with st.expander("イベントの追加/削除"):
        ev_name = st.text_input("イベント名")
        ev_date = st.date_input("日付")
        if st.button("イベント追加"):
            st.session_state.events.append({"name": ev_name, "date": ev_date.strftime("%Y-%m-%d")})
            save_json(EVENT_FILE, st.session_state.events)
            st.rerun()
        if st.session_state.events:
            idx = st.selectbox("削除するイベント", range(len(st.session_state.events)), format_func=lambda x: st.session_state.events[x]['name'])
            if st.button("選択したイベントを削除"):
                st.session_state.events.pop(idx)
                save_json(EVENT_FILE, st.session_state.events)
                st.rerun()

    st.divider()
    st.header("📋 Reminder Edit")
    new_reminder = st.text_area("リマインダー内容", value=st.session_state.reminder_text)
    if st.button("リマインダー更新"):
        st.session_state.reminder_text = new_reminder
        save_json(REMINDER_FILE, new_reminder)
        st.rerun()

    st.divider()
    st.subheader("💾 Backup (Spreadsheet)")
    full_config = {"portfolio": st.session_state.portfolio, "events": st.session_state.events, "reminder_text": st.session_state.reminder_text}
    
    # 【改修箇所】固定URLを使用するように変更
    if st.button("設定をエクスポート"):
        export_to_spreadsheet(full_config)

    if st.button("設定をインポート"):
        imported_data = import_from_spreadsheet()
        if imported_data:
            backup_portfolio()
            st.session_state.portfolio = imported_data.get("portfolio", {})
            st.session_state.events = imported_data.get("events", [])
            st.session_state.reminder_text = imported_data.get("reminder_text", "")
            save_json(DB_FILE, st.session_state.portfolio)
            save_json(EVENT_FILE, st.session_state.events)
            save_json(REMINDER_FILE, st.session_state.reminder_text)
            st.success("インポート完了")
            st.rerun()

    st.divider()
    st.header("📸 AI Scanner")
    up_files = st.file_uploader("証券口座のスクショをアップロード", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    if up_files and st.button("AI解析実行"):
        with st.spinner("AIが銘柄を抽出中..."):
            try:
                extracted_data = analyze_multiple_images(up_files)
                backup_portfolio()
                st.session_state.portfolio = extracted_data
                save_json(DB_FILE, st.session_state.portfolio)
                st.rerun()
            except Exception as e: st.error(f"エラー: {e}")

# --- 5. メイン画面 ---
st.title("🚀 Strategist Dashboard")

if st.session_state.events:
    st.write("📌 **重要スケジュール**")
    cols = st.columns(len(st.session_state.events))
    for i, event in enumerate(st.session_state.events):
        try:
            target_date = datetime.strptime(event['date'], "%Y-%m-%d")
            days_left = (target_date - datetime.now()).days
            cols[i].markdown(f"<small>{event['name']}</small>", unsafe_allow_html=True)
            cols[i].metric("", event['date'], f"あと {days_left} 日")
        except: pass

st.divider()
st.header("📉 Portfolio Monitor")
if st.button('最新価格に更新'): st.rerun()

prices_dict = get_live_prices(st.session_state.portfolio.keys())
rate = prices_dict.get("USDJPY", 159.2)

rows = []
total_profit_jpy = 0
total_profit_usd_only_us_stocks = 0

for i, (key, info) in enumerate(st.session_state.portfolio.items()):
    p_data = prices_dict.get(key)
    if p_data and info.get('shares', 0) >= 0:
        cur, prev = p_data["current"], p_data["prev_close"]
        day_change_pct = f"({(cur - prev) / prev * 100:+.2f}%)" if prev else ""
        display_name = f"{key.split('_')[0]} {info.get('name','')}"
        
        if info['shares'] == 0:
            p_jpy = 0
        else:
            if "_SHORT" in key:
                label, p_jpy = "信用(売建)", (info['cost'] - cur) * info['shares']
            elif "_MARGIN_LONG" in key:
                label, p_jpy = "信用(買建)", (cur - info['cost']) * info['shares']
            else:
                label = "現物"
                if info.get('currency') == "USD":
                    p_usd = (cur - info['cost']) * info['shares']
                    p_jpy = p_usd * rate
                    total_profit_usd_only_us_stocks += p_usd
                else: p_jpy = (cur - info['cost']) * info['shares']

        total_profit_jpy += p_jpy
        cost_display = f"${info['cost']:,}" if info.get('currency') == "USD" else f"¥{info['cost']:,}"
        cur_display = f"{('$' if info.get('currency') == 'USD' else '¥')}{cur:,.2f} {day_change_pct}"
        
        rows.append({
            "No.": i + 1, "銘柄": display_name, "数量": info['shares'], "区分": label if info['shares'] > 0 else "決済済",
            "取得単価": cost_display, "現在値 (前日比)": cur_display, "損益(円)": f"¥{p_jpy:,.0f}"
        })

m_col1, m_col2 = st.columns(2)
m_col1.metric("総合計損益 (JPY)", f"¥{total_profit_jpy:,.0f}", delta=f"USD/JPY: {rate:.2f}")
m_col2.metric("米国株合計損益 (USD)", f"${total_profit_usd_only_us_stocks:,.2f}")

if rows: st.table(pd.DataFrame(rows))
else: st.info("銘柄がありません")

st.divider()
st.subheader("📋 Reminder")
st.info(st.session_state.reminder_text)
