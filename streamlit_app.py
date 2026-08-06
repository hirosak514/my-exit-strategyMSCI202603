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
import csv
from streamlit_autorefresh import st_autorefresh

# --- 0. データの保存・読み込み ---
DB_FILE = "portfolio.json"
EVENT_FILE = "events.json"
REMINDER_FILE = "reminder.json"
CONFIG_FILE = "config.json"

# 【改修箇所】対象のスプレッドシートURLをここに記述してください
FIXED_SHEET_URL = "https://docs.google.com/spreadsheets/d/17kAFl14q8EaaQ6kvezlAe1Yzr71Yo673T61--_cyESQ/edit"

# --- バックアップ履歴 設定 ---
MAX_BACKUPS = 1000          # 保持する最大バックアップ件数（超えたら古い順に削除）
INITIAL_CACHE_SIZE = 10     # 初回「1つ前の設定」で読み込む件数（直近N件）
WINDOW_CACHE_RADIUS = 10    # 範囲外アクセス時に読み込む前後件数（前後N件＝計2N件）
TIMESTAMP_FMT = "%Y-%m-%d %H:%M:%S"       # スプレッドシート保存用の内部フォーマット
DISPLAY_FMT = "%Y年%m月%d日 %H:%M"         # ユーザー表示用のフォーマット

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

def ensure_header(ws):
    """A1セルが 'timestamp' でなければヘッダー行を先頭に挿入する"""
    try:
        first_cell = ws.acell("A1").value
    except Exception:
        first_cell = None
    if first_cell != "timestamp":
        # 旧形式（A1に直接JSONを書いていた版）が残っている場合、
        # そのままだと2行目に旧データが紛れ込みます。事前に手動でシートを
        # クリアしておくことを推奨します。
        ws.insert_row(["timestamp", "data"], 1)

def get_backup_count(ws):
    """タイムスタンプ列（A列）だけを読み、件数を安価に取得する"""
    try:
        col = ws.col_values(1)  # ヘッダー込み
        return max(0, len(col) - 1)
    except Exception:
        return 0

def fetch_backup_range(ws, start_idx, end_idx, total):
    """
    絶対インデックス start_idx〜end_idx（1=最古 … total=最新）の
    範囲だけをまとめて取得する。
    """
    start_idx = max(1, start_idx)
    end_idx = min(total, end_idx)
    if start_idx > end_idx:
        return {}, start_idx, end_idx

    start_row = start_idx + 1  # +1 はヘッダー行分
    end_row = end_idx + 1
    values = ws.get(f"A{start_row}:B{end_row}")

    cache = {}
    for offset, row in enumerate(values):
        idx = start_idx + offset
        if len(row) >= 2 and row[1]:
            try:
                cache[idx] = {"timestamp": row[0], "data": json.loads(row[1])}
            except Exception:
                continue
    return cache, start_idx, end_idx

def overwrite_backup_at_index(idx, data):
    """
    絶対インデックス idx の行を、新しい内容で上書きする。
    - 主：セッションに保持している idx（=行番号-1）から直接特定
    - 副：上書き直前に実際のタイムスタンプとキャッシュ上のタイムスタンプを照合し、
          ズレていれば列A全体を検索してフォールバックする
    戻り値: 成功時は書き込んだ行のタイムスタンプ文字列、失敗時は None
    """
    gc = get_gspread_client()
    if not gc: return None
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)

        cached_entry = st.session_state.backup_cache.get(idx)
        expected_ts = cached_entry["timestamp"] if cached_entry else None
        row_number = idx + 1  # +1 はヘッダー行分

        actual_row = ws.row_values(row_number)
        actual_ts = actual_row[0] if actual_row else None

        if expected_ts and actual_ts != expected_ts:
            # 行がズレている可能性 → タイムスタンプで再検索してフォールバック
            col = ws.col_values(1)
            try:
                row_number = col.index(expected_ts) + 1  # 1-based行番号（col[0]がヘッダー）
            except ValueError:
                st.error("上書き対象のバックアップがシート上に見つかりませんでした（削除された可能性があります）。"
                          "「◀ 1つ前の設定」で対象を読み込み直してください。")
                return None

        # タイムスタンプは維持したまま、データ部分のみ上書きする
        keep_ts = expected_ts or actual_ts or datetime.now().strftime(TIMESTAMP_FMT)
        ws.update(f"A{row_number}:B{row_number}", [[keep_ts, json.dumps(data, ensure_ascii=False)]])
        return keep_ts
    except Exception as e:
        st.error(f"上書き失敗: {e}")
        return None

def export_to_spreadsheet(data):
    """新規バックアップを1行追記する（上書きしない）。1000件超は古い順に削除。"""
    gc = get_gspread_client()
    if not gc: return
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)

        timestamp = datetime.now().strftime(TIMESTAMP_FMT)
        ws.append_row([timestamp, json.dumps(data, ensure_ascii=False)])

        total = get_backup_count(ws)
        if total > MAX_BACKUPS:
            excess = total - MAX_BACKUPS
            # 2行目（最古データ）から excess 件分をまとめて削除
            ws.delete_rows(2, 1 + excess)

        display_ts = datetime.strptime(timestamp, TIMESTAMP_FMT).strftime(DISPLAY_FMT)
        st.success(f"バックアップを保存しました（{display_ts}）")
    except Exception as e:
        st.error(f"エクスポート失敗: {e}")

def load_backup_window(center_idx=None, initial=False):
    """
    initial=True: 直近 INITIAL_CACHE_SIZE 件を読み込み、最新（=1つ前）を適用
    initial=False: center_idx の前後 WINDOW_CACHE_RADIUS 件（計最大2N件）を読み込む
    """
    gc = get_gspread_client()
    if not gc: return
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)

        total = get_backup_count(ws)
        st.session_state.backup_total = total
        if total == 0:
            st.warning("バックアップ履歴が見つかりません")
            return

        if initial:
            start_idx = max(1, total - INITIAL_CACHE_SIZE + 1)
            end_idx = total
            target_idx = total  # 最新 = "1つ前"の最初の到達点
        else:
            start_idx = max(1, center_idx - WINDOW_CACHE_RADIUS)
            end_idx = min(total, center_idx + WINDOW_CACHE_RADIUS)
            target_idx = center_idx

        cache, s, e = fetch_backup_range(ws, start_idx, end_idx, total)
        if not cache:
            st.warning("該当する履歴データが見つかりませんでした")
            return

        st.session_state.backup_cache = cache
        st.session_state.cache_min = s
        st.session_state.cache_max = e
        apply_backup_index(target_idx)
    except Exception as e:
        st.error(f"履歴取得失敗: {e}")

def apply_backup_index(idx):
    """キャッシュ済みの idx 番目のバックアップをセッションに反映する"""
    entry = st.session_state.backup_cache.get(idx)
    if not entry:
        st.warning("そのデータはまだキャッシュされていません")
        return
    backup_portfolio()
    data = entry["data"]
    st.session_state.portfolio = data.get("portfolio", {})
    st.session_state.events = data.get("events", [])
    st.session_state.reminder_text = data.get("reminder_text", "")
    st.session_state.backup_index = idx
    save_json(DB_FILE, st.session_state.portfolio)
    save_json(EVENT_FILE, st.session_state.events)
    save_json(REMINDER_FILE, st.session_state.reminder_text)

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
# --- バックアップ履歴（前後移動）用の状態 ---
if 'backup_cache' not in st.session_state:
    st.session_state.backup_cache = {}      # {絶対インデックス: {"timestamp":..., "data":...}}
if 'cache_min' not in st.session_state:
    st.session_state.cache_min = None
if 'cache_max' not in st.session_state:
    st.session_state.cache_max = None
if 'backup_index' not in st.session_state:
    st.session_state.backup_index = None    # 現在表示中の絶対インデックス（1=最古）
if 'backup_total' not in st.session_state:
    st.session_state.backup_total = None

# 復元用バックアップ
def backup_portfolio():
    st.session_state.prev_portfolio = copy.deepcopy(st.session_state.portfolio)

# --- 2. API設定 ---
current_api_key = st.session_state.api_key or st.secrets.get("GEMINI_API_KEY", "")
if current_api_key:
    genai.configure(api_key=current_api_key)

# --- 3. 解析・価格取得関数 ---
def search_and_add_stock(query, add_type_label, shares, cost):
    """
    証券コード または 銘柄名 で検索し、見つかれば portfolio に新規追加する。
    戻り値: 成功時は追加したキー文字列、失敗時は None
    """
    type_suffix = {"現物": "", "信用(買建)": "_MARGIN_LONG", "信用(売建)": "_SHORT"}.get(add_type_label, "")
    q = query.strip()
    if not q:
        return None

    candidates = []  # (表示コード, yfinanceティッカー, 通貨)
    if q.isdigit() and len(q) == 4:
        candidates.append((q, f"{q}.T", "JPY"))
    candidates.append((q.upper(), q.upper(), "USD"))

    # 1. コード直接指定として試す
    for code, ticker, currency in candidates:
        try:
            info = yf.Ticker(ticker).info
            name = info.get("shortName") or info.get("longName")
            price = info.get("regularMarketPrice") or info.get("previousClose")
            if name and price:
                key = f"{code}{type_suffix}"
                st.session_state.portfolio[key] = {
                    "name": name, "shares": shares, "cost": cost, "currency": currency
                }
                return key
        except Exception:
            continue

    # 2. 直接指定で見つからなければ、銘柄名によるあいまい検索
    try:
        search_result = yf.Search(q, max_results=5)
        quotes = getattr(search_result, "quotes", []) or []
        for item in quotes:
            symbol = item.get("symbol")
            name = item.get("shortname") or item.get("longname")
            if not symbol:
                continue
            is_japan = symbol.endswith(".T")
            code = symbol.replace(".T", "") if is_japan else symbol
            currency = "JPY" if is_japan else "USD"
            key = f"{code}{type_suffix}"
            st.session_state.portfolio[key] = {
                "name": name or code, "shares": shares, "cost": cost, "currency": currency
            }
            return key
    except Exception:
        pass

    return None

def parse_stock_csv(uploaded_file):
    """
    1行目が '#' で始まるコメント行（例: # トレンド銘柄...,保存日時:...,市場:jp）の場合はスキップし、
    'code,name' ヘッダーを持つCSVから銘柄コード・銘柄名のリストを抽出する。
    戻り値: [(code, name), ...]
    """
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    try:
        text = raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = raw.decode('shift_jis', errors='ignore')

    lines = [l for l in text.splitlines() if l.strip() and not l.lstrip().startswith('#')]
    if not lines:
        return []

    reader = csv.DictReader(lines)
    # ヘッダーの大文字小文字・前後空白ゆれを吸収
    fieldmap = {(f or '').strip().lower(): f for f in (reader.fieldnames or [])}
    code_field = fieldmap.get('code')
    name_field = fieldmap.get('name')

    results = []
    for row in reader:
        code = (row.get(code_field) or '').strip() if code_field else ''
        name = (row.get(name_field) or '').strip() if name_field else ''
        if code:
            results.append((code, name))
    return results

def build_entries_from_csv(rows, default_shares=100):
    """
    [(code, name), ...] から portfolio 用のエントリ辞書を作る。
    価格は現在株価を自動取得（取得できない場合は0）。区分は「現物」固定。
    """
    entries = {}
    for code, name in rows:
        is_japan = code.isdigit() and len(code) == 4
        ticker = f"{code}.T" if is_japan else ("7013.T" if code == "IHI" else code.upper())
        currency = "JPY" if is_japan else "USD"

        price = 0.0
        try:
            hist = yf.Ticker(ticker).history(period="5d")
            if not hist.empty:
                price = float(hist['Close'].iloc[-1])
        except Exception:
            price = 0.0

        key = f"{code}_現物"
        entries[key] = {
            "name": name or code,
            "shares": default_shares,
            "cost": round(price, 2),
            "currency": currency
        }
    return entries

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

PRICE_CACHE_TTL_SECONDS = 300  # 5分

@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def _fetch_prices_cached(keys_tuple):
    """get_live_prices の結果を5分間キャッシュする。取得時刻も併せて返す。"""
    prices = get_live_prices(list(keys_tuple))
    return prices, datetime.now()

def get_prices_with_cache(portfolio_keys):
    """
    ポートフォリオのキー集合から価格を取得する（5分キャッシュ付き）。
    戻り値: (prices_dict, last_updated_datetime)
    """
    keys_tuple = tuple(sorted(portfolio_keys))
    if not keys_tuple:
        return {"USDJPY": 159.2}, datetime.now()
    return _fetch_prices_cached(keys_tuple)

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

    new_symbol_query = st.text_input(
        "🔍 銘柄コード/名称で新規追加（入力時のみ有効）",
        value="", key="add_stock_query",
        placeholder="例：7203 / NVDA / トヨタ"
    )
    add_mode = bool(new_symbol_query.strip())

    if add_mode:
        add_type = st.selectbox("区分", ["現物", "信用(買建)", "信用(売建)"], key="add_stock_type")
        new_shares = st.number_input("数量（新規）", value=0.0, min_value=0.0, key="add_stock_shares")
        new_cost = st.number_input("取得単価（新規）", value=0.0, min_value=0.0, key="add_stock_cost")
        selected_no = None
        target_key = None
    else:
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
            target_key = None

    btn_col1, btn_col2, btn_col3 = st.columns(3)
    mod_ready = btn_col1.button("修正")
    rev_ready = btn_col2.button("復元", type="primary")
    del_ready = btn_col3.button("削除")

    if mod_ready:
        if add_mode:
            with st.spinner(f"「{new_symbol_query}」を検索中..."):
                backup_portfolio()
                added_key = search_and_add_stock(new_symbol_query, add_type, new_shares, new_cost)
            if added_key:
                save_json(DB_FILE, st.session_state.portfolio)
                st.success(f"銘柄「{added_key}」を追加しました")
                del st.session_state["add_stock_query"]
                st.rerun()
            else:
                st.session_state.prev_portfolio = None  # 追加失敗時はバックアップを取り消す
                st.error(f"「{new_symbol_query}」に該当する銘柄が見つかりませんでした")
        elif selected_no:
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
        if add_mode:
            st.warning("削除は既存銘柄を選択している場合のみ有効です")
        elif selected_no:
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
    
    exp_col1, exp_col2 = st.columns(2)
    new_backup_clicked = exp_col1.button("🆕 新規バックアップ")
    overwrite_clicked = exp_col2.button("♻️ 上書きバックアップ")

    if new_backup_clicked:
        export_to_spreadsheet(full_config)

    if overwrite_clicked:
        if st.session_state.backup_index is None:
            st.warning("上書き対象が選択されていません。先に「◀ 1つ前の設定」で対象のバックアップを表示してください。")
        else:
            written_ts = overwrite_backup_at_index(st.session_state.backup_index, full_config)
            if written_ts:
                # ローカルキャッシュも同時に更新しておく（表示との整合性を保つため）
                entry = st.session_state.backup_cache.get(st.session_state.backup_index)
                if entry:
                    entry["data"] = full_config
                try:
                    display_ts = datetime.strptime(written_ts, TIMESTAMP_FMT).strftime(DISPLAY_FMT)
                except Exception:
                    display_ts = written_ts
                st.success(f"バックアップを上書きしました（{display_ts}）")

    # --- 前後ナビゲーション ---
    nav_col1, nav_col2 = st.columns(2)
    prev_clicked = nav_col1.button("◀ 1つ前の設定")
    next_clicked = nav_col2.button("1つ後の設定 ▶")

    if prev_clicked:
        if st.session_state.backup_index is None:
            # 初回：直近 INITIAL_CACHE_SIZE 件を読み込み、最新のものを表示
            load_backup_window(initial=True)
        else:
            target = st.session_state.backup_index - 1
            if target < 1:
                st.warning("これ以上前の履歴はありません")
            elif target < st.session_state.cache_min:
                load_backup_window(center_idx=target)
            else:
                apply_backup_index(target)
        st.rerun()

    if next_clicked:
        if st.session_state.backup_index is None:
            st.info("先に「◀ 1つ前の設定」を押してください")
        else:
            total = st.session_state.backup_total or st.session_state.backup_index
            target = st.session_state.backup_index + 1
            if target > total:
                st.warning("これより新しい履歴はありません（最新のバックアップです）")
            elif target > st.session_state.cache_max:
                load_backup_window(center_idx=target)
            else:
                apply_backup_index(target)
        st.rerun()

    # --- 現在表示中のバックアップ日時を表示 ---
    if st.session_state.backup_index is not None:
        entry = st.session_state.backup_cache.get(st.session_state.backup_index)
        if entry:
            try:
                dt = datetime.strptime(entry["timestamp"], TIMESTAMP_FMT)
                display_ts = dt.strftime(DISPLAY_FMT)
            except Exception:
                display_ts = entry["timestamp"]
            total_disp = st.session_state.backup_total or "?"
            st.caption(f"📅 表示中のバックアップ：**{display_ts}**（{st.session_state.backup_index} / {total_disp} 件目）")

    st.divider()
    st.header("📄 CSV読み込み")
    csv_file = st.file_uploader(
        "銘柄リストCSVをアップロード（ドラッグ&ドロップ可）",
        type=["csv"], key="csv_uploader"
    )
    csv_mode = st.radio(
        "読み込み方法",
        ["現在の画面に追加", "新規リストとして作成（既存を置き換え）"],
        key="csv_mode"
    )
    if st.button("csv読み込み"):
        if csv_file:
            with st.spinner("CSVを解析し、現在株価を取得中..."):
                try:
                    rows = parse_stock_csv(csv_file)
                    if not rows:
                        st.warning("CSVから銘柄を読み取れませんでした。フォーマットをご確認ください。")
                    else:
                        new_entries = build_entries_from_csv(rows, default_shares=100)
                        backup_portfolio()
                        if csv_mode == "現在の画面に追加":
                            st.session_state.portfolio.update(new_entries)
                        else:
                            st.session_state.portfolio = new_entries
                        save_json(DB_FILE, st.session_state.portfolio)
                        st.success(f"{len(new_entries)}件の銘柄を読み込みました")
                        st.rerun()
                except Exception as e:
                    st.error(f"CSV読み込みエラー: {e}")
        else:
            st.warning("CSVファイルがアップロードされていません。")

    st.divider()
    st.header("📸 AI Scanner")
    up_files = st.file_uploader("証券口座のスクショをアップロード", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    # 【改修３】解析ボタンを常に表示（if文の構造を変更）
    if st.button("AI解析実行"):
        if up_files:
            with st.spinner("AIが銘柄を抽出中..."):
                try:
                    extracted_data = analyze_multiple_images(up_files)
                    backup_portfolio()
                    st.session_state.portfolio = extracted_data
                    save_json(DB_FILE, st.session_state.portfolio)
                    st.rerun()
                except Exception as e: st.error(f"エラー: {e}")
        else:
            st.warning("画像がアップロードされていません。")

# --- 5. メイン画面 ---
st.title("🚀 Strategist Dashboard")

# 【改修１】Portfolio Monitor を最上部に配置
st.header("📉 Portfolio Monitor")

# 5分（300,000ミリ秒）ごとに自動で画面を再実行する
st_autorefresh(interval=PRICE_CACHE_TTL_SECONDS * 1000, key="auto_price_refresh")

col_refresh, col_ts = st.columns([1, 3])
if col_refresh.button('最新価格に更新'):
    # キャッシュを明示的に破棄してから再実行（手動更新は必ず最新値を取りに行く）
    _fetch_prices_cached.clear()
    st.rerun()

prices_dict, last_updated = get_prices_with_cache(st.session_state.portfolio.keys())
col_ts.caption(f"🕒 最終更新: {last_updated.strftime('%Y年%m月%d日 %H:%M:%S')}（5分ごとに自動更新）")
rate = prices_dict.get("USDJPY", 159.2)

rows = []
total_profit_jpy = 0
total_profit_usd_only_us_stocks = 0

for i, (key, info) in enumerate(st.session_state.portfolio.items()):
    shares = info.get('shares', 0)
    if shares < 0:
        continue

    p_data = prices_dict.get(key)
    display_name = f"{key.split('_')[0]} {info.get('name','')}"

    # --- 価格取得の成否を判定（NaN・取得失敗の両方をここで吸収する） ---
    cur, prev = None, None
    price_available = False
    if p_data:
        cur, prev = p_data.get("current"), p_data.get("prev_close")
        if cur is not None and not pd.isna(cur):
            price_available = True

    day_change_pct = ""
    if price_available and prev is not None and not pd.isna(prev) and prev != 0:
        day_change_pct = f"({(cur - prev) / prev * 100:+.2f}%)"

    if shares == 0:
        p_jpy = 0
        label = "決済済"
    elif not price_available:
        # 価格が取得できない銘柄は損益0円として扱い、合計に影響を与えない
        p_jpy = 0
        label = "信用(売建)" if "_SHORT" in key else ("信用(買建)" if "_MARGIN_LONG" in key else "現物")
    else:
        # --- [修正版] 損益計算ロジック ---
        if "_SHORT" in key:
            label = "信用(売建)"
            diff = info['cost'] - cur
        elif "_MARGIN_LONG" in key:
            label = "信用(買建)"
            diff = cur - info['cost']
        else:
            label = "現物"
            diff = cur - info['cost']

        if info.get('currency') == "USD":
            p_usd = diff * shares
            p_jpy = p_usd * rate
            total_profit_usd_only_us_stocks += p_usd
        else:
            p_jpy = diff * shares
        # ---------------------------------

    total_profit_jpy += p_jpy

    cost_display = f"${info['cost']:,}" if info.get('currency') == "USD" else f"¥{info['cost']:,}"
    if shares == 0:
        cur_display = "-"
    elif price_available:
        cur_display = f"{('$' if info.get('currency') == 'USD' else '¥')}{cur:,.2f} {day_change_pct}"
    else:
        cur_display = "⚠️ 価格取得失敗"

    display_label = label if shares > 0 else "決済済"
    if shares > 0 and not price_available:
        display_label += "（未反映）"

    rows.append({
        "No.": i + 1, "銘柄": display_name, "数量": shares, "区分": display_label,
        "取得単価": cost_display, "現在値 (前日比)": cur_display,
        "損益(円)": f"¥{p_jpy:,.0f}" if (shares == 0 or price_available) else "¥0（未反映）"
    })

m_col1, m_col2 = st.columns(2)
m_col1.metric("総合計損益 (JPY)", f"¥{total_profit_jpy:,.0f}", delta=f"USD/JPY: {rate:.2f}")
m_col2.metric("米国株合計損益 (USD)", f"${total_profit_usd_only_us_stocks:,.2f}")

if rows: st.table(pd.DataFrame(rows))
else: st.info("銘柄がありません")

st.divider()

# 【改修２】重要スケジュールを複数行に表示して重なりを防止
if st.session_state.events:
    st.write("📌 **重要スケジュール**")
    
    # 1行あたり最大3つのイベントを表示する
    MAX_COLS = 3
    events = st.session_state.events
    
    for i in range(0, len(events), MAX_COLS):
        # 3つずつのチャンクに分ける
        chunk = events[i:i + MAX_COLS]
        cols = st.columns(MAX_COLS)
        
        for j, event in enumerate(chunk):
            try:
                target_date = datetime.strptime(event['date'], "%Y-%m-%d")
                days_left = (target_date - datetime.now()).days
                # フォントサイズの調整と重なり防止のため改行を含むマークダウンを使用
                cols[j].markdown(f"**{event['name']}**", unsafe_allow_html=True)
                cols[j].metric("", event['date'], f"あと {days_left} 日")
            except: pass

st.divider()
st.subheader("📋 Reminder")
st.info(st.session_state.reminder_text)