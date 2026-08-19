import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime, timezone, timedelta, time as dt_time
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
import streamlit.components.v1 as components
import time
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- 0. データの保存・読み込み ---
DB_FILE = "portfolio.json"
EVENT_FILE = "events.json"
REMINDER_FILE = "reminder.json"
CONFIG_FILE = "config.json"

def now_jst():
    """
    サーバーの実行環境（Streamlit CloudなどはUTC）によらず、常に日本時間(JST)の
    naive datetimeを返す。datetime.now()の代わりに、表示・タイムスタンプ保存用途では
    こちらを使用する。
    """
    return datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(hours=9)

# 【改修箇所】対象のスプレッドシートURLをここに記述してください
FIXED_SHEET_URL = "https://docs.google.com/spreadsheets/d/17kAFl14q8EaaQ6kvezlAe1Yzr71Yo673T61--_cyESQ/edit"

# --- バックアップ履歴 設定 ---
MAX_BACKUPS = 1000          # 保持する最大バックアップ件数（超えたら古い順に削除）
INITIAL_CACHE_SIZE = 10     # 初回「1つ前の設定」で読み込む件数（直近N件）
WINDOW_CACHE_RADIUS = 10    # 範囲外アクセス時に読み込む前後件数（方向が分からない場合の対称フォールバック）
DIRECTIONAL_LOOKAHEAD = 30  # 進行方向側に多めに読み込む件数（同方向への連続移動を高速化する）
DIRECTIONAL_LOOKBEHIND = 3  # 進行方向と逆側は最小限だけ読み込む
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
def get_gspread_client(silent=False):
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Streamlit secrets または環境変数から認証情報を取得
        creds_dict = st.secrets["gcp_service_account"]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(credentials)
    except Exception as e:
        if not silent:
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
        keep_ts = expected_ts or actual_ts or now_jst().strftime(TIMESTAMP_FMT)
        ws.update(f"A{row_number}:B{row_number}", [[keep_ts, json.dumps(data, ensure_ascii=False)]])
        return keep_ts
    except Exception as e:
        st.error(f"上書き失敗: {e}")
        return None

def delete_backup_at_index(idx):
    """
    絶対インデックス idx の行をスプレッドシートから削除する。
    - 主：セッションに保持している idx（=行番号-1）から直接特定
    - 副：削除直前に実際のタイムスタンプとキャッシュ上のタイムスタンプを照合し、
          ズレていれば列A全体を検索してフォールバックする
    戻り値: 成功時は True、失敗時は False
    """
    gc = get_gspread_client()
    if not gc: return False
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
            col = ws.col_values(1)
            try:
                row_number = col.index(expected_ts) + 1
            except ValueError:
                st.error("削除対象のバックアップがシート上に見つかりませんでした（既に削除された可能性があります）。")
                return False

        ws.delete_rows(row_number)
        return True
    except Exception as e:
        st.error(f"削除失敗: {e}")
        return False

def export_to_spreadsheet(data):
    """新規バックアップを1行追記する（上書きしない）。1000件超は古い順に削除。"""
    gc = get_gspread_client()
    if not gc: return
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)

        timestamp = now_jst().strftime(TIMESTAMP_FMT)
        ws.append_row([timestamp, json.dumps(data, ensure_ascii=False)])

        total = get_backup_count(ws)
        if total > MAX_BACKUPS:
            excess = total - MAX_BACKUPS
            # 2行目（最古データ）から excess 件分をまとめて削除
            ws.delete_rows(2, 1 + excess)
            total = get_backup_count(ws)  # 削除後の正しい件数に更新

        # --- 重要：セッションが保持している先読みキャッシュ・総件数は
        # この時点で古くなっているため、次回のナビゲーション操作が
        # サーバーの最新状態を正しく参照できるようリセットしておく ---
        st.session_state.backup_total = total
        st.session_state.backup_cache = {}
        st.session_state.cache_min = None
        st.session_state.cache_max = None
        st.session_state.backup_index = None

        display_ts = datetime.strptime(timestamp, TIMESTAMP_FMT).strftime(DISPLAY_FMT)
        st.success(f"バックアップを保存しました（{display_ts}）")
    except Exception as e:
        st.error(f"エクスポート失敗: {e}")

def preload_backup_cache():
    """
    アプリ起動時に一度だけ呼び出す。直近 INITIAL_CACHE_SIZE 件を先読みして
    キャッシュしておくことで、初回の「1つ前の設定」操作をAPI呼び出しなしで
    即座に反映できるようにする。ライブの portfolio には一切手を加えない。
    """
    gc = get_gspread_client(silent=True)  # 未設定環境でも起動時にエラーを出さない
    if not gc:
        return
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)

        total = get_backup_count(ws)
        st.session_state.backup_total = total
        if total == 0:
            return

        start_idx = max(1, total - INITIAL_CACHE_SIZE + 1)
        cache, s, e = fetch_backup_range(ws, start_idx, total, total)
        if cache:
            st.session_state.backup_cache = cache
            st.session_state.cache_min = s
            st.session_state.cache_max = e
    except Exception:
        # 起動時のプリロードはあくまで先読み最適化のためのものなので、
        # 失敗してもアプリ本体の起動は妨げない（黙って諦める）
        pass

def load_backup_window(center_idx=None, initial=False, direction=None):
    """
    initial=True: 直近 INITIAL_CACHE_SIZE 件を読み込み、最新（=1つ前）を適用
    initial=False: center_idx を基準に読み込む。
        direction="next"（1つ後方向へ移動中）: 進行方向側に多め(DIRECTIONAL_LOOKAHEAD件)、
            逆側は最小限(DIRECTIONAL_LOOKBEHIND件)だけ読み込む
            → 同じ方向へ連続移動してもキャッシュ切れになりにくくする
        direction="prev"（1つ前方向へ移動中）: 上記の前後を逆にする
        direction=None: 前後対称（WINDOW_CACHE_RADIUS件ずつ）に読み込む

    既存のキャッシュと範囲が連続・重複する場合は、不足分だけをスプレッドシートから
    取得してマージする（境界付近を行き来した際に同じデータを再取得しないようにするため）。
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
            if direction == "next":
                start_idx = max(1, center_idx - DIRECTIONAL_LOOKBEHIND)
                end_idx = min(total, center_idx + DIRECTIONAL_LOOKAHEAD)
            elif direction == "prev":
                start_idx = max(1, center_idx - DIRECTIONAL_LOOKAHEAD)
                end_idx = min(total, center_idx + DIRECTIONAL_LOOKBEHIND)
            else:
                start_idx = max(1, center_idx - WINDOW_CACHE_RADIUS)
                end_idx = min(total, center_idx + WINDOW_CACHE_RADIUS)
            target_idx = center_idx

        existing_min = st.session_state.cache_min
        existing_max = st.session_state.cache_max

        if existing_min is not None and existing_max is not None \
                and start_idx <= existing_max + 1 and end_idx >= existing_min - 1:
            # 既存キャッシュと連続・重複 → 不足分だけ取得してマージする
            fetch_ranges = []
            if start_idx < existing_min:
                fetch_ranges.append((start_idx, existing_min - 1))
            if end_idx > existing_max:
                fetch_ranges.append((existing_max + 1, end_idx))
            new_min = min(existing_min, start_idx)
            new_max = max(existing_max, end_idx)
            merged_cache = st.session_state.backup_cache  # コピーせず直接追記する（キャッシュ肥大時のO(n)コピーを回避）
        else:
            # 既存キャッシュと繋がらない → その範囲を丸ごと取得
            fetch_ranges = [(start_idx, end_idx)]
            new_min, new_max = start_idx, end_idx
            merged_cache = {}

        for s_idx, e_idx in fetch_ranges:
            if s_idx > e_idx:
                continue
            cache, _, _ = fetch_backup_range(ws, s_idx, e_idx, total)
            if cache:
                merged_cache.update(cache)

        if not merged_cache:
            st.warning("該当する履歴データが見つかりませんでした")
            return

        st.session_state.backup_cache = merged_cache
        st.session_state.cache_min = new_min
        st.session_state.cache_max = new_max
        apply_backup_index(target_idx)
    except Exception as e:
        st.error(f"履歴取得失敗: {e}")

def apply_backup_index(idx):
    """
    キャッシュ済みの idx 番目のバックアップをセッションに反映する。
    閲覧中はディスクへの保存を行わない（毎回のボタン操作を軽くするため）。
    メモリ上の「復元」用スナップショット（backup_portfolio）は従来通り取っておく。
    """
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

def render_backup_nav_controls(key_prefix=""):
    """
    「🔄 スプレッドシートを再読み込み」「◀」「▶」の3ボタンを描画する。
    サイドバー・メイン画面など複数箇所から呼び出せるよう、ウィジェットキーの重複を避けるために
    key_prefix を付与する（例: "sidebar_", "main_"）。
    """
    if st.button("🔄 スプレッドシートを再読み込み", key=f"{key_prefix}reload_btn"):
        st.session_state.backup_cache = {}
        st.session_state.cache_min = None
        st.session_state.cache_max = None
        if st.session_state.backup_index is not None:
            # 現在表示中の位置を保ったまま、その周辺を最新の内容で読み直す
            load_backup_window(center_idx=st.session_state.backup_index)
        else:
            load_backup_window(initial=True)
        st.success("スプレッドシートを再読み込みしました")
        st.rerun()

    with st.container(key=f"{key_prefix}nav_arrows_row"):
        nav_col1, nav_col2 = st.columns(2)
        prev_clicked = nav_col1.button("◀", key=f"{key_prefix}prev_btn", use_container_width=True)
        next_clicked = nav_col2.button("▶", key=f"{key_prefix}next_btn", use_container_width=True)

    if prev_clicked:
        if st.session_state.backup_index is None:
            if st.session_state.cache_min is not None and st.session_state.backup_total:
                # 起動時にプリロード済みのキャッシュをそのまま使う（API呼び出し不要）
                apply_backup_index(st.session_state.backup_total)
            else:
                # プリロードが未実施・失敗していた場合のフォールバック
                load_backup_window(initial=True)
        else:
            target = st.session_state.backup_index - 1
            if target < 1:
                st.warning("これ以上前の履歴はありません")
            elif target < st.session_state.cache_min:
                load_backup_window(center_idx=target, direction="prev")
            else:
                apply_backup_index(target)
        st.rerun()

    if next_clicked:
        if st.session_state.backup_index is None:
            if st.session_state.cache_min is not None and st.session_state.backup_total:
                # 起動時にプリロード済みのキャッシュをそのまま使う（API呼び出し不要）
                apply_backup_index(st.session_state.backup_total)
            else:
                # プリロードが未実施・失敗していた場合のフォールバック
                load_backup_window(initial=True)
        else:
            total = st.session_state.backup_total or st.session_state.backup_index
            target = st.session_state.backup_index + 1
            if target > total:
                st.warning("これより新しい履歴はありません（最新のバックアップです）")
            elif target > st.session_state.cache_max:
                load_backup_window(center_idx=target, direction="next")
            else:
                apply_backup_index(target)
        st.rerun()

def jump_to_backup_index(target_idx):
    """
    指定した絶対インデックスへジャンプする。
    既にキャッシュ範囲内ならAPI呼び出しなしで即座に反映し、範囲外ならその周辺を取得する。
    """
    if st.session_state.backup_total is None or target_idx is None:
        st.warning("バックアップ履歴の総件数が不明です。先に「🔄 スプレッドシートを再読み込み」を押してください。")
        return
    target_idx = max(1, min(st.session_state.backup_total, int(target_idx)))
    if st.session_state.cache_min is not None and st.session_state.cache_max is not None \
            and st.session_state.cache_min <= target_idx <= st.session_state.cache_max:
        apply_backup_index(target_idx)
    else:
        load_backup_window(center_idx=target_idx)

def find_backup_index_by_date(target_date):
    """
    スプレッドシートのA列（タイムスタンプ）だけを読み、target_date（日付のみ、時刻は無視）に
    一致する最後のバックアップの絶対インデックスを返す。
    完全一致が無ければ、日付が一番近いものを返す。
    戻り値: (idx, actual_date) または (None, None)（失敗時・データなし）
    """
    gc = get_gspread_client()
    if not gc:
        return None, None
    try:
        sh = gc.open_by_url(FIXED_SHEET_URL)
        ws = sh.get_worksheet(0)
        ensure_header(ws)
        col = ws.col_values(1)  # ヘッダー込み。A列だけなので軽量
    except Exception:
        return None, None
    if len(col) <= 1:
        return None, None

    parsed = []  # [(idx, date), ...]
    for i, ts in enumerate(col[1:], start=1):
        try:
            dt = datetime.strptime(ts, TIMESTAMP_FMT)
            parsed.append((i, dt.date()))
        except Exception:
            continue
    if not parsed:
        return None, None

    # 同じ日付が複数あれば、一番新しい（その日の最後の）ものを選ぶ
    exact_matches = [p for p in parsed if p[1] == target_date]
    if exact_matches:
        return exact_matches[-1]

    # 完全一致が無ければ、日付が一番近いものにフォールバック
    return min(parsed, key=lambda p: abs((p[1] - target_date).days))

def render_backup_jump_controls(key_prefix=""):
    """
    「⏮ 最古 / ⏭ 最新」「No.でジャンプ」「日付でジャンプ」の3種類のジャンプ機能を描画する。
    バックアップ件数が数百〜千件規模になっても、目的の位置に少ないクリックで到達できるようにする。
    """
    if "_jump_msg" in st.session_state:
        st.info(st.session_state.pop("_jump_msg"))

    st.markdown("**バックアップへジャンプ**")

    # --- 最古 / 最新 ---
    jump_col1, jump_col2 = st.columns(2)
    if jump_col1.button("⏮ 最古", key=f"{key_prefix}jump_oldest_btn", use_container_width=True):
        if st.session_state.backup_total is None:
            load_backup_window(initial=True)
        jump_to_backup_index(1)
        st.rerun()
    if jump_col2.button("⏭ 最新", key=f"{key_prefix}jump_newest_btn", use_container_width=True):
        if st.session_state.backup_total is None:
            load_backup_window(initial=True)
        else:
            jump_to_backup_index(st.session_state.backup_total)
        st.rerun()

    if not st.session_state.backup_total:
        st.caption("履歴を一度読み込むと、No.・日付でのジャンプも使えるようになります。")
        return

    # --- No.直接指定でジャンプ ---
    no_col1, no_col2 = st.columns([2, 1])
    max_no = st.session_state.backup_total
    default_no = min(st.session_state.backup_index or 1, max_no)
    jump_no = no_col1.number_input(
        "No.でジャンプ", min_value=1, max_value=max_no, value=default_no,
        step=1, key=f"{key_prefix}jump_no_input", label_visibility="collapsed"
    )
    if no_col2.button("移動", key=f"{key_prefix}jump_no_btn", use_container_width=True):
        jump_to_backup_index(int(jump_no))
        st.rerun()

    # --- 日付でジャンプ ---
    date_col1, date_col2 = st.columns([2, 1])
    jump_date = date_col1.date_input(
        "日付でジャンプ", value=now_jst().date(),
        key=f"{key_prefix}jump_date_input", label_visibility="collapsed"
    )
    if date_col2.button("検索", key=f"{key_prefix}jump_date_btn", use_container_width=True):
        idx, actual_date = find_backup_index_by_date(jump_date)
        if idx is None:
            st.warning("該当するバックアップが見つかりませんでした")
        else:
            jump_to_backup_index(idx)
            if actual_date != jump_date:
                st.session_state["_jump_msg"] = (
                    f"指定日にはデータが無かったため、最も近い"
                    f"{actual_date.strftime('%Y年%m月%d日')}のバックアップへ移動しました"
                )
        st.rerun()

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
if 'twelvedata_key' not in st.session_state:
    st.session_state.twelvedata_key = load_json(CONFIG_FILE, {}).get("twelvedata_key", "")
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
if 'startup_backup_preloaded' not in st.session_state:
    st.session_state.startup_backup_preloaded = False

if not st.session_state.startup_backup_preloaded:
    # アプリ起動時に一度だけ、直近のバックアップ履歴を先読みしてキャッシュしておく
    preload_backup_cache()
    st.session_state.startup_backup_preloaded = True

# 復元用バックアップ
def backup_portfolio():
    st.session_state.prev_portfolio = copy.deepcopy(st.session_state.portfolio)

# --- 2. API設定 ---
current_api_key = st.session_state.api_key or st.secrets.get("GEMINI_API_KEY", "")
if current_api_key:
    genai.configure(api_key=current_api_key)

current_twelvedata_key = st.session_state.twelvedata_key or st.secrets.get("TWELVEDATA_API_KEY", "")

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

def _read_csv_text(uploaded_file):
    """アップロードされたCSVファイルをUTF-8-SIG/Shift_JIS両対応でテキストとして読み込む"""
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('shift_jis', errors='ignore')

def detect_csv_type(text):
    """
    コメント行（先頭'#'）を除いた最初のヘッダー行を見て、CSVの種類を判定する。
    'holdings': 保有株数・取得単価を含む保有明細形式
    'trend'   : code,name のみのトレンド銘柄リスト形式
    'unknown' : 判定できない形式
    """
    lines = [l for l in text.splitlines() if l.strip() and not l.lstrip().startswith('#')]
    if not lines:
        return 'unknown'
    header = lines[0]
    if ('保有株数' in header or '数量' in header) and '取得単価' in header:
        return 'holdings'
    header_lower = header.lower()
    if 'code' in header_lower and 'name' in header_lower:
        return 'trend'
    return 'unknown'

def parse_stock_csv(uploaded_file):
    """
    1行目が '#' で始まるコメント行（例: # トレンド銘柄...,保存日時:...,市場:jp）の場合はスキップし、
    'code,name' ヘッダーを持つCSVから銘柄コード・銘柄名のリストを抽出する。
    戻り値: [(code, name), ...]
    """
    text = _read_csv_text(uploaded_file)
    return parse_stock_csv_text(text)

def parse_stock_csv_text(text):
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

def parse_holdings_csv_text(text):
    """
    '銘柄コード,銘柄名,保有株数,取得単価（加重平均）,現在値,評価額,評価損益' のような
    保有明細形式のCSVから、銘柄コード・銘柄名・保有株数・取得単価を抽出する。
    列名の表記ゆれ（「数量」等）にもある程度対応する。
    戻り値: [(code, name, shares, cost), ...]
    """
    lines = [l for l in text.splitlines() if l.strip() and not l.lstrip().startswith('#')]
    if not lines:
        return []

    reader = csv.DictReader(lines)
    fieldnames = reader.fieldnames or []

    def find_field(*keywords):
        # キーワードを優先度順（＝より具体的なものを先）に、全列に対して試す。
        # フィールド優先で回すと、緩いキーワードが別項目の列に誤マッチすることがあるため注意。
        for kw in keywords:
            for fn in fieldnames:
                name = (fn or '').strip()
                if kw in name:
                    return fn
        return None

    code_field = find_field('銘柄コード', 'コード', 'code')
    name_field = find_field('銘柄名', '名称', 'name')
    shares_field = find_field('保有株数', '数量', 'shares')
    cost_field = find_field('取得単価', 'cost')

    results = []
    for row in reader:
        code = (row.get(code_field) or '').strip() if code_field else ''
        if not code:
            continue
        name = (row.get(name_field) or '').strip() if name_field else ''
        shares_raw = (row.get(shares_field) or '').strip() if shares_field else ''
        cost_raw = (row.get(cost_field) or '').strip() if cost_field else ''
        try:
            shares = float(shares_raw.replace(',', '')) if shares_raw else 0.0
        except ValueError:
            shares = 0.0
        try:
            cost = float(cost_raw.replace(',', '')) if cost_raw else 0.0
        except ValueError:
            cost = 0.0
        results.append((code, name, shares, cost))
    return results

def build_entries_from_holdings_csv(rows):
    """
    [(code, name, shares, cost), ...] から portfolio 用のエントリ辞書を作る。
    株数・取得単価はCSVの値をそのまま使うため、価格取得のAPI呼び出しは発生しない。
    区分は「現物」固定。
    """
    entries = {}
    for code, name, shares, cost in rows:
        is_japan = code.isdigit() and len(code) == 4
        currency = "JPY" if is_japan else "USD"
        key = f"{code}_現物"
        entries[key] = {
            "name": name or code,
            "shares": shares,
            "cost": round(cost, 2),
            "currency": currency
        }
    return entries

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

TWELVEDATA_BASE_URL = "https://api.twelvedata.com"

def build_twelvedata_symbol(portfolio_key):
    """内部の portfolio キー（例: '7203_現物'）から Twelve Data 用のシンボル文字列を作る"""
    code = portfolio_key.split('_')[0]
    is_japan = code.isdigit() and len(code) == 4
    if code == "IHI":
        return "7013:TSE"
    if is_japan:
        return f"{code}:TSE"
    return code.upper()

def fetch_twelvedata_quotes(portfolio_keys, api_key):
    """
    Twelve Data の /quote エンドポイントで複数銘柄をまとめて取得する。
    戻り値: {portfolio_key: {"current":..., "prev_close":...}} （取得できたものだけを含む）
    取得できなかった銘柄はこの戻り値に含まれないので、呼び出し元でYahooにフォールバックする。
    """
    if not api_key or not portfolio_keys:
        return {}

    key_to_symbol = {k: build_twelvedata_symbol(k) for k in portfolio_keys}
    symbol_to_keys = {}
    for k, sym in key_to_symbol.items():
        symbol_to_keys.setdefault(sym, []).append(k)

    symbols = list(symbol_to_keys.keys())
    results = {}

    CHUNK = 100  # Twelve Dataは1リクエストあたり最大120銘柄まで対応（余裕を持たせる）
    for i in range(0, len(symbols), CHUNK):
        chunk_symbols = symbols[i:i + CHUNK]
        params = {"symbol": ",".join(chunk_symbols), "apikey": api_key}
        try:
            resp = requests.get(f"{TWELVEDATA_BASE_URL}/quote", params=params, timeout=10)
            data = resp.json()
        except Exception:
            continue  # このチャンクは全滅 → 呼び出し元でYahooにフォールバックされる

        # 単一銘柄の場合は結果がそのままdictで返り、複数銘柄の場合はsymbolをキーにしたdict of dictで返る
        if isinstance(data, dict) and "symbol" in data:
            entries = {data.get("symbol"): data}
        elif isinstance(data, dict):
            entries = data
        else:
            entries = {}

        for sym, entry in entries.items():
            if not isinstance(entry, dict):
                continue
            if entry.get("status") == "error" or "close" not in entry:
                continue
            try:
                cur = float(entry["close"])
            except (TypeError, ValueError):
                continue
            prev = entry.get("previous_close")
            try:
                prev = float(prev) if prev is not None else None
            except (TypeError, ValueError):
                prev = None

            matched_keys = symbol_to_keys.get(sym)
            if not matched_keys:
                # レスポンスのsymbol表記が微妙に異なる場合（例: "7203:TSE" vs "7203"）への保険
                base = str(sym).split(':')[0]
                for cand_sym, ks in symbol_to_keys.items():
                    if cand_sym.split(':')[0] == base:
                        matched_keys = ks
                        break
            if matched_keys:
                for k in matched_keys:
                    results[k] = {"current": cur, "prev_close": prev}

    return results

def _fetch_single_symbol_price(key, td_api_key):
    """
    1銘柄分の価格を、優先順位（Twelve Data → .info → metadata → fast_info → history）で取得する。
    この関数自体はキャッシュされない「生」の取得処理。time.sleepもここに置くことで、
    キャッシュ経由の呼び出し（下の _fetch_single_symbol_price_cached）がヒットした場合は
    実行されず、実際にネットワークへ取りに行った時だけ待機するようにする。
    戻り値: (price_dict または None, debug_msg, error_msg)
    """
    symbol = key.split('_')[0]
    is_japan = symbol.isdigit() and len(symbol) == 4
    ticker = f"{symbol}.T" if is_japan else ("7013.T" if symbol == "IHI" else symbol)
    error_msg = ""

    # --- 0. Twelve Data（設定されている場合のみ） ---
    if td_api_key:
        try:
            td_result = fetch_twelvedata_quotes([key], td_api_key)
            if key in td_result:
                r = td_result[key]
                time.sleep(0.1)
                return r, f"[twelvedata] 現在値={r['current']} 前日終値={r['prev_close']}", ""
        except Exception as e:
            error_msg += f" / [twelvedata] {type(e).__name__}: {e}"

    got_price = False
    result = None
    debug_msg = ""

    # --- 1. .info（v7/finance/quote エンドポイント）を最優先で試す ---
    # (get_history_metadata / fast_info / history は、実は全て同じ「チャート用API」を
    #  内部で共有しており、見た目は別ルートでも実体は同一データだった。
    #  .info は quoteSummary + v7/finance/quote という完全に別のAPIを叩くため、
    #  ここでようやく本当に独立した気配値が期待できる。ただし重い呼び出しなので
    #  失敗時は下位のフォールバックに切り替える)
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        cur = info.get("regularMarketPrice") if info else None
        if cur is None and info:
            cur = info.get("currentPrice")
        prev = info.get("regularMarketPreviousClose") if info else None
        if prev is None and info:
            prev = info.get("previousClose")
        if cur is not None and not pd.isna(cur):
            result = {"current": cur, "prev_close": prev}
            debug_msg = f"[info] 現在値={cur} 前日終値={prev}"
            got_price = True
    except Exception as e:
        error_msg += f" / [info] {type(e).__name__}: {e}"

    # --- 2. .info が使えない場合は get_history_metadata() の regularMarketPrice を試す ---
    if not got_price:
        try:
            stock = yf.Ticker(ticker)
            md = stock.get_history_metadata()
            cur = md.get("regularMarketPrice") if md else None
            prev = md.get("chartPreviousClose") if md else None
            if prev is None and md:
                prev = md.get("previousClose")
            if cur is not None and not pd.isna(cur):
                result = {"current": cur, "prev_close": prev}
                debug_msg = f"[metadata] 現在値={cur} 前日終値={prev}"
                got_price = True
        except Exception as e:
            error_msg += f" / [metadata] {type(e).__name__}: {e}"

    if not got_price:
        # --- 3. metadataも使えない場合は fast_info を試す ---
        try:
            stock = yf.Ticker(ticker)
            fi = stock.fast_info
            cur = fi.get("lastPrice") if fi else None
            prev = fi.get("previousClose") if fi else None
            if cur is not None and not pd.isna(cur):
                result = {"current": cur, "prev_close": prev}
                debug_msg = f"[fast_info] 現在値={cur:.2f} 前日終値={prev}"
                got_price = True
        except Exception as e:
            error_msg += f" / [fast_info] {type(e).__name__}: {e}"

    if not got_price:
        # --- 4. それでもダメなら従来の日足取得にフォールバック ---
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period="5d")
            if not hist.empty:
                last_close = hist['Close'].iloc[-1]
                last_dt = hist.index[-1]
                result = {
                    "current": last_close,
                    "prev_close": hist['Close'].iloc[-2] if len(hist) >= 2 else None
                }
                debug_msg = f"[history フォールバック] {last_dt} 終値={last_close:.2f}（{len(hist)}本取得）"
                got_price = True
            else:
                error_msg += " / 空のデータが返されました（レート制限の可能性）"
        except Exception as e:
            error_msg += f" / [history] {type(e).__name__}: {e}"

    time.sleep(0.3)  # 実際にネットワークへ取りに行った場合のみ待機（レート制限対策）
    return result, debug_msg, error_msg.strip(" /")

def _fetch_usdjpy_rate(td_api_key):
    """USD/JPY為替レートを取得する（Twelve Data優先、ダメならYahooにフォールバック）"""
    if td_api_key:
        try:
            resp = requests.get(f"{TWELVEDATA_BASE_URL}/quote", params={"symbol": "USD/JPY", "apikey": td_api_key}, timeout=10)
            fx_data = resp.json()
            if isinstance(fx_data, dict) and fx_data.get("status") != "error" and "close" in fx_data:
                return float(fx_data["close"]), ""
        except Exception as e:
            pass  # 下のフォールバックへ

    try:
        usdjpy_md = yf.Ticker("JPY=X").get_history_metadata()
        rate = usdjpy_md.get("regularMarketPrice", 159.2) if usdjpy_md else 159.2
        return rate, ""
    except Exception:
        try:
            usdjpy = yf.Ticker("JPY=X").history(period="5d")
            rate = usdjpy['Close'].iloc[-1] if not usdjpy.empty else 159.2
            return rate, ""
        except Exception as e:
            return 159.2, f"{type(e).__name__}: {e}"

PRICE_CACHE_TTL_SECONDS = 900  # 15分

@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def _fetch_single_symbol_price_cached(key, td_api_key):
    """
    銘柄1つ単位で価格をキャッシュする。
    バックアップ履歴を移動して保有銘柄構成が変わっても、共通する銘柄はこのキャッシュを
    使い回せるため、銘柄セット全体をキーにしていた以前の方式より無駄な再取得が大幅に減る。
    """
    return _fetch_single_symbol_price(key, td_api_key)

@st.cache_data(ttl=PRICE_CACHE_TTL_SECONDS, show_spinner=False)
def _fetch_usdjpy_rate_cached(td_api_key):
    return _fetch_usdjpy_rate(td_api_key)

def get_live_prices(portfolio_keys, td_api_key=None):
    prices = {}
    fetch_errors = {}   # デバッグ用：銘柄ごとの失敗理由を記録
    fetch_debug = {}    # デバッグ用：成功時も含め、取得元・値を記録
    portfolio_keys = list(portfolio_keys)

    if portfolio_keys:
        # 銘柄ごとの取得はネットワークI/O待ちが支配的なので、並列化して大幅に高速化する。
        # （st.cache_dataでヒットする銘柄は即座に返るが、未キャッシュの銘柄が多いバックアップに
        #  切り替えた際、逐次処理だと銘柄数×待機時間の分だけ直列に遅くなっていたため）
        max_workers = min(8, len(portfolio_keys))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_key = {
                executor.submit(_fetch_single_symbol_price_cached, key, td_api_key): key
                for key in portfolio_keys
            }
            for future in as_completed(future_to_key):
                key = future_to_key[future]
                try:
                    result, debug_msg, error_msg = future.result()
                except Exception as e:
                    result, debug_msg, error_msg = None, "", f"{type(e).__name__}: {e}"
                prices[key] = result
                if debug_msg:
                    fetch_debug[key] = debug_msg
                if error_msg:
                    fetch_errors[key] = error_msg

    rate, rate_error = _fetch_usdjpy_rate_cached(td_api_key)
    prices["USDJPY"] = rate
    if rate_error:
        fetch_errors["USDJPY"] = rate_error

    prices["_fetch_errors"] = fetch_errors
    prices["_fetch_debug"] = fetch_debug
    return prices

def get_prices_with_cache(portfolio_keys, td_api_key=None):
    """
    ポートフォリオのキー集合から価格を取得する（銘柄単位で15分キャッシュ）。
    戻り値: (prices_dict, last_updated_datetime)
    """
    keys_tuple = tuple(sorted(portfolio_keys))
    if not keys_tuple:
        return {"USDJPY": 159.2}, now_jst()
    prices = get_live_prices(keys_tuple, td_api_key=td_api_key)
    return prices, now_jst()

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
    /* 仮想注文パネル：PC幅（769px以上）のときだけ画面右下に固定表示する。
       スマホ幅では固定を解除し、通常のページの流れに沿って表示することで
       コンテンツに重ならないようにする */
    @media (min-width: 769px) {
        div[class*="st-key-floating_sim_panel"] {
            position: fixed;
            bottom: 20px;
            right: 20px;
            width: 320px;
            background-color: #0e1117;
            border: 1px solid #444;
            border-radius: 10px;
            padding: 14px 16px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.6);
            z-index: 9999;
        }
    }
    /* ◀▶ナビゲーションボタンは、画面幅が狭くても常に横並びのままにする
       （Streamlitの列は既定で狭い画面だと縦積みに折り返されるため、それを無効化する） */
    div[class*="nav_arrows_row"] [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
    }
    div[class*="nav_arrows_row"] [data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width: 0 !important;
        width: 50% !important;
        flex: 1 1 0 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- キーボードショートカット（← → で バックアップ履歴を前後に移動） ---
# 親ドキュメント側の __backupNavKeyListenerInstalled フラグにより、実際のリスナー登録は
# 初回成功時の1回だけに抑えられる。そのため、Python側で「1回だけ描画」に絞り込む最適化は
# 撤回する（1回目の描画・インストールが何らかの理由で失敗すると、以降ずっと矢印キーが
# 反応しなくなる不具合につながっていたため）。
components.html("""
<script>
(function() {
    function tryInstall() {
        let doc;
        try {
            doc = window.parent.document;
        } catch (err) {
            return false; // 親フレームがまだ準備できていない → 後でリトライ
        }
        if (!doc) { return false; }
        // 複数回のrerunや複数回のリトライでリスナーが重複登録されるのを防ぐためのフラグ
        if (doc.__backupNavKeyListenerInstalled) { return true; }

        function clickButtonByText(text) {
            const buttons = Array.from(doc.querySelectorAll('button'));
            const matches = buttons.filter(btn => btn.innerText.trim() === text);
            if (matches.length === 0) { return; }
            // サイドバー・メイン画面に同名ボタンが重複しているため、
            // 非表示（折りたたみ等で offsetParent が null）のものは避け、見えている方を優先する
            const visible = matches.find(btn => btn.offsetParent !== null);
            (visible || matches[0]).click();
        }

        doc.addEventListener('keydown', function(e) {
            const active = doc.activeElement;
            const tag = active ? active.tagName.toLowerCase() : '';
            // 入力欄にフォーカスがある間は矢印キー本来の動作（カーソル移動等）を妨げない
            if (tag === 'input' || tag === 'textarea' || (active && active.isContentEditable)) {
                return;
            }
            if (e.key === 'ArrowRight') {
                clickButtonByText('▶');
                e.preventDefault();
            } else if (e.key === 'ArrowLeft') {
                clickButtonByText('◀');
                e.preventDefault();
            }
        });

        doc.__backupNavKeyListenerInstalled = true;
        return true;
    }

    // 起動直後は親フレームの準備が間に合わずインストールに失敗することがあるため、
    // 成功するまで短い間隔でリトライする（最大 約5秒間）
    if (!tryInstall()) {
        let attempts = 0;
        const maxAttempts = 20; // 250ms × 20 = 5秒
        const retryTimer = setInterval(function() {
            attempts++;
            if (tryInstall() || attempts >= maxAttempts) {
                clearInterval(retryTimer);
            }
        }, 250);
    }
})();
</script>
""", height=0)

with st.sidebar:
    st.header("🔑 Settings")
    new_api_key = st.text_input("Gemini API Key", value=st.session_state.api_key, type="password")
    if st.button("APIキーを保存"):
        st.session_state.api_key = new_api_key
        cfg = load_json(CONFIG_FILE, {})
        cfg["gemini_key"] = new_api_key
        save_json(CONFIG_FILE, cfg)
        st.success("APIキーを保存しました")
        st.rerun()

    new_td_key = st.text_input("Twelve Data API Key", value=st.session_state.twelvedata_key, type="password",
                                help="株価取得の第一候補として使用します。未設定・取得失敗時はYahoo Financeにフォールバックします。")
    if st.button("Twelve Data APIキーを保存"):
        st.session_state.twelvedata_key = new_td_key
        cfg = load_json(CONFIG_FILE, {})
        cfg["twelvedata_key"] = new_td_key
        save_json(CONFIG_FILE, cfg)
        st.success("Twelve Data APIキーを保存しました")
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

    st.markdown('<hr style="margin: 0.5rem 0;">', unsafe_allow_html=True)
    st.subheader("💾 Backup (Spreadsheet)")
    full_config = {"portfolio": st.session_state.portfolio, "events": st.session_state.events, "reminder_text": st.session_state.reminder_text}

    @st.dialog("確認")
    def confirm_overwrite_dialog(data_to_write, target_idx):
        st.write("データを上書きしますか？")
        col_ok, col_cancel = st.columns(2)
        if col_ok.button("OK", use_container_width=True):
            written_ts = overwrite_backup_at_index(target_idx, data_to_write)
            if written_ts:
                entry = st.session_state.backup_cache.get(target_idx)
                if entry:
                    entry["data"] = data_to_write
                try:
                    display_ts = datetime.strptime(written_ts, TIMESTAMP_FMT).strftime(DISPLAY_FMT)
                except Exception:
                    display_ts = written_ts
                st.session_state["_overwrite_success_msg"] = f"バックアップを上書きしました（{display_ts}）"
            st.rerun()
        if col_cancel.button("Cancel", use_container_width=True):
            st.rerun()

    exp_col1, exp_col2 = st.columns(2)
    new_backup_clicked = exp_col1.button("🆕 新規バックアップ")
    overwrite_clicked = exp_col2.button("♻️ 上書きバックアップ")

    if "_overwrite_success_msg" in st.session_state:
        st.success(st.session_state.pop("_overwrite_success_msg"))

    if new_backup_clicked:
        export_to_spreadsheet(full_config)

    if overwrite_clicked:
        if st.session_state.backup_index is None:
            st.warning("上書き対象が選択されていません。先に「◀ 1つ前の設定」で対象のバックアップを表示してください。")
        else:
            confirm_overwrite_dialog(full_config, st.session_state.backup_index)

    # --- キャッシュ強制更新・前後ナビゲーション ---
    render_backup_nav_controls(key_prefix="sidebar_")

    # --- ジャンプ機能（件数が多くなっても目的のバックアップへ素早く到達できるようにする） ---
    render_backup_jump_controls(key_prefix="sidebar_")

    # --- 現在表示中のバックアップ情報を表示（保存No. / 保存日付） ---
    if st.session_state.backup_index is not None:
        entry = st.session_state.backup_cache.get(st.session_state.backup_index)
        if entry:
            try:
                dt = datetime.strptime(entry["timestamp"], TIMESTAMP_FMT)
                display_ts = dt.strftime(DISPLAY_FMT)
            except Exception:
                display_ts = entry["timestamp"]
            total_disp = st.session_state.backup_total or "?"
            st.info(f"📌 保存No.: **{st.session_state.backup_index} / {total_disp}**\n\n"
                    f"📅 保存日付: **{display_ts}**")
    else:
        st.caption("バックアップ履歴はまだ読み込まれていません（「◀ 1つ前の設定」を押すと表示されます）")

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
            try:
                text = _read_csv_text(csv_file)
                csv_type = detect_csv_type(text)

                if csv_type == 'holdings':
                    with st.spinner("CSVを解析中（保有株数・取得単価をファイルから反映）..."):
                        rows = parse_holdings_csv_text(text)
                        if not rows:
                            st.warning("CSVから銘柄を読み取れませんでした。フォーマットをご確認ください。")
                        else:
                            new_entries = build_entries_from_holdings_csv(rows)
                            backup_portfolio()
                            if csv_mode == "現在の画面に追加":
                                st.session_state.portfolio.update(new_entries)
                            else:
                                st.session_state.portfolio = new_entries
                            save_json(DB_FILE, st.session_state.portfolio)
                            st.success(f"{len(new_entries)}件の銘柄を読み込みました（保有株数・取得単価をファイルから反映）")
                            st.rerun()

                elif csv_type == 'trend':
                    with st.spinner("CSVを解析し、現在株価を取得中..."):
                        rows = parse_stock_csv_text(text)
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
                            st.success(f"{len(new_entries)}件の銘柄を読み込みました（数量100・現在株価で登録）")
                            st.rerun()

                else:
                    st.warning("CSVの形式を判定できませんでした。「code,name」形式、"
                               "または「銘柄コード,銘柄名,保有株数,取得単価...」形式に対応しています。")
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

# 15分（900,000ミリ秒）ごとに自動で画面を再実行する
st_autorefresh(interval=PRICE_CACHE_TTL_SECONDS * 1000, key="auto_price_refresh")

def fetch_open_price_for_date(code, currency, reference_date):
    """
    基準日(reference_date)を基準に始値を取得する。
    基準日の日足があればその始値、無ければ基準日以前で直近の日足の始値を使う
    （基準日が休日や、基準日=当日でまだ開場前の場合等に自動的に直近の値へフォールバックする）。
    取得できなければNoneを返す。
    """
    ticker = f"{code}.T" if (currency == "JPY" and code != "IHI") else ("7013.T" if code == "IHI" else code.upper())
    try:
        hist = yf.Ticker(ticker).history(period="1mo")  # 基準日が7日前でも余裕を持って1ヶ月分取得
        if hist.empty:
            return None
        hist = hist.copy()
        hist['_date'] = hist.index.date
        candidates = hist[hist['_date'] <= reference_date]
        if candidates.empty:
            return None
        return float(candidates['Open'].iloc[-1])
    except Exception:
        return None

def fetch_close_price_for_date(code, currency, reference_date):
    """
    基準日を基準に終値を取得する。
    ・基準日が「今日」の場合：今日の終値はまだ確定していないとみなし、前営業日以前の直近の確定終値を使う
    ・基準日が過去日の場合：その日（無ければ基準日以前で直近の過去営業日）の終値を使う
    取得できなければNoneを返す。
    """
    ticker = f"{code}.T" if (currency == "JPY" and code != "IHI") else ("7013.T" if code == "IHI" else code.upper())
    today_jst = now_jst().date()
    try:
        hist = yf.Ticker(ticker).history(period="1mo")
        if hist.empty:
            return None
        hist = hist.copy()
        hist['_date'] = hist.index.date
        if reference_date >= today_jst:
            candidates = hist[hist['_date'] < today_jst]  # 当日はまだ未確定なので除外
        else:
            candidates = hist[hist['_date'] <= reference_date]
        if candidates.empty:
            return None
        return float(candidates['Close'].iloc[-1])
    except Exception:
        return None

def fetch_mid_price_for_date(code, reference_date):
    """
    基準日を基準に日本株の「仲値」（前場の終値＝前引け、11:30頃の株価）を取得する。
    Yahoo Financeの1分足は直近7日分しか取得できないため、基準日は7日前までに限定される前提。
    ・基準日が「今日」かつ実際の現在時刻がまだ11:30前の場合：今日はスキップし、直近の過去営業日を使う
    ・それ以外：基準日（無ければ基準日以前で直近の過去営業日）の11:30以前の最後の足を採用
      （その日に11:30以前の足が無ければ、その日の最後の足で代用）
    取得できなければNoneを返す。
    """
    ticker = f"{code}.T" if code != "IHI" else "7013.T"
    try:
        hist = yf.Ticker(ticker).history(period="7d", interval="1m")
        if hist.empty:
            return None

        idx = hist.index
        if idx.tz is not None:
            idx_jst = idx.tz_convert('Asia/Tokyo')
        else:
            idx_jst = idx.tz_localize('UTC').tz_convert('Asia/Tokyo')
        hist = hist.copy()
        hist.index = idx_jst

        today_jst = now_jst().date()
        now_time_jst = now_jst().time()
        all_dates = sorted(set(hist.index.date))
        candidate_dates = sorted([d for d in all_dates if d <= reference_date], reverse=True)

        for d in candidate_dates:
            if d == today_jst and now_time_jst < dt_time(11, 30):
                continue  # 当日はまだ前引け前 → この日はスキップして前の営業日へ
            day_data = hist[hist.index.date == d]
            morning_day = day_data[day_data.index.time <= dt_time(11, 30)]
            if not morning_day.empty:
                return float(morning_day['Close'].iloc[-1])
            if not day_data.empty:
                return float(day_data['Close'].iloc[-1])

        return None
    except Exception:
        return None

def simulate_equal_investment(total_amount, input_currency, prices_dict, rate, order_mode="即注文", reference_date=None):
    """
    現在のポートフォリオ銘柄に、投資金額を均等配分する。
    - 投資額はJPY基準に変換した上で銘柄数で等分する
    - 各銘柄は自国通貨に変換し、約定価格で1株単位（切り捨て）の株数を算出する
    - order_mode: "即注文"（現在値、reference_dateは無視）/ "始値注文"（基準日の始値）/
      "仲値注文"（日本株のみ、基準日の前場終値＝前引け値）/ "終値注文"（基準日の終値）
      いずれも、基準日にデータが無い場合は基準日以前で直近の過去営業日に自動フォールバックする。
      "仲値注文"で米国株の場合は現在値にフォールバックする。
    戻り値: (success, {key: {"shares":..., "exec_price":...}} または None, 必要最低金額(入力通貨換算) または 0)
    """
    keys = list(st.session_state.portfolio.keys())
    if not keys:
        return False, None, 0

    if reference_date is None:
        reference_date = now_jst().date()

    # 現在価格を取得できた銘柄のみを対象にする（約定価格の基準として現在値の生存確認に使う）
    valid_entries = {}
    for key in keys:
        info = st.session_state.portfolio[key]
        p_data = prices_dict.get(key)
        cur = p_data.get("current") if p_data else None
        if cur is None or pd.isna(cur):
            continue

        currency = info.get("currency", "JPY")
        code = key.split('_')[0]
        exec_price = float(cur)  # デフォルト＝即注文

        if order_mode == "始値注文":
            open_price = fetch_open_price_for_date(code, currency, reference_date)
            if open_price is not None and not pd.isna(open_price):
                exec_price = open_price
            # 取得できない場合は現在値のままフォールバック
        elif order_mode == "終値注文":
            close_price = fetch_close_price_for_date(code, currency, reference_date)
            if close_price is not None and not pd.isna(close_price):
                exec_price = close_price
            # 取得できない場合は現在値のままフォールバック
        elif order_mode == "仲値注文":
            if currency == "JPY":
                mid = fetch_mid_price_for_date(code, reference_date)
                if mid is not None:
                    exec_price = mid
                # 仲値が取れない場合は現在値のままフォールバック
            # 米国株は仲値注文非対応のため現在値のままフォールバック

        valid_entries[key] = {"price": exec_price, "currency": currency}

    if not valid_entries:
        return False, None, 0

    n_valid = len(valid_entries)
    total_jpy = total_amount if input_currency == "JPY" else total_amount * rate
    equal_per_stock_jpy = total_jpy / n_valid

    # 各銘柄の価格をJPY換算し、最も高い銘柄を特定（＝均等配分の上での制約条件）
    max_price_jpy = 0
    for e in valid_entries.values():
        price_jpy = e["price"] * rate if e["currency"] == "USD" else e["price"]
        max_price_jpy = max(max_price_jpy, price_jpy)

    if equal_per_stock_jpy < max_price_jpy:
        min_required_jpy = max_price_jpy * n_valid
        min_required_display = min_required_jpy if input_currency == "JPY" else min_required_jpy / rate
        return False, None, min_required_display

    result = {}
    for key, e in valid_entries.items():
        equal_amount_native = equal_per_stock_jpy if e["currency"] == "JPY" else equal_per_stock_jpy / rate
        shares = int(equal_amount_native // e["price"])  # 1株単位（切り捨て）
        result[key] = {"shares": shares, "exec_price": e["price"]}

    return True, result, 0

col_refresh, col_ts, col_quick_nav = st.columns([1, 2, 1.3])
if col_refresh.button('最新価格に更新'):
    # キャッシュを明示的に破棄してから再実行（手動更新は必ず最新値を取りに行く）
    _fetch_single_symbol_price_cached.clear()
    _fetch_usdjpy_rate_cached.clear()
    st.rerun()

prices_dict, last_updated = get_prices_with_cache(st.session_state.portfolio.keys(), td_api_key=current_twelvedata_key)
col_ts.caption(f"🕒 最終更新: {last_updated.strftime('%Y年%m月%d日 %H:%M:%S')}（15分ごとに自動更新）")
rate = prices_dict.get("USDJPY", 159.2)

with col_quick_nav:
    render_backup_nav_controls(key_prefix="main_")

    if st.session_state.backup_index is not None and st.session_state.backup_total:
        position_str = f"{st.session_state.backup_index}/{st.session_state.backup_total}"
    else:
        position_str = "-/-"
    # Streamlitのcolumnsは狭い画面で縦積みになるため、位置番号とリマインダー本文は
    # 生のHTML(flex)で横並びにする（①のナビボタンと同じ理由）
    reminder_html = (st.session_state.reminder_text or "").replace("\n", "<br>")
    st.markdown(f"""
    <div style="display:flex; align-items:flex-start; gap:10px;
                background-color:rgba(28,131,225,0.1); border:1px solid rgba(28,131,225,0.4);
                border-radius:8px; padding:10px 14px; margin-top:4px;">
        <div style="font-weight:bold; white-space:nowrap; color:#4dabf7; flex-shrink:0;">{position_str}</div>
        <div style="flex:1; word-break:break-word;">{reminder_html}</div>
    </div>
    """, unsafe_allow_html=True)

if "_delete_success_msg" in st.session_state:
    st.success(st.session_state.pop("_delete_success_msg"))

rows = []
row_keys = []  # rows[i] に対応する実際の portfolio キー（削除対象の特定に使う）
total_profit_jpy = 0
total_profit_usd_only_us_stocks = 0
total_cost_basis_jpy = 0  # 含み損益率(%)算出用：価格取得できた銘柄の取得金額合計（円換算）

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
            total_cost_basis_jpy += info['cost'] * shares * rate
        else:
            p_jpy = diff * shares
            total_cost_basis_jpy += info['cost'] * shares
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

    row_keys.append(key)
    rows.append({
        "選択": False,
        "No.": i + 1, "銘柄": display_name, "数量": shares, "区分": display_label,
        "取得単価": cost_display, "現在値 (前日比)": cur_display,
        "損益(円)": p_jpy if (shares == 0 or price_available) else 0
    })

m_col1, m_col2 = st.columns(2)
with m_col1:
    sub_amount, sub_pct = st.columns([2, 1])
    sub_amount.metric("総合計損益 (JPY)", f"¥{total_profit_jpy:,.0f}", delta=f"USD/JPY: {rate:.2f}")
    if total_cost_basis_jpy != 0:
        total_pl_pct = total_profit_jpy / total_cost_basis_jpy * 100
        sub_pct.metric("含み損益率", f"{total_pl_pct:+.2f}%")
    else:
        sub_pct.metric("含み損益率", "—")
m_col2.metric("米国株合計損益 (USD)", f"${total_profit_usd_only_us_stocks:,.2f}")

if rows:
    df_display = pd.DataFrame(rows)
    edited_df = st.data_editor(
        df_display,
        column_config={
            "選択": st.column_config.CheckboxColumn("選択", default=False, width="small"),
            "損益(円)": st.column_config.NumberColumn("損益(円)", format="¥%,.0f"),
        },
        disabled=["No.", "銘柄", "数量", "区分", "取得単価", "現在値 (前日比)", "損益(円)"],
        hide_index=True,
        use_container_width=True,
        key="portfolio_table_editor"
    )
    selected_keys = [row_keys[idx] for idx, checked in enumerate(edited_df["選択"]) if checked]

    @st.dialog("確認")
    def confirm_delete_stocks_dialog(keys_to_delete):
        st.write(f"以下の{len(keys_to_delete)}銘柄を削除しますか？")
        for k in keys_to_delete:
            nm = st.session_state.portfolio.get(k, {}).get("name", "")
            st.write(f"- {k.split('_')[0]} {nm}")
        st.caption("削除後も「復元」ボタンで直前の状態に戻せます。")
        col_ok, col_cancel = st.columns(2)
        if col_ok.button("OK", use_container_width=True):
            backup_portfolio()
            for k in keys_to_delete:
                st.session_state.portfolio.pop(k, None)
            save_json(DB_FILE, st.session_state.portfolio)
            st.session_state["_delete_stocks_success_msg"] = f"{len(keys_to_delete)}銘柄を削除しました"
            st.rerun()
        if col_cancel.button("Cancel", use_container_width=True):
            st.rerun()

    if st.button("🗑️ 銘柄削除", disabled=(len(selected_keys) == 0)):
        confirm_delete_stocks_dialog(selected_keys)

    if "_delete_stocks_success_msg" in st.session_state:
        st.success(st.session_state.pop("_delete_stocks_success_msg"))
else:
    st.info("銘柄がありません")

# --- シート削除（現在表示中のバックアップと対応するスプレッドシートのデータを削除） ---
@st.dialog("確認")
def confirm_delete_sheet_dialog(target_idx):
    st.write("現在表示中のバックアップと、対応するスプレッドシートのデータを削除しますか？")
    st.caption("この操作は元に戻せません。")
    col_ok, col_cancel = st.columns(2)
    if col_ok.button("OK", use_container_width=True):
        deleted = delete_backup_at_index(target_idx)
        if deleted:
            # ローカルの表示状態・キャッシュを破棄し、live編集用のportfolioに戻す
            st.session_state.portfolio = load_json(DB_FILE, {})
            st.session_state.events = load_json(EVENT_FILE, [])
            st.session_state.reminder_text = load_json(REMINDER_FILE, "- ターゲット日程を入力してください")
            st.session_state.backup_cache = {}
            st.session_state.cache_min = None
            st.session_state.cache_max = None
            st.session_state.backup_total = None
            st.session_state.backup_index = None
            st.session_state["_delete_success_msg"] = "削除しました。データを再読み込みしました。"
        st.rerun()
    if col_cancel.button("Cancel", use_container_width=True):
        st.rerun()

if st.button("🗑️ シート削除"):
    if st.session_state.backup_index is None:
        st.warning("削除対象が選択されていません。先に「◀ 1つ前の設定」で対象のバックアップを表示してください。")
    else:
        confirm_delete_sheet_dialog(st.session_state.backup_index)

# --- 仮想注文パネル（PC幅では画面右下に固定表示、スマホ幅では通常フロー） ---
with st.container(key="floating_sim_panel"):
    with st.expander("💰 仮想注文", expanded=False):
        sim_amt_col, sim_cur_col = st.columns([2, 1])
        sim_amount = sim_amt_col.number_input(
            "投資金額", min_value=0.0, value=0.0, step=1000.0,
            key="sim_amount", label_visibility="collapsed", placeholder="投資金額"
        )
        sim_currency_label = sim_cur_col.selectbox(
            "単位", ["円", "ドル"], key="sim_currency", label_visibility="collapsed"
        )
        sim_currency = "JPY" if sim_currency_label == "円" else "USD"

        sim_order_mode = st.radio(
            "注文方法", ["即注文", "始値注文", "仲値注文", "終値注文"], index=0,
            key="sim_order_mode", horizontal=True
        )
        if sim_order_mode == "仲値注文":
            st.caption("※ 仲値注文は日本株のみ対応です（米国株は現在値で約定します）")

        def business_days_back(base_date, n):
            """base_dateからn営業日（土日を除く。日本の祝日は考慮しない簡易版）だけさかのぼった日付を返す"""
            d = base_date
            count = 0
            while count < n:
                d -= timedelta(days=1)
                if d.weekday() < 5:  # 0=月曜〜4=金曜
                    count += 1
            return d

        sim_min_date = business_days_back(now_jst().date(), 7)
        sim_reference_date = st.date_input(
            "基準日",
            value=now_jst().date(),
            min_value=sim_min_date,
            max_value=now_jst().date(),
            key="sim_reference_date",
            disabled=(sim_order_mode == "即注文"),
            help="始値注文・仲値注文・終値注文の基準となる日付（直近7営業日まで）。即注文では使用しません。"
        )
        if sim_order_mode == "仲値注文" and sim_reference_date < now_jst().date() - timedelta(days=7):
            st.caption("⚠️ 仲値注文は暦日ベースで直近7日分のデータしか取得できないため、"
                       "この基準日では取得できず現在値にフォールバックする可能性があります。")

        sim_btn_col1, sim_btn_col2 = st.columns(2)
        virtual_order_clicked = sim_btn_col1.button("仮想注文", use_container_width=True)
        revert_clicked = sim_btn_col2.button("戻す", use_container_width=True)

        if virtual_order_clicked:
            if not sim_amount or sim_amount <= 0:
                st.warning("投資金額を入力してください")
            else:
                success, result, min_needed = simulate_equal_investment(
                    sim_amount, sim_currency, prices_dict, rate,
                    order_mode=sim_order_mode, reference_date=sim_reference_date
                )
                if success:
                    backup_portfolio()
                    for key, r in result.items():
                        st.session_state.portfolio[key]['shares'] = r['shares']
                        st.session_state.portfolio[key]['cost'] = round(r['exec_price'], 2)
                    save_json(DB_FILE, st.session_state.portfolio)
                    ref_disp = sim_reference_date.strftime('%Y年%m月%d日') if sim_order_mode != "即注文" else ""
                    st.success(f"{len(result)}銘柄に均等配分しました（{sim_order_mode}{' ' + ref_disp if ref_disp else ''}）")
                    st.rerun()
                else:
                    unit = "円" if sim_currency == "JPY" else "ドル"
                    st.error(f"投資金額が不足しています。最低 {min_needed:,.0f}{unit} 必要です。")

        if revert_clicked:
            if st.session_state.prev_portfolio is not None:
                st.session_state.portfolio = copy.deepcopy(st.session_state.prev_portfolio)
                st.session_state.prev_portfolio = None
                save_json(DB_FILE, st.session_state.portfolio)
                st.success("仮想注文前の状態に戻しました")
                st.rerun()
            else:
                st.error("戻せる状態がありません")

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
                days_left = (target_date - now_jst()).days
                # フォントサイズの調整と重なり防止のため改行を含むマークダウンを使用
                cols[j].markdown(f"**{event['name']}**", unsafe_allow_html=True)
                cols[j].metric("", event['date'], f"あと {days_left} 日")
            except: pass

st.divider()
# --- 価格取得エラーの診断表示（原因調査用・画面最下部） ---
st.divider()
_fetch_errors = prices_dict.get("_fetch_errors", {})
_fetch_debug = prices_dict.get("_fetch_debug", {})
if _fetch_errors:
    with st.expander(f"⚠️ 価格取得に失敗した銘柄があります（{len(_fetch_errors)}件）"):
        for err_key, err_msg in _fetch_errors.items():
            st.caption(f"**{err_key}**: {err_msg}")
if _fetch_debug:
    with st.expander("🔍 取得データの詳細（デバッグ用）", expanded=False):
        st.caption("表示中の「終値」がYahoo Financeから返ってきた最新の値です。"
                   "「最新価格に更新」を数分あけて複数回押し、ここの日時・値が動くか確認してください。")
        for dbg_key, dbg_msg in _fetch_debug.items():
            st.caption(f"**{dbg_key}**: {dbg_msg}")