#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NBA Stats → Google Sheets automation
- 2025-26 Regular Season
- Last 10 games (overall; no location split)
- PerGame & Per100Possessions
- General (Advanced, Four Factors, Misc, Scoring, Opponent, Defense) + Clutch (Advanced, Four Factors, Misc, Scoring, Opponent)
- Writes to Google Sheet "NBA Model" (configurable)
- Run Log tab with outcomes and notes
- Sheets API write throttling to avoid 429s
- Optional proxy rotation for NBA API
- 6AM New York daily guard (bypass with RUN_ANYTIME=1)
"""

import os
import time
import random
import contextlib
import inspect
from datetime import datetime

import pandas as pd
import pytz

# =========================
# ENV / CONFIG
# =========================
SEASON           = os.getenv("SEASON", "2025-26")
SEASON_TYPE      = os.getenv("SEASON_TYPE", "Regular Season")
LAST_N_GAMES     = int(os.getenv("LAST_N_GAMES", "10"))
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "NBA Model")
# ADD THIS:
LEAGUE_ID        = os.getenv("LEAGUE_ID", "00")  # "00" = NBA only

CLUTCH_TIME      = os.getenv("CLUTCH_TIME", "Last 5 Minutes")
AHEAD_BEHIND     = os.getenv("AHEAD_BEHIND", "Ahead or Behind")
POINT_DIFF       = int(os.getenv("POINT_DIFF", "5"))

# Optional proxies (comma-separated ports in PROXY_PORTS)
PROXY_USER  = os.getenv("PROXY_USER", "")
PROXY_PASS  = os.getenv("PROXY_PASS", "")
PROXY_HOST  = os.getenv("PROXY_HOST", "")
PROXY_PORTS = [p.strip() for p in os.getenv("PROXY_PORTS", "").split(",") if p.strip()]

PER_MODES = ["PerGame", "Per100Possessions"]

GENERAL_MEASURE_TYPES = ["Advanced", "FourFactors", "Misc", "Scoring", "Opponent", "Defense"]
GENERAL_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
    "Defense": "Defense",
}

CLUTCH_MEASURE_TYPES = ["Advanced", "FourFactors", "Misc", "Scoring", "Opponent"]
CLUTCH_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
}

# Sheets write throttle to avoid 429 (“per minute” write quota)
SHEETS_WRITE_PAUSE_SEC = float(os.getenv("SHEETS_WRITE_PAUSE_SEC", "1.2"))
def sheets_pause():
    if SHEETS_WRITE_PAUSE_SEC > 0:
        time.sleep(SHEETS_WRITE_PAUSE_SEC)

# Run Log
RUN_LOG_SHEET     = os.getenv("RUN_LOG_SHEET", "Run_Log")
RUN_LOG_KEEP_LAST = int(os.getenv("RUN_LOG_KEEP_LAST", "5000"))
RUN_ANYTIME       = os.getenv("RUN_ANYTIME", "0")  # "1" bypasses 6am guard (useful for manual runs)
DEBUG_NBA_SIGNATURES = os.getenv("DEBUG_NBA_SIGNATURES", "0")  # "1" to print constructor params
NY_TZ = pytz.timezone("America/New_York")
RUN_LOG = []  # rows: [timestamp_ny, status, tab, note]

def log_result(tab: str, status: str, note: str = ""):
    ts = datetime.now(NY_TZ).strftime("%Y-%m-%d %H:%M:%S")
    RUN_LOG.append([ts, status, tab, (note or "")[:300]])

def sleep_backoff(attempt: int):
    time.sleep(0.9 * attempt + random.random() * 0.5)

# =========================
# GOOGLE SHEETS AUTH
# =========================
# Strip proxies from auth path to prevent weird failures
for k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"):
    os.environ.pop(k, None)

import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

sa_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
if not sa_path or not os.path.exists(sa_path):
    raise RuntimeError("GOOGLE_APPLICATION_CREDENTIALS not set or file not found.")

creds = Credentials.from_service_account_file(sa_path, scopes=SCOPES)
gc = gspread.authorize(creds)
print("✅ Google Sheets auth OK")

def get_sheet():
    try:
        return gc.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return gc.create(SPREADSHEET_NAME)

def write_df(sh, title: str, df: pd.DataFrame):
    sheets_pause()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        sheets_pause()
        ws = sh.add_worksheet(title=title, rows=max(len(df)+5, 50), cols=max(len(df.columns)+5, 26))
    sheets_pause()
    ws.clear()
    if df.empty:
        ws.update("A1", [["(no data)"]])
        print(f"⚠️ Wrote empty tab: {title}")
        return
    sheets_pause()
    set_with_dataframe(ws, df, include_index=False, include_column_header=True, resize=True)
    print(f"✅ Wrote tab: {title} ({len(df)} rows)")

def ensure_run_log_sheet(sh):
    try:
        sheets_pause()
        _ = sh.worksheet(RUN_LOG_SHEET)
    except gspread.WorksheetNotFound:
        sheets_pause()
        ws = sh.add_worksheet(title=RUN_LOG_SHEET, rows=200, cols=4)
        sheets_pause()
        ws.update("A1:D1", [["timestamp_ny", "status", "tab", "note"]])

def flush_run_log(sh):
    if not RUN_LOG:
        return
    try:
        sheets_pause()
        ws = sh.worksheet(RUN_LOG_SHEET)
        sheets_pause()
        ws.append_rows(RUN_LOG, value_input_option="RAW")
        # prune
        try:
            sheets_pause()
            vals = ws.get_all_values()
            total = len(vals) - 1
            if total > RUN_LOG_KEEP_LAST:
                to_delete = total - RUN_LOG_KEEP_LAST
                sheets_pause()
                ws.delete_rows(2, 1 + to_delete)
        except Exception:
            pass
    except Exception as e:
        print(f"⚠️ Run_Log write skipped: {e}")
    finally:
        RUN_LOG.clear()

# =========================
# NBA API SETUP
# =========================
from nba_api.stats.library.http import NBAStatsHTTP
from nba_api.stats.endpoints import leaguedashteamstats, leaguedashteamclutch

NBAStatsHTTP.HEADERS = {
    "Host": "stats.nba.com",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://www.nba.com/stats/",
    "Origin": "https://www.nba.com",
    "x-nba-stats-origin": "stats",
    "x-nba-stats-token": "true",
}

@contextlib.contextmanager
def use_proxy():
    """Context manager to apply one proxy port for a single request burst."""
    if PROXY_USER and PROXY_PASS and PROXY_HOST and PROXY_PORTS:
        port = random.choice(PROXY_PORTS)
        proxy = f"http://{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{port}"
        old_http = os.environ.get("HTTP_PROXY")
        old_https = os.environ.get("HTTPS_PROXY")
        os.environ["HTTP_PROXY"] = proxy
        os.environ["HTTPS_PROXY"] = proxy
        try:
            yield
        finally:
            if old_http is not None:
                os.environ["HTTP_PROXY"] = old_http
            else:
                os.environ.pop("HTTP_PROXY", None)
            if old_https is not None:
                os.environ["HTTPS_PROXY"] = old_https
            else:
                os.environ.pop("HTTPS_PROXY", None)
    else:
        yield


def dump_signatures():
    try:
        g_params = list(inspect.signature(leaguedashteamstats.LeagueDashTeamStats.__init__).parameters.keys())
        c_params = list(inspect.signature(leaguedashteamclutch.LeagueDashTeamClutch.__init__).parameters.keys())
        print("🧭 LeagueDashTeamStats params:", g_params)
        print("🧭 LeagueDashTeamClutch params:", c_params)
    except Exception as e:
        print("🧭 Param dump failed:", e)
        
def set_kw(kwargs: dict, cls, candidates, value) -> bool:
    """
    Try each candidate name against the class constructor.
    If a name exists in the signature, set kwargs[name] = value and return True.
    Otherwise return False.
    """
    params = set(inspect.signature(cls.__init__).parameters.keys())
    for name in candidates:
        if name in params:
            kwargs[name] = value
            return True
    return False
    
# =========================
# FETCHERS
# =========================
def fetch_general(per_mode: str, measure_label: str) -> pd.DataFrame:
    """
    LeagueDashTeamStats caller for this nba_api build.

    - Per-mode: per_mode_detailed ("PerGame" / "Per100Possessions")
    - Measure: measure_type_detailed_defense ("Advanced", "Four Factors", etc.)
    - Game scope: game_scope_simple_nullable="Last 10"  -> Last 10 games
    """
    api_measure = GENERAL_LABEL_TO_API[measure_label]

    for attempt in range(1, 8):
        try:
            with use_proxy():
                resp = leaguedashteamstats.LeagueDashTeamStats(
                    season=SEASON,
                    season_type_all_star=SEASON_TYPE,
                    per_mode_detailed=per_mode,
                    measure_type_detailed_defense=api_measure,
                    game_scope_simple_nullable="Last 10",  # <-- key change
                    pace_adjust="N",
                    plus_minus="N",
                    rank="N",
                )

                df_list = resp.get_data_frames()
                if not df_list:
                    raise RuntimeError("No data frames returned from LeagueDashTeamStats")

                df = df_list[0]
                if df is None or df.empty:
                    raise RuntimeError("Empty dataframe from LeagueDashTeamStats")

                # Strip non-NBA teams (WNBA etc.)
                if "TEAM_ID" in df.columns:
                    df = df[df["TEAM_ID"].astype(str).str.startswith("161061")]

                return df.reset_index(drop=True)

        except Exception as e:
            if attempt == 7:
                raise
            print(f"[general retry {attempt}] {measure_label} {per_mode}: {e}")
            sleep_backoff(attempt)


def fetch_clutch(per_mode: str, measure_label: str) -> pd.DataFrame:
    """
    LeagueDashTeamClutch caller for this specific nba_api build.

    Uses:
      - per_mode_detailed
      - measure_type_detailed_defense
      - last_n_games (for L10 clutch sample)
    """
    api_measure = CLUTCH_LABEL_TO_API[measure_label]

    for attempt in range(1, 8):
        try:
            with use_proxy():
                resp = leaguedashteamclutch.LeagueDashTeamClutch(
                    season=SEASON,
                    season_type_all_star=SEASON_TYPE,
                    per_mode_detailed=per_mode,                 # "PerGame" / "Per100Possessions"
                    measure_type_detailed_defense=api_measure,  # "Advanced", "Four Factors", etc.
                    clutch_time=CLUTCH_TIME,
                    ahead_behind=AHEAD_BEHIND,
                    point_diff=POINT_DIFF,
                    last_n_games=LAST_N_GAMES,
                    pace_adjust="N",
                    plus_minus="N",
                    rank="N",
                    # rest left default / nullable
                )

                df = resp.get_data_frames()[0]
                if df is None or df.empty:
                    raise RuntimeError("Empty dataframe from LeagueDashTeamClutch")

                if "TEAM_ID" in df.columns:
                    df = df[df["TEAM_ID"].astype(str).str.startswith("161061")]

                return df.reset_index(drop=True)

        except Exception as e:
            if attempt == 7:
                raise
            print(f"[clutch retry {attempt}] {measure_label} {per_mode}: {e}")
            sleep_backoff(attempt)



# =========================
# MAIN
# =========================
def main():
    sh = get_sheet()
    ensure_run_log_sheet(sh)

    # Optional: print constructor signatures to debug env (first run)
    if DEBUG_NBA_SIGNATURES == "1":
        dump_signatures()

    # 6–8 AM ET run window, unless manual trigger
    if RUN_ANYTIME != "1":
        now_et = datetime.now(NY_TZ)
        if not (6 <= now_et.hour < 8):  # allow 2-hour window
            note = f"Skip: NY time {now_et.strftime('%Y-%m-%d %H:%M:%S')}"
            print(note)
            log_result("RUN_GUARD", "SKIP", note)
            flush_run_log(sh)
            return

    # GENERAL
    for per_mode in PER_MODES:
        for label in GENERAL_MEASURE_TYPES:
            tab = f"NBA_GEN_{label}_{per_mode}_L{LAST_N_GAMES}"
            try:
                df = fetch_general(per_mode, label)
                write_df(sh, tab, df)
                log_result(tab, "OK", f"{len(df)} rows")
                sheets_pause()
            except Exception as e:
                print(f"❌ {tab} -> {e}")
                log_result(tab, "FAIL", str(e))

    # CLUTCH
    for per_mode in PER_MODES:
        for label in CLUTCH_MEASURE_TYPES:
            tab = f"NBA_CLUTCH_{label}_{per_mode}_L{LAST_N_GAMES}"
            try:
                df = fetch_clutch(per_mode, label)
                write_df(sh, tab, df)
                log_result(tab, "OK", f"{len(df)} rows")
                sheets_pause()
            except Exception as e:
                print(f"❌ {tab} -> {e}")
                log_result(tab, "FAIL", str(e))

    flush_run_log(sh)
    print("✅ Done")

if __name__ == "__main__":
    main()