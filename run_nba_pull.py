#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NBA Stats → Google Sheets automation
- Runs daily at 6AM America/New_York (GitHub Actions)
- Last 10 games (no location split)
- 2025-26 Regular Season
- Logs every tab result into Run_Log sheet
- Handles Google Sheets quotas (throttle)
- Proxy rotation + retry for NBA API
"""

import os, sys, time, random, contextlib, inspect
from datetime import datetime
import pandas as pd
import pytz

############################################
# CONFIG
############################################
SEASON           = os.getenv("SEASON", "2025-26")
SEASON_TYPE      = os.getenv("SEASON_TYPE", "Regular Season")
LAST_N_GAMES     = int(os.getenv("LAST_N_GAMES", "10"))
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "NBA Model")

CLUTCH_TIME      = os.getenv("CLUTCH_TIME", "Last 5 Minutes")
AHEAD_BEHIND     = os.getenv("AHEAD_BEHIND", "Ahead or Behind")
POINT_DIFF       = int(os.getenv("POINT_DIFF", "5"))

# PROXIES (optional)
PROXY_USER  = os.getenv("PROXY_USER", "")
PROXY_PASS  = os.getenv("PROXY_PASS", "")
PROXY_HOST  = os.getenv("PROXY_HOST", "")
PROXY_PORTS = [p for p in os.getenv("PROXY_PORTS", "").split(",") if p.strip()]

PER_MODES = ["PerGame", "Per100Possessions"]

GENERAL_MEASURE_TYPES = [
    "Advanced","FourFactors","Misc","Scoring","Opponent","Defense"
]
GENERAL_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
    "Defense": "Defense",
}

CLUTCH_MEASURE_TYPES = ["Advanced","FourFactors","Misc","Scoring","Opponent"]
CLUTCH_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
}

# SHEETS RATE LIMIT THROTTLE
SHEETS_WRITE_PAUSE_SEC = float(os.getenv("SHEETS_WRITE_PAUSE_SEC", "1.2"))

def sheets_pause():
    if SHEETS_WRITE_PAUSE_SEC > 0:
        time.sleep(SHEETS_WRITE_PAUSE_SEC)

# RUN LOG
RUN_LOG_SHEET     = os.getenv("RUN_LOG_SHEET", "Run_Log")
RUN_LOG_KEEP_LAST = int(os.getenv("RUN_LOG_KEEP_LAST", "5000"))
RUN_ANYTIME       = os.getenv("RUN_ANYTIME", "0")  # bypass 6AM guard for manual test
NY_TZ = pytz.timezone("America/New_York")
RUN_LOG = []

############################################
# LOGGING
############################################
def log_result(tab, status, note=""):
    ts = datetime.now(NY_TZ).strftime("%Y-%m-%d %H:%M:%S")
    RUN_LOG.append([ts, status, tab, (note or "")[:300]])

############################################
# GOOGLE AUTH
############################################
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# remove proxy env to not break auth
for k in ("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy"):
    os.environ.pop(k, None)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
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

############################################
# SHEETS HELPERS
############################################
def write_df(sh, title, df):
    sheets_pause()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        sheets_pause()
        ws = sh.add_worksheet(title=title, rows=max(len(df)+5,50), cols=max(len(df.columns)+5,26))
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
        ws.update("A1:D1", [["timestamp_ny","status","tab","note"]])

def flush_run_log(sh):
    if not RUN_LOG: 
        return
    try:
        sheets_pause()
        ws = sh.worksheet(RUN_LOG_SHEET)
        sheets_pause()
        ws.append_rows(RUN_LOG, value_input_option="RAW")
        # prune
        sheets_pause()
        vals = ws.get_all_values()
        total = len(vals)-1
        if total > RUN_LOG_KEEP_LAST:
            to_delete = total - RUN_LOG_KEEP_LAST
            sheets_pause()
            ws.delete_rows(2, 1+to_delete)
    except Exception as e:
        print(f"⚠️ Run_Log failed (ignored): {e}")
    finally:
        RUN_LOG.clear()

############################################
# NBA API SETUP
############################################
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
            if old_http is not None: os.environ["HTTP_PROXY"] = old_http
            else: os.environ.pop("HTTP_PROXY", None)
            if old_https is not None: os.environ["HTTPS_PROXY"] = old_https
            else: os.environ.pop("HTTPS_PROXY", None)
    else:
        yield

def sleep_backoff(n):
    time.sleep(0.9*n + random.random()*0.5)

def set_kw(kwargs: dict, cls, candidates, value):
    """Pick the first param name that exists on cls.__init__ and set it."""
    params = set(inspect.signature(cls.__init__).parameters.keys())
    for name in candidates:
        if name in params:
            kwargs[name] = value
            return True
    return False

############################################
# FETCHERS
############################################
def fetch_general(per_mode, measure_label):
    # Map our label (e.g., "FourFactors") to API phrase ("Four Factors")
    api_measure = GENERAL_LABEL_TO_API[measure_label]
    C = leaguedashteamstats.LeagueDashTeamStats

    for attempt in range(1, 8):
        try:
            with use_proxy():
                kwargs = {
                    "season": SEASON,
                    "pace_adjust": "N",
                    "plus_minus": "N",
                    "rank": "N",
                }

                # season_type / season_type_all_star
                set_kw(kwargs, C, ["season_type_all_star", "season_type"], SEASON_TYPE)

                # per_mode (detailed fallback)
                set_kw(kwargs, C, ["per_mode_detailed", "per_mode"], per_mode)

                # IMPORTANT: some nba_api builds ONLY expose measure_type_detailed_defense
                # so try it first even for non-defense, then fall back.
                candidates = ["measure_type_detailed_defense", "measure_type_detailed", "measure_type"]
                if not set_kw(kwargs, C, candidates, api_measure):
                    raise RuntimeError("No compatible measure_type kw found")

                # Last 10 overall
                if not set_kw(kwargs, C, ["last_n_games", "last_n_games_nullable"], LAST_N_GAMES):
                    set_kw(kwargs, C, ["game_scope", "game_scope_nullable"], "Last 10")

                resp = C(**kwargs)
                df = resp.get_data_frames()[0]
                if df is None or df.empty:
                    raise RuntimeError("Empty dataframe from API")
                return df

        except Exception as e:
            if attempt == 7:
                raise
            print(f"[general retry {attempt}] {api_measure} {per_mode}: {e}")
            sleep_backoff(attempt)


def fetch_clutch(per_mode, measure_label):
    api_measure = CLUTCH_LABEL_TO_API[measure_label]
    C = leaguedashteamclutch.LeagueDashTeamClutch

    for attempt in range(1, 8):
        try:
            with use_proxy():
                kwargs = {
                    "season": SEASON,
                    "clutch_time": CLUTCH_TIME,
                    "ahead_behind": AHEAD_BEHIND,
                    "point_diff": POINT_DIFF,
                    "pace_adjust": "N",
                    "plus_minus": "N",
                    "rank": "N",
                    # Many builds DO NOT accept game_scope → omit it.
                    "last_n_games": LAST_N_GAMES,
                }

                # season_type key differs by build
                set_kw(kwargs, C, ["season_type_all_star", "season_type"], SEASON_TYPE)

                # per_mode may be per_mode_time on this endpoint
                set_kw(kwargs, C, ["per_mode_time", "per_mode"], per_mode)

                # measure_type may be measure_type_time on this endpoint — try both, plus the defense variant just in case
                if not set_kw(kwargs, C, ["measure_type_time", "measure_type", "measure_type_detailed_defense"], api_measure):
                    raise RuntimeError("No compatible clutch measure_type kw found")

                resp = C(**kwargs)
                df = resp.get_data_frames()[0]
                if df is None or df.empty:
                    raise RuntimeError("Empty dataframe from clutch API")
                return df

        except Exception as e:
            if attempt == 7:
                raise
            print(f"[clutch retry {attempt}] {api_measure} {per_mode}: {e}")
            sleep_backoff(attempt)


############################################
# MAIN
############################################
def main():
    now = datetime.now(NY_TZ)
    sh = get_sheet()
    ensure_run_log_sheet(sh)

    # 6AM guard
    if RUN_ANYTIME != "1" and now.hour != 6:
        msg = f"Skip: NY time {now}"
        print(msg)
        log_result("RUN_GUARD", "SKIP", msg)
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
