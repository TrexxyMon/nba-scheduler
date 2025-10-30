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
        return
    sheets_pause()
    set_with_dataframe(ws, df, include_index=False, include_column_header=True, resize=True)

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
    "User-Agent": "Mozilla/5.0",
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
        try: yield
        finally:
            if old_http: os.environ["HTTP_PROXY"]=old_http
            else: os.environ.pop("HTTP_PROXY",None)
            if old_https: os.environ["HTTPS_PROXY"]=old_https
            else: os.environ.pop("HTTPS_PROXY",None)
    else:
        yield

def sleep_backoff(n):
    time.sleep(0.9*n + random.random()*0.5)

def kw_last10(cls, n):
    params = inspect.signature(cls.__init__).parameters
    if "last_n_games" in params:
        return {"last_n_games": n}
    if "game_scope" in params:
        return {"game_scope": "Last 10"}
    return {}

############################################
# FETCHERS
############################################
def fetch_general(per_mode, measure):
    C = leaguedashteamstats.LeagueDashTeamStats
    for attempt in range(1,8):
        try:
            with use_proxy():
                resp = C(
                    season=SEASON,
                    pace_adjust="N", plus_minus="N", rank="N",
                    **{"season_type_all_star":SEASON_TYPE} if "season_type_all_star" in inspect.signature(C.__init__).parameters else {"season_type":SEASON_TYPE},
                    **({"per_mode_detailed":per_mode} if "per_mode_detailed" in inspect.signature(C.__init__).parameters else {"per_mode":per_mode}),
                    **({"measure_type_detailed":measure} if "measure_type_detailed" in inspect.signature(C.__init__).parameters else {"measure_type":measure}),
                    **kw_last10(C,LAST_N_GAMES)
                )
                return resp.get_data_frames()[0]
        except Exception as e:
            if attempt==7: raise
            print(f"[general retry {attempt}] {measure} {per_mode}: {e}")
            sleep_backoff(attempt)

def fetch_clutch(per_mode, measure):
    C = leaguedashteamclutch.LeagueDashTeamClutch
    for attempt in range(1,8):
        try:
            with use_proxy():
                resp = C(
                    season=SEASON,
                    season_type=SEASON_TYPE,
                    clutch_time=CLUTCH_TIME,
                    ahead_behind=AHEAD_BEHIND,
                    point_diff=POINT_DIFF,
                    pace_adjust="N", plus_minus="N", rank="N",
                    game_scope="Last 10", last_n_games=LAST_N_GAMES,
                    per_mode=per_mode, measure_type=measure
                )
                return resp.get_data_frames()[0]
        except Exception as e:
            if attempt==7: raise
            print(f"[clutch retry {attempt}] {measure} {per_mode}: {e}")
            sleep_backoff(attempt)

############################################
# MAIN
############################################
def main():
    now = datetime.now(NY_TZ)
    sh = get_sheet()
    ensure_run_log_sheet(sh)

    # 6AM guard
    if RUN_ANYTIME!="1" and now.hour != 6:
        msg = f"Skip: NY time {now}"
        print(msg)
        log_result("RUN_GUARD","SKIP",msg)
        flush_run_log(sh)
        return

    # GENERAL
    for per_mode in PER_MODES:
        for label in GENERAL_MEASURE_TYPES:
            api_measure = GENERAL_LABEL_TO_API[label]
            tab = f"NBA_GEN_{label}_{per_mode}_L{LAST_N_GAMES}"
            try:
                df = fetch_general(per_mode, api_measure)
                write_df(sh, tab, df)
                log_result(tab,"OK",f"{len(df)} rows")
            except Exception as e:
                log_result(tab,"FAIL",str(e))

    # CLUTCH
    for per_mode in PER_MODES:
        for label in CLUTCH_MEASURE_TYPES:
            api_measure = CLUTCH_LABEL_TO_API[label]
            tab = f"NBA_CLUTCH_{label}_{per_mode}_L{LAST_N_GAMES}"
            try:
                df = fetch_clutch(per_mode, api_measure)
                write_df(sh, tab, df)
                log_result(tab,"OK",f"{len(df)} rows")
            except Exception as e:
                log_result(tab,"FAIL",str(e))

    flush_run_log(sh)
    print("✅ Done")

if __name__ == "__main__":
    main()
