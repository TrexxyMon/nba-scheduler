#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NBA Stats → Google Sheets (Daily @ 6AM America/New_York via GitHub Actions)
- Season: 2025-26 (env overrideable)
- Window: Last 10 games overall (no location split)
- Pulls:
    * General: Advanced, Four Factors, Misc, Scoring, Opponent, Defense
    * Clutch:  Advanced, Four Factors, Misc, Scoring, Opponent
    * PerMode: PerGame, Per100Possessions
- Features:
    * Google Service Account auth (GOOGLE_APPLICATION_CREDENTIALS)
    * Proxy rotation + retry/backoff
    * Robust parameter handling across nba_api versions
    * Overwrites tabs
    * Run Log tab (append-only + auto-trim)
    * 6AM New York guard with manual-bypass (RUN_ANYTIME=1)
"""

import os, random, time, contextlib, inspect, sys
from datetime import datetime
import pandas as pd
import pytz

# ---------- Config (override via env in Actions) ----------
SEASON             = os.getenv("SEASON", "2025-26")
SEASON_TYPE        = os.getenv("SEASON_TYPE", "Regular Season")
LAST_N_GAMES       = int(os.getenv("LAST_N_GAMES", "10"))
SPREADSHEET_NAME   = os.getenv("SPREADSHEET_NAME", "NBA Model")

CLUTCH_TIME        = os.getenv("CLUTCH_TIME", "Last 5 Minutes")
AHEAD_BEHIND       = os.getenv("AHEAD_BEHIND", "Ahead or Behind")  # must be this for overall
POINT_DIFF         = int(os.getenv("POINT_DIFF", "5"))

# Proxies (optional; leave blank to skip)
PROXY_USER         = os.getenv("PROXY_USER", "")
PROXY_PASS         = os.getenv("PROXY_PASS", "")
PROXY_HOST         = os.getenv("PROXY_HOST", "gate.decodo.com")
PROXY_PORTS        = [int(p) for p in os.getenv(
    "PROXY_PORTS",
    "10001,10002,10003,10004,10005,10006,10007,10008,10009,10010"
).split(",") if p.strip()]

PER_MODES = ["PerGame", "Per100Possessions"]

GENERAL_MEASURE_TYPES = ["Advanced","FourFactors","Misc","Scoring","Opponent","Defense"]
GENERAL_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
    "Defense": "Defense",
}
CLUTCH_MEASURE_TYPES  = ["Advanced","FourFactors","Misc","Scoring","Opponent"]
CLUTCH_LABEL_TO_API = {
    "Advanced": "Advanced",
    "FourFactors": "Four Factors",
    "Misc": "Misc",
    "Scoring": "Scoring",
    "Opponent": "Opponent",
}

# ---------- Run Log helpers ----------
RUN_LOG_SHEET = os.getenv("RUN_LOG_SHEET", "Run_Log")
RUN_LOG_KEEP_LAST = int(os.getenv("RUN_LOG_KEEP_LAST", "5000"))
RUN_ANYTIME = os.getenv("RUN_ANYTIME", "0")  # '1' bypasses the 6am guard

NY_TZ = pytz.timezone("America/New_York")
RUN_LOG = []

def log_result(tab: str, status: str, note: str = ""):
    ts = datetime.now(NY_TZ).strftime("%Y-%m-%d %H:%M:%S")
    RUN_LOG.append([ts, status, tab, (note or "")[:300]])

# ---------- Google Sheets auth ----------
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# Ensure no stray proxies leak into Google auth
for k in ("HTTP_PROXY","HTTPS_PROXY","http_proxy","https_proxy"):
    os.environ.pop(k, None)

# Explicit scopes are REQUIRED in CI (Drive scope needed for open(name) lookup)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

sa_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
if not sa_path or not os.path.exists(sa_path):
    raise RuntimeError("GOOGLE_APPLICATION_CREDENTIALS not set or file not found.")

creds = Credentials.from_service_account_file(sa_path, scopes=SCOPES)
gc = gspread.authorize(creds)
print("✅ Google Sheets auth OK (scoped)")


def get_or_create_spreadsheet(name: str):
    try:
        return gc.open(name)
    except gspread.SpreadsheetNotFound:
        return gc.create(name)

def write_df(sh, title: str, df: pd.DataFrame):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=max(len(df)+5, 50), cols=max(len(df.columns)+5, 26))
    ws.clear()
    if df.empty:
        ws.update("A1", [["(no data)"]])
        print(f"⚠️ Wrote empty tab: {title}")
        return
    set_with_dataframe(ws, df, include_index=False, include_column_header=True, resize=True)
    print(f"✅ Wrote tab: {title} ({len(df)} rows)")

def flush_run_log(sh):
    if not RUN_LOG:
        return
    try:
        try:
            ws = sh.worksheet(RUN_LOG_SHEET)
            new_ws = False
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=RUN_LOG_SHEET, rows=100, cols=4)
            new_ws = True
        if new_ws:
            ws.update("A1:D1", [["timestamp_ny", "status", "tab", "note"]])
        ws.append_rows(RUN_LOG, value_input_option="RAW")
        # prune to last N
        try:
            existing = ws.get_all_values()
            total = len(existing) - 1
            if total > RUN_LOG_KEEP_LAST:
                to_delete = total - RUN_LOG_KEEP_LAST
                ws.delete_rows(2, 1 + to_delete)
        except Exception:
            pass
    finally:
        RUN_LOG.clear()

# ---------- NBA headers + proxy/backoff ----------
from nba_api.stats.library.http import NBAStatsHTTP

NBAStatsHTTP.HEADERS = {
    "Host": "stats.nba.com",
    "Connection": "keep-alive",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Origin": "https://www.nba.com",
    "Referer": "https://www.nba.com/stats/",
    "x-nba-stats-origin": "stats",
    "x-nba-stats-token": "true",
    "Pragma": "no-cache",
    "Cache-Control": "no-cache",
}
print("🛠️ NBA headers set")

@contextlib.contextmanager
def use_proxy_port(port: int):
    if PROXY_USER and PROXY_PASS and PROXY_HOST and port:
        proxy = f"http://{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{port}"
        prev_http = os.environ.get("HTTP_PROXY")
        prev_https = os.environ.get("HTTPS_PROXY")
        os.environ["HTTP_PROXY"]  = proxy
        os.environ["HTTPS_PROXY"] = proxy
        try:
            yield
        finally:
            if prev_http is None: os.environ.pop("HTTP_PROXY", None)
            else: os.environ["HTTP_PROXY"] = prev_http
            if prev_https is None: os.environ.pop("HTTPS_PROXY", None)
            else: os.environ["HTTPS_PROXY"] = prev_https
    else:
        yield

def sleep_backoff(attempt: int, base: float = 0.9, jitter: float = 0.6):
    time.sleep(base*attempt + random.random()*jitter)

# ---------- General fetcher (Last 10 overall) ----------
from nba_api.stats.endpoints import leaguedashteamstats

def _choose_kw(cls, candidates):
    params = set(inspect.signature(cls.__init__).parameters.keys())
    for k in candidates:
        if k in params: return k
    return None

def _kwargs_last10_overall(cls, n: int):
    kw = {}
    all_kw  = _choose_kw(cls, ["last_n_games", "last_n_games_nullable"])
    game_scope_kw = _choose_kw(cls, ["game_scope", "game_scope_nullable"])
    if all_kw:
        kw[all_kw] = n
        return kw
    if game_scope_kw:
        kw[game_scope_kw] = "Last 10"
        return kw
    return kw

def fetch_leaguedashteamstats_last10_overall(
    season: str, season_type: str, per_mode: str, measure_type: str,
    timeout: int = 60, max_retries: int = 10
) -> pd.DataFrame:
    Stats = leaguedashteamstats.LeagueDashTeamStats
    per_mode_kw = _choose_kw(Stats, ["per_mode_detailed", "per_mode"])
    measure_kw  = _choose_kw(Stats, ["measure_type_detailed", "measure_type_detailed_defense", "measure_type"])
    season_type_kw = _choose_kw(Stats, ["season_type_all_star", "season_type"])
    for attempt in range(1, max_retries+1):
        try:
            with use_proxy_port(random.choice(PROXY_PORTS) if PROXY_PORTS else None):
                kwargs = {
                    "season": season,
                    "pace_adjust": "N",
                    "plus_minus": "N",
                    "rank": "N",
                    "timeout": timeout,
                }
                if season_type_kw: kwargs[season_type_kw] = season_type
                if per_mode_kw:    kwargs[per_mode_kw]    = per_mode
                if measure_kw:     kwargs[measure_kw]     = measure_type
                kwargs.update(_kwargs_last10_overall(Stats, LAST_N_GAMES))
                resp = Stats(**kwargs)
                df = resp.get_data_frames()[0]
                if df is None or df.empty:
                    raise RuntimeError("Empty dataframe from API")
                return df
        except Exception as e:
            print(f"[general retry {attempt}] {measure_type} {per_mode}: {e}")
            if attempt == max_retries:
                raise
            sleep_backoff(attempt)

# ---------- Clutch fetcher (Last 10 overall) ----------
from nba_api.stats.endpoints import leaguedashteamclutch as clutch_ep

def _has_param(cls, name: str) -> bool:
    return name in inspect.signature(cls.__init__).parameters

def fetch_leaguedashteamclutch_last10_overall(
    season: str,
    season_type: str,
    per_mode: str,
    measure_type: str,
    clutch_time: str,
    ahead_behind: str,   # "Ahead or Behind"
    point_diff: int,
    max_retries: int = 10,
    timeout: int = 60
) -> pd.DataFrame:
    C = clutch_ep.LeagueDashTeamClutch
    desired = {
        "season": season,
        "clutch_time": clutch_time,
        "ahead_behind": ahead_behind,
        "point_diff": point_diff,
        "pace_adjust": "N",
        "plus_minus": "N",
        "rank": "N",
        "game_scope": "Last 10",
        "last_n_games": LAST_N_GAMES,
        "season_type": season_type,
        "per_mode_time": per_mode,
        "per_mode": per_mode,
        "measure_type_time": measure_type,
        "measure_type": measure_type,
    }
    kwargs = {k: v for k, v in desired.items() if _has_param(C, k)}
    for attempt in range(1, max_retries+1):
        try:
            with use_proxy_port(random.choice(PROXY_PORTS) if PROXY_PORTS else None):
                resp = C(timeout=timeout, **kwargs)
            df = resp.get_data_frames()[0]
            if df is None or df.empty:
                raise RuntimeError("Empty dataframe from clutch API")
            return df
        except Exception as e:
            print(f"[clutch retry {attempt}] {measure_type} {per_mode}: {e}")
            if attempt == max_retries:
                raise
            sleep_backoff(attempt)

# ---------- Main ----------
def main():
    sh = get_or_create_spreadsheet(SPREADSHEET_NAME)

    # 6AM guard AFTER auth so we can log SKIP and create Run_Log
    now_ny = datetime.now(NY_TZ)
    if RUN_ANYTIME != "1" and now_ny.hour != 6:
        msg = f"NY time {now_ny.strftime('%Y-%m-%d %H:%M:%S')} not 06:00 — skipping run"
        print(msg)
        log_result("RUN_GUARD", "SKIP", msg)
        flush_run_log(sh)
        return

    try:
        # General
        for per_mode in PER_MODES:
            for label in GENERAL_MEASURE_TYPES:
                api_measure = GENERAL_LABEL_TO_API[label]
                tab = f"NBA_GEN_{label}_{per_mode}_L{LAST_N_GAMES}"
                try:
                    df = fetch_leaguedashteamstats_last10_overall(
                        season=SEASON,
                        season_type=SEASON_TYPE,
                        per_mode=per_mode,
                        measure_type=api_measure
                    )
                    write_df(sh, tab, df)
                    log_result(tab, "OK", f"{len(df)} rows")
                except Exception as e:
                    print("❌", tab, "->", e)
                    log_result(tab, "FAIL", str(e))

        # Clutch
        for per_mode in PER_MODES:
            for label in CLUTCH_MEASURE_TYPES:
                api_measure = CLUTCH_LABEL_TO_API[label]
                tab = f"NBA_CLUTCH_{label}_{per_mode}_L{LAST_N_GAMES}"
                try:
                    df = fetch_leaguedashteamclutch_last10_overall(
                        season=SEASON,
                        season_type=SEASON_TYPE,
                        per_mode=per_mode,
                        measure_type=api_measure,
                        clutch_time=CLUTCH_TIME,
                        ahead_behind=AHEAD_BEHIND,
                        point_diff=POINT_DIFF
                    )
                    write_df(sh, tab, df)
                    log_result(tab, "OK", f"{len(df)} rows")
                except Exception as e:
                    print("❌", tab, "->", e)
                    log_result(tab, "FAIL", str(e))

        print("🎯 All done.")
    finally:
        flush_run_log(sh)

if __name__ == "__main__":
    main()
