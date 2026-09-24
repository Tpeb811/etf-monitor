# -*- coding: utf-8 -*-
"""
ETF 趋势监控 V3.9【基于V3.8新增综合研判】
====================================================
(此处省略你原有的长篇注释，保持原样即可)
"""
import os, re, io, sys, json, time, datetime, warnings, shutil
warnings.filterwarnings("ignore")
import requests
import numpy as np
import pandas as pd
import akshare as ak
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import LineChart, Reference

# ============================ 全局配置 ============================
MONITOR_ROOT = r"D:/zzxt/etf_monitor"
REPORT = r"D:/zzxt/ETF趋势监控报告.xlsx"
ETF_LIST = [
    "159593","510300","159338","510500","159901","159915","512100","159898","512880",
    "563300","588000","511380","159905","159928","510150","589720","510050","159698"
]
BACKFILL_START = "2026-01-01"
BENCHMARK = "sz399106"
TOP_N = 10
MAX_HOLDINGS = 100
MIN_COVER_PCT = 50.0
HOLD_TOPLINE = 100
W_ATR, W_SLOPE, W_RS = 0.40, 0.30, 0.30
FETCH_START = "2024-01-01"
MAX_SUSPEND_REUSE = 5
CACHE_GAP_THRESHOLD_DAY = 20
FULL_REFETCH_INTERVAL_DAY = 30
MID_GAP_THRESHOLD = 3
CHUNK_MONTH = 6
ENABLE_INTRADAY_SNAP = False

POS_LOOKBACK = 250
POS_LOW_LINE = 20.0
POS_HIGH_LINE = 80.0
POS_USE_FIX_DATE = True
POS_FIX_START_DATE = "2024-01-01"
NEWHL_LOOKBACK = 120
NEWHL_RECENT = 20
MOM_WINDOW = 20
BREADTH_LINE = 20.0
TB_MIN_COVERAGE = 0.7

if not os.path.exists("D:/"):
    _base = os.path.dirname(os.path.abspath(__file__))
    MONITOR_ROOT = os.path.join(_base, "etf_monitor")
    REPORT = os.path.join(_base, "ETF趋势监控报告.xlsx")

PX_CACHE = os.path.join(MONITOR_ROOT, "行情缓存")
HIST_CSV = os.path.join(MONITOR_ROOT, "趋势历史.csv")
HOLD_JSON = os.path.join(MONITOR_ROOT, "持仓档案.json")
LOG_FILE = os.path.join(MONITOR_ROOT, "运行日志.log")
META_JSON = os.path.join(MONITOR_ROOT, "缓存元信息.json")
LOCK_FILE = os.path.join(MONITOR_ROOT, "运行锁.lock")

_s = W_ATR + W_SLOPE + W_RS
W_ATR, W_SLOPE, W_RS = W_ATR/_s, W_SLOPE/_s, W_RS/_s
os.makedirs(PX_CACHE, exist_ok=True)

# ---------------------- 日志 ----------------------
def log(m, file_log=True):
    msg = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {m}"
    print(msg, flush=True)
    if file_log:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f: f.write(msg + "\n")
        except Exception: pass

# === 中间省略你原有的 1~7 部分的所有函数（get_holdings_em, fetch_by_segment, calc_etf_scores_range 等等） ===
# === 请把你原文件中的这些函数原封不动地粘贴在这里 ===
# === 包括：get_holdings_em, tx_symbol, _load_meta, _save_meta, _fetch_chunk, fetch_by_segment, update_price_cache, _mid_gap_check, get_benchmark, get_intraday_k_for_codes, stock_trend_series, load_holdings_archive, save_holdings_archive, holdings_signature, check_holdings_changes, holdings_on, load_history, save_history, calc_etf_scores_range, calc_etf_intraday_point, stock_position_metrics, calc_etf_top_bottom, comprehensive_verdict, write_report, main, _main_body ===
# ---------------------- 1. 持仓抓取（动态扩样：前十必取，覆盖≥50%或最多前100） ----------------------
def get_holdings_em(code):
    """返回 (季度, DataFrame[code,name,w], 说明)。债券型ETF返回(None,None,说明)"""
    year = datetime.date.today().year
    note = ""
    for y in (year, year-1):
        for retry in range(2):
            try:
                r = requests.get(
                    "https://fundf10.eastmoney.com/FundArchivesDatas.aspx",
                    params={"type":"jjcc","code":code,"topline":HOLD_TOPLINE,"year":y,"month":""},
                    headers={"Referer":f"https://fundf10.eastmoney.com/ccmx_{code}.html",
                             "User-Agent":"Mozilla/5.0"}, timeout=15)
                m = re.search(r'content:"(.*)",arryear', r.text, re.S)
                if not m: continue
                html = m.group(1).replace('\\"','"').replace("\\'","'").replace("\\/","/")
                for b in re.split(r"<h4[^>]*>", html)[1:]:
                    quarter = re.sub(r"<[^>]+>","",b.split("</h4>")[0]).replace("&nbsp;"," ").strip()
                    quarter = re.sub(r"\s+"," ",quarter)
                    if "<table" not in b: continue
                    tbl = pd.read_html(io.StringIO("<table"+b.split("<table",1)[1]))[0]
                    cc = [c for c in tbl.columns if "代码" in str(c)]
                    nc = [c for c in tbl.columns if "名称" in str(c)]
                    pc = [c for c in tbl.columns if "净值" in str(c)]
                    if not (cc and nc and pc): continue
                    out = pd.DataFrame({
                        "code": tbl[cc[0]].astype(str).str.extract(r"(\d+)")[0].str.zfill(6),
                        "name": tbl[nc[0]].astype(str),
                        "w": pd.to_numeric(tbl[pc[0]].astype(str).str.rstrip("%"), errors="coerce"),
                    }).dropna(subset=["code","w"])
                    out = out[out["code"].str.match(r"^(00|30|60|68)")].reset_index(drop=True)
                    if len(out) < 3:
                        return quarter, None, "无股票持仓(债券/货币型)"
                    # ---- 动态扩样：前十必取；覆盖<50%则向下取，直至≥50%或满100只 ----
                    cum = out["w"].cumsum()
                    n = TOP_N
                    if len(out) > TOP_N:
                        while n < min(len(out), MAX_HOLDINGS) and cum.iloc[n-1] < MIN_COVER_PCT:
                            n += 1
                    sel = out.head(n)
                    cover = float(sel["w"].sum())
                    if cover < MIN_COVER_PCT and len(out) <= TOP_N:
                        note = (f"覆盖{cover:.1f}%<{MIN_COVER_PCT:.0f}%，但该季度仅披露前十大"
                                f"(一/三季报规则)，按前十大计算")
                    elif cover < MIN_COVER_PCT:
                        note = f"已扩样至前{len(sel)}大，覆盖{cover:.1f}%仍<{MIN_COVER_PCT:.0f}%(披露上限)"
                    else:
                        note = f"前{len(sel)}大，覆盖{cover:.1f}%"
                    return quarter, sel, note
            except Exception as e:
                log(f"  {code} 持仓接口异常(y={y},retry={retry}):{e}")
                time.sleep(1)
    return None, None, "持仓获取失败"

# ---------------------- 2. 行情缓存（分段下载 + 尾端缺口 + 中段缺口 + 月度全量重拉） ----------------------
def tx_symbol(c6): return ("sh" if c6.startswith(("6","9")) else "sz") + c6

def _load_meta():
    if os.path.exists(META_JSON):
        try: return json.load(open(META_JSON, encoding="utf-8"))
        except Exception: return {}
    return {}

def _save_meta(m): json.dump(m, open(META_JSON,"w",encoding="utf-8"))

def _fetch_chunk(symbol, s_date, e_date, adjust="qfq"):
    for retry in range(3):
        try:
            df = ak.stock_zh_a_hist_tx(symbol=symbol, start_date=s_date.replace("-",""),
                                       end_date=e_date.replace("-",""), adjust=adjust)
            time.sleep(0.35)
            return df
        except Exception:
            time.sleep(1.2)
    return None

def fetch_by_segment(symbol, start_dt, end_dt, chunk_month=CHUNK_MONTH):
    pieces, cur = [], start_dt
    import calendar
    while cur <= end_dt:
        y, m = cur.year, cur.month + chunk_month
        while m > 12: m -= 12; y += 1
        try: nxt = datetime.date(y, m, cur.day)
        except ValueError: nxt = datetime.date(y, m, calendar.monthrange(y,m)[1])
        seg_end = min(nxt, end_dt)
        seg = _fetch_chunk(symbol, cur.isoformat(), seg_end.isoformat())
        if seg is not None and len(seg): pieces.append(seg)
        cur = seg_end + datetime.timedelta(days=1)
    if not pieces: return pd.DataFrame()
    full = pd.concat(pieces).rename(columns=str.lower)
    full["date"] = pd.to_datetime(full["date"])
    return full.drop_duplicates("date").sort_values("date").reset_index(drop=True)

def update_price_cache(code6, bench_dates, force_full_all):
    fp = os.path.join(PX_CACHE, code6 + ".csv")
    force_full = force_full_all
    old, warn = None, []
    today = datetime.date.today()
    if os.path.exists(fp):
        old = pd.read_csv(fp, parse_dates=["date"])
        if len(old):
            gap = (today - old["date"].max().date()).days
            if gap > CACHE_GAP_THRESHOLD_DAY:
                log(f"[CACHE_FIX] {code6} 缓存尾端缺口{gap}天，触发全量分段重拉")
                force_full = True
    fetch_start = datetime.date.fromisoformat(FETCH_START)
    if (not force_full) and old is not None and len(old):
        fetch_start = (old["date"].max() + datetime.timedelta(days=1)).date()
        if fetch_start >= today:
            df_total = old
            _mid_gap_check(code6, df_total, bench_dates, warn)
            # =========【新增告警1：无需下载直接返回缓存的分支】=========
            bench_max_dt = bench_dates.max()
            local_max_dt = df_total["date"].max()
            if local_max_dt < bench_max_dt:
                msg = f"[CACHE_WARN] {code6} 行情落后！本地最新:{local_max_dt.date()},基准最新:{bench_max_dt.date()}"
                log(msg)
                warn.append(msg)
            # ========================================================
            return df_total, warn
    df_new = fetch_by_segment(tx_symbol(code6), fetch_start, today)
    if len(df_new):
        net_gap = (today - df_new["date"].max().date()).days
        if net_gap > CACHE_GAP_THRESHOLD_DAY:
            log(f"[CACHE_WARN] {code6} 网络返回尾端缺口{net_gap}天")
    df_total = df_new if (force_full or old is None or not len(old)) else pd.concat([old, df_new])
    df_total = df_total.drop_duplicates("date").sort_values("date").reset_index(drop=True)
    if len(df_total): df_total.to_csv(fp, index=False)
    _mid_gap_check(code6, df_total, bench_dates, warn)
    # =========【新增告警2：完成下载、拼接、保存后的分支】=========
    if len(df_total) and len(bench_dates):
        bench_max_dt = bench_dates.max()
        local_max_dt = df_total["date"].max()
        if local_max_dt < bench_max_dt:
            msg = f"[CACHE_WARN] {code6} 行情落后！本地最新:{local_max_dt.date()},基准最新:{bench_max_dt.date()}"
            log(msg)
            warn.append(msg)
    # ========================================================
    return df_total, warn

def _mid_gap_check(code6, df, bench_dates, warn):
    """修正2：中段缺口校验——缓存日期与基准交易日比对"""
    if df is None or not len(df) or bench_dates is None or not len(bench_dates): return
    s = set(df["date"].dt.normalize())
    b = bench_dates.dt.normalize()
    overlap = b[(b >= df["date"].min()) & (b <= df["date"].max())]
    missing = [d for d in overlap if d not in s]
    if len(missing) > MID_GAP_THRESHOLD:
        msg = f"[CACHE_WARN] {code6} 中段缺口{len(missing)}个交易日(如{missing[0].date()}等)，请核查缓存"
        log(msg); warn.append(msg)

def get_benchmark():
    fp = os.path.join(PX_CACHE, f"BENCH_{BENCHMARK[2:]}.csv")

    for i in range(3):
        try:
            d = ak.stock_zh_index_daily_tx(symbol=BENCHMARK)
            d["date"] = pd.to_datetime(d["date"])
            d = d[d["date"] >= FETCH_START].sort_values("date").reset_index(drop=True)
            d.to_csv(fp, index=False)
            return d
        except Exception: time.sleep(1.5)
    if os.path.exists(fp): return pd.read_csv(fp, parse_dates=["date"])
    raise RuntimeError(f"基准行情获取失败({BENCHMARK})")

# ---------------------- 盘中快照（默认关闭，保留V3.6逻辑） ----------------------
def get_intraday_k_for_codes(code_list, base_date):
    spot_df = None
    for retry in range(2):
        try: spot_df = ak.stock_zh_a_spot_em(); break
        except Exception as e: log(f"[WARN] 实时盘口失败 retry={retry}:{e}"); time.sleep(1.0)
    if spot_df is None: return dict()
    spot_df["代码"] = spot_df["代码"].astype(str).str.zfill(6)
    spot_map = {r["代码"]: {"open":r["今开"],"high":r["最高"],"low":r["最低"],"close":r["最新价"]}
                for _,r in spot_df.iterrows()}
    res, base_dt = {}, pd.Timestamp(base_date)
    for c6 in code_list:
        fp = os.path.join(PX_CACHE, c6 + ".csv")
        if not os.path.exists(fp): continue
        df = pd.read_csv(fp, parse_dates=["date"])
        df = df[df["date"].dt.date != base_date].copy()
        sp = spot_map.get(c6)
        if sp is None or pd.isna(sp["close"]) or sp["close"] <= 0:
            res[c6] = df; continue
        df = pd.concat([df, pd.DataFrame([{"date":base_dt,"open":sp["open"],"high":sp["high"],
                                           "low":sp["low"],"close":sp["close"]}])], ignore_index=True)
        res[c6] = df.sort_values("date").reset_index(drop=True)
    return res

# ---------------------- 3. 个股趋势分 ----------------------
def stock_trend_series(d, bench_r60):
    d = d.set_index("date").sort_index()
    c,h,l = d["close"], d["high"], d["low"]
    ma20,ma60,ma120 = (c.rolling(n).mean() for n in (20,60,120))
    tr = pd.concat([h-l,(h-c.shift()).abs(),(l-c.shift()).abs()],axis=1).max(axis=1)
    atr = tr.rolling(14).mean()
    atr_s = (((50+(c-ma20)/atr*10)+(50+(c-ma60)/atr*10)+(50+(c-ma120)/atr*10))/3).clip(0,100)
    slope_s = (50+(ma60-ma60.shift(20))/ma60*5000).clip(0,100)
    excess = (c.pct_change(60) - bench_r60.reindex(d.index)).dropna()
    rs_s = excess.rolling(252, min_periods=30).apply(lambda x:(x<x.iloc[-1]).mean()*100)
    return (atr_s*W_ATR + slope_s*W_SLOPE + rs_s*W_RS).round(2)

# ---------------------- 4. 持仓档案 ----------------------
def load_holdings_archive():
    if os.path.exists(HOLD_JSON): return json.load(open(HOLD_JSON, encoding="utf-8"))
    return {}

def save_holdings_archive(a): json.dump(a, open(HOLD_JSON,"w",encoding="utf-8"), ensure_ascii=False, indent=1)

def holdings_signature(items): return [(i["code"], round(float(i["w"]),2)) for i in items]

def check_holdings_changes(archive, today):
    changes, notes = [], {}
    for code in ETF_LIST:
        quarter, h, note = get_holdings_em(code)
        if h is None:
            log(f"  {code} {note}，跳过"); continue
        notes[code] = note
        items = h.to_dict("records")
        entry = archive.get(code)
        if entry is None:
            archive[code] = {"quarter": quarter, "vintages":[{"from":BACKFILL_START,"items":items}]}
            log(f"  {code} 首次建档（{quarter}；{note}）")
        else:
            entry["quarter"] = quarter
            cur = entry["vintages"][-1]["items"]
            if holdings_signature(cur) != holdings_signature(items):
                entry["vintages"].append({"from": today, "items": items})
                changes.append({"日期":today,"ETF":code,"季度":quarter,
                    "变更前":"、".join(i["name"] for i in cur),
                    "变更后":"、".join(i["name"] for i in items),
                    "说明":"变更生效日≈检测日，实际以季报截止日为准；变更日之后按新权重计算"})
                log(f"  {code} 持仓变更！已登记（{quarter}；{note}）")
            else:
                log(f"  {code} 持仓无变化（{quarter}；{note}）")
        time.sleep(0.3)
    return archive, changes, notes

def holdings_on(archive, code, date_str):
    vs = archive[code]["vintages"]; cur = vs[0]["items"]
    for v in vs:
        if v["from"] <= date_str: cur = v["items"]
    return cur

# ---------------------- 5. 历史趋势 ----------------------
def load_history():
    if os.path.exists(HIST_CSV): return pd.read_csv(HIST_CSV, dtype={"ETF":str})
    return pd.DataFrame(columns=["日期","ETF","趋势分","覆盖权重%"])

def save_history(df):
    df = df.drop_duplicates(["日期","ETF"], keep="last").sort_values(["日期","ETF"])
    df.to_csv(HIST_CSV, index=False, encoding="utf-8-sig")
    return df

# ---------------------- 6. ETF分数计算 ----------------------
def calc_etf_scores_range(code, bench_r60, trade_dates, archive):
    wsets = {}
    for v in archive[code]["vintages"]:
        key = json.dumps(holdings_signature(v["items"]))
        if key not in wsets: wsets[key] = {i["code"]: float(i["w"]) for i in v["items"]}
    cache = {}
    def series_for(wmap):
        key = json.dumps(sorted(wmap.items()))
        if key in cache: return cache[key]
        cols = {}
        for c6 in wmap:
            fp = os.path.join(PX_CACHE, c6 + ".csv")
            if not os.path.exists(fp): continue
            d = pd.read_csv(fp, parse_dates=["date"])
            if len(d) < 130: continue
            cols[c6] = stock_trend_series(d, bench_r60)
        cache[key] = pd.DataFrame(cols).reindex(trade_dates) if cols else None
        return cache[key]
    out = pd.Series(index=trade_dates, dtype=float)
    cover_pct_series = pd.Series(index=trade_dates, dtype=float)
    day_cover_info, day_stock_score = {}, {}
    for dt in trade_dates:
        ds = dt.strftime("%Y-%m-%d")
        items = holdings_on(archive, code, ds)
        wmap = {i["code"]: float(i["w"]) for i in items}
        name_map = {i["code"]: i["name"] for i in items}
        sdf = series_for(wmap)
        if sdf is None: continue
        valid, missing, reuse, sdict = [], [], {}, {}
        for c in sdf.columns:
            val = sdf.loc[dt, c]
            if pd.notna(val):
                valid.append(c); sdict[c] = {"name":name_map[c],"score":val}; continue
            lv = sdf[c][:dt].last_valid_index()
            if lv is None: missing.append(c); continue
            delta = (dt - lv).days
            if delta <= MAX_SUSPEND_REUSE:
                reuse[c] = sdf[c].loc[lv]; valid.append(c)
                sdict[c] = {"name":name_map[c],"score":reuse[c]}
            else:
                missing.append(c)
        total_w = sum(wmap.values()); valid_w = sum(wmap[c] for c in valid)
        cover = valid_w/total_w if total_w > 1e-6 else 0.0
        cover_pct_series[dt] = round(total_w, 2)
        day_cover_info[dt] = {"cover_pct":cover,"missing":missing,"suspend":list(reuse)}
        day_stock_score[dt] = sdict
        if valid_w < 1e-6: continue
        wv = pd.Series({c:wmap[c] for c in valid})
        sv = pd.Series({c:(reuse[c] if c in reuse else sdf.loc[dt,c]) for c in valid})
        out[dt] = round(sv.mul(wv).sum()/valid_w, 1)
    return out, cover_pct_series, day_cover_info, day_stock_score

def calc_etf_intraday_point(code, archive, intraday_dt, bench_r60):
    ds = intraday_dt.strftime("%Y-%m-%d")
    items = holdings_on(archive, code, ds)
    wmap = {i["code"]: float(i["w"]) for i in items}
    name_map = {i["code"]: i["name"] for i in items}
    try: kmap = get_intraday_k_for_codes(list(wmap), intraday_dt.date())
    except Exception as e:
        log(f"[INTRA_ERROR]{e}"); return (np.nan,0.0,["盘中行情异常"],[],{})
    if not kmap: return (np.nan,0.0,["实时盘口为空"],[],{})
    valid, missing, reuse, smap, sdict = [], [], {}, {}, {}
    for c6 in wmap:
        if c6 not in kmap or len(kmap[c6]) < 130: missing.append(c6); continue
        s = stock_trend_series(kmap[c6], bench_r60)
        val = s.loc[intraday_dt] if intraday_dt in s.index else np.nan
        if pd.notna(val):
            valid.append(c6); smap[c6]=val; sdict[c6]={"name":name_map[c6],"score":val}; continue
        lv = s[:intraday_dt].last_valid_index()
        if lv is not None and (intraday_dt-lv).days <= MAX_SUSPEND_REUSE:
            valid.append(c6); smap[c6]=s.loc[lv]; reuse[c6]=s.loc[lv]
            sdict[c6]={"name":name_map[c6],"score":s.loc[lv]}
        else: missing.append(c6)
    total_w = sum(wmap.values()); valid_w = sum(wmap[c] for c in valid)
    if valid_w < 1e-6: return (np.nan, 0.0, missing, list(reuse), {})
    wv = pd.Series({c:wmap[c] for c in valid})
    return (round(pd.Series(smap).mul(wv).sum()/valid_w,1), valid_w/total_w, missing, list(reuse), sdict)

# ---------------------- 6.5 V3.8 顶底观察 ----------------------
def stock_position_metrics(c6, snap_date):
    """从缓存日K计算个股顶底指标；直接读取本地缓存，不触发网络下载。
    数据截止 snap_date；停牌股最新K线距今≤MAX_SUSPEND_REUSE天则沿用，否则返回None。"""
    fp = os.path.join(PX_CACHE, c6 + ".csv")
    if not os.path.exists(fp):
        return None
    # 读取本地已经下载好的缓存K线，无网络请求
    d = pd.read_csv(fp, parse_dates=["date"]).sort_values("date")
    d = d[d["date"] <= snap_date].copy()

    if len(d) < 130:
        return None
    if (pd.Timestamp(snap_date) - d["date"].iloc[-1]).days > MAX_SUSPEND_REUSE:
        return None

    c, h, l = d["close"], d["high"], d["low"]

    # 根据开关选择区间模式，全部内存运算，不读网络
    if POS_USE_FIX_DATE:
        # 固定起始日期模式
        anchor_dt = pd.Timestamp(POS_FIX_START_DATE)
        win = d.loc[d["date"] >= anchor_dt]
        # 容错：新股数据晚于锚定日期，取全部可用数据
        if len(win) < 20:
            win = d
    else:
        # 原版：滚动最近N个交易日窗口
        win = d.tail(POS_LOOKBACK)

    lo, hi = win["low"].min(), win["high"].max()
    if not (hi > lo > 0):
        return None
    pos = (c.iloc[-1] - lo) / (hi - lo) * 100

    if len(c) <= 2 * MOM_WINDOW:
        return None
    mom_now = (c.iloc[-1] / c.iloc[-1 - MOM_WINDOW] - 1) * 100
    mom_prev = (c.iloc[-1 - MOM_WINDOW] / c.iloc[-1 - 2 * MOM_WINDOW] - 1) * 100
    if pd.isna(mom_now) or pd.isna(mom_prev):
        return None

    ref = d.tail(NEWHL_LOOKBACK + NEWHL_RECENT)
    recent = d.tail(NEWHL_RECENT)
    new_high = bool(len(ref) and recent["high"].max() >= ref["high"].max() - 1e-9)
    new_low = bool(len(ref) and recent["low"].min() <= ref["low"].min() + 1e-9)

    return {"pos": pos, "mom_now": mom_now, "mom_prev": mom_prev,
            "new_high": new_high, "new_low": new_low}


def calc_etf_top_bottom(code, archive, snap_date):
    """ETF顶底观察：权重股指标按持仓权重加权汇总，输出位置分/动量/结构占比/信号"""
    items = holdings_on(archive, code, snap_date.strftime("%Y-%m-%d"))
    total_disclosed_w = sum(float(it["w"]) for it in items)   # 全部披露持仓权重之和
    rows = []
    for it in items:
        m = stock_position_metrics(it["code"], snap_date)
        if m is None: continue
        rows.append({"code": it["code"], "name": it["name"], "w": float(it["w"]), **m})
    if not rows: return None
    df = pd.DataFrame(rows)
    W, tw = df["w"], df["w"].sum()
    if tw < 1e-6: return None
    coverage = tw / total_disclosed_w if total_disclosed_w > 1e-6 else 0.0   # 数据覆盖率
    pos_w = (df["pos"] * W).sum() / tw
    mom_now_w = (df["mom_now"] * W).sum() / tw
    mom_prev_w = (df["mom_prev"] * W).sum() / tw
    nh = df.loc[df["new_high"], "w"].sum() / tw * 100
    nl = df.loc[df["new_low"], "w"].sum() / tw * 100
    decel = mom_now_w - mom_prev_w
    if pos_w <= POS_LOW_LINE and decel > 0 and nl <= BREADTH_LINE:
        sig = "底部衰竭观察"
    elif pos_w >= POS_HIGH_LINE and decel < 0 and nh <= BREADTH_LINE:
        sig = "顶部背离预警"
    elif pos_w <= POS_LOW_LINE:
        sig = "低位区"
    elif pos_w >= POS_HIGH_LINE:
        sig = "高位区"
    else:
        sig = "—"
    # 数据覆盖率不足时，信号追加"（样本不足）"后缀，降低可信度
    if coverage < TB_MIN_COVERAGE and sig != "—":
        sig = sig + "（样本不足）"
    low3, high3 = df.nsmallest(3, "pos"), df.nlargest(3, "pos")
    return {"位置分": round(pos_w, 1), "近20日动量%": round(mom_now_w, 2),
            "前20日动量%": round(mom_prev_w, 2), "动量变化": round(decel, 2),
            "创新低%": round(nl, 1), "创新高%": round(nh, 1), "顶底信号": sig,
            "数据覆盖%": round(coverage * 100, 1),
            "低位个股TOP3": "、".join(f"{r['name']}({r['pos']:.0f}%)" for _, r in low3.iterrows()),
            "高位个股TOP3": "、".join(f"{r['name']}({r['pos']:.0f}%)" for _, r in high3.iterrows())}


# ---------------------- 6.6 V3.9 综合研判（趋势分 × 顶底信号/位置/动量 交叉验证） ----------------------
def comprehensive_verdict(trend, sig, pos=np.nan, mom_now=np.nan, decel=np.nan):
    """输出 实战标签[触发条件](确认纪律备注)。
    注意：经验规则，未经历史回测，仅作观察参考，不构成投资建议。
    确认纪律：买入类标签出现后，次日趋势分不回落（或上穿60）再动手；风控类标签宜早不宜迟。"""
    if pd.isna(trend): return "—"
    sig = str(sig or "—")
    weak = "（样本不足）" in sig
    base_sig = sig.replace("（样本不足）", "")
    tr = f"趋势{trend:.1f}"
    # 1) 强势回调·可建仓加仓（新增）：趋势强 + 位置中位(20~70) + 近20日回调 + 回调放缓
    if (trend >= 60 and pd.notna(pos) and 20 < pos <= 70
            and pd.notna(mom_now) and mom_now < 0 and pd.notna(decel) and decel > 0):
        v = f"🟢 强势回调·可建仓加仓[{tr}+位置{pos:.0f}中位+近20日回调{mom_now:+.1f}%但放缓](次日趋势分不回落再动手)"
    # 2) 底部企稳·可左侧建仓（新增）：底部衰竭 + 近20日已不再下跌
    elif "底部衰竭观察" in base_sig and pd.notna(mom_now) and mom_now >= 0:
        v = f"🟢 底部企稳·可左侧建仓[{tr}+底部衰竭+近20日止跌{mom_now:+.1f}%](次日趋势分不回落再动手)"
    # 3) 底部衰竭但仍在跌：按趋势分层
    elif "底部衰竭观察" in base_sig:
        if trend >= 60:   v = f"🟢 深跌强修复·关注[{tr}+底部衰竭·位置≤20](等回踩不破前低再介入)"
        elif trend >= 40: v = f"🟡 拐点观察·待确认[{tr}+底部衰竭·仍在跌](等趋势分上穿60再动手)"
        else:             v = f"🔴 弱势假底·勿接刀[{tr}+底部衰竭·仍在跌](反弹视为减仓机会)"
    # 4) 顶部背离：按趋势分层
    elif "顶部背离预警" in base_sig:
        if trend >= 60:   v = f"⚠️ 高位减仓预警[{tr}+顶部背离](分批减，不一次清仓)"
        elif trend >= 40: v = f"🟡 转弱观察[{tr}+顶部背离](趋势分跌破40确认转弱)"
        else:             v = f"🔴 弱反弹见顶·回避[{tr}+顶部背离](反弹不参与)"
    else:
        return "—"
    return v + "·样本不足" if weak else v

# ---------------------- 7. Excel报告 ----------------------
def write_report(hist, snapshot_df, is_intraday_snap, snap_date_str, changes, checks, tb_df):
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    scale = ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                           mid_type="num", mid_value=50, mid_color="FFEB84",
                           end_type="num", end_value=100, end_color="63BE7B")
    ws = wb.active; ws.title = "当日快照"; ws.sheet_view.showGridLines = False
    title = (f"当日快照【盘中未收盘，仅供观察，不写历史】（{snap_date_str}）" if is_intraday_snap
             else f"当日快照（{snap_date_str}，收盘数据；当天多次运行只保留最后一次）")
    ws.cell(row=2, column=2, value=title).font = Font(size=13, bold=True,
                                                      color="FF0000" if is_intraday_snap else "000000")
    note_text = """【指标简要说明】
    趋势分信号：强势(≥60)/震荡(40~60)/弱势(<40)，代表大方向；
    位置分：成分股加权，价格在指定起点以来波段区间的百分位0‑100；
    动量变化：加速度(近20日动量-前20日动量)；>0力度改善，<0力度衰减，不等于涨跌方向；
    顶底信号：底部衰竭观察/顶部背离预警仅为反转前置观察信号，不能单独作为买卖依据；
    综合研判：趋势分×顶底信号/位置/动量交叉验证的经验标签（强势回调·可建仓加仓/底部企稳·可左侧建仓/深跌强修复/拐点观察/弱势假底/高位减仓/转弱观察/弱反弹见顶），[]内为触发条件，()内为确认纪律，未经回测，仅供参考；
    确认纪律：买入类标签出现后，次日趋势分不回落（或上穿60）再动手；风控类标签宜早不宜迟；
    创新低%/创新高%：满足条件成分股权重占比；带（样本不足）代表数据可信度下降。"""

    cols = ["ETF", "名称", "趋势分", "信号", "位置分", "动量变化", "顶底信号", "综合研判", "创新低%", "创新高%", "覆盖权重%",
            "数据覆盖%", "最高分个股", "最低分个股"]
    for j, t in enumerate(cols):
        c = ws.cell(row=4, column=2 + j, value=t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill

    # ========== 列宽代码放在这里，表头循环之后 ==========
    for j, wd in enumerate([10, 16, 8, 8, 8, 10, 13, 52, 9, 9, 12, 12, 24, 24]):
        ws.column_dimensions[chr(66 + j)].width = wd

    for i, (_, r) in enumerate(snapshot_df.sort_values("趋势分", ascending=False).iterrows()):
        for j, k in enumerate(cols):
            ws.cell(5 + i, 2 + j, r[k])

    # 修复条件格式这一行
    if len(snapshot_df):
        ws.conditional_formatting.add(f"D5:D{4 + len(snapshot_df)}", scale)
    for j,wd in enumerate([10,16,8,8,8,10,13,52,9,9,12,12,24,24]): ws.column_dimensions[chr(66+j)].width = wd

    # ===== 指标说明：移到表格下方 + 自动换行分行显示 =====
    last_data_row = 4 + len(snapshot_df)        # 数据占用的最后一行
    note_row = last_data_row + 2                # 空一行后，写说明
    ncell = ws.cell(note_row, 2, value=note_text)
    ncell.font = Font(size=9, color="444444")
    ncell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    # 合并整行宽度（表头从第2列到第2+13-1列），让说明横跨表格
    ws.merge_cells(start_row=note_row, start_column=2,
                   end_row=note_row, end_column=1 + len(cols))
    ws.row_dimensions[note_row].height = 90     # 合并+换行时Excel不会自动撑高，须手给行高

    # ---- V3.8 顶底观察 sheet ----
    ts = wb.create_sheet("顶底观察"); ts.sheet_view.showGridLines = False
    ts.cell(2,2,"顶底观察（位置分=权重股在近250日高低区间中的加权位置；"
                "底部衰竭观察=位置≤20+动量回升+创新低占比≤20%；顶部背离预警=位置≥80+涨速放缓+创新高占比≤20%；不构成投资建议）").font = Font(size=12,bold=True)
    tcols = ["ETF","名称","趋势分","位置分","顶底信号","近20日动量%","前20日动量%","动量变化","创新低%","创新高%","数据覆盖%","低位个股TOP3","高位个股TOP3"]
    for j,t in enumerate(tcols):
        c = ts.cell(4,2+j,t); c.font = Font(bold=True,color="FFFFFF"); c.fill = fill
    if tb_df is not None and len(tb_df):
        tb_sorted = tb_df.sort_values("位置分")
        for i,(_,r) in enumerate(tb_sorted.iterrows()):
            for j,k in enumerate(tcols): ts.cell(5+i,2+j,r[k])
            sig = r["顶底信号"]
            if "底部衰竭观察" in sig or "顶部背离预警" in sig:
                for j in range(len(tcols)):
                    ts.cell(5+i,2+j).font = Font(bold=True,color="1F7A1F" if "底部衰竭观察" in sig else "C00000")
        ts.conditional_formatting.add(f"E5:E{4+len(tb_sorted)}", scale)
    for j,wd in enumerate([10,16,8,8,16,11,11,9,9,9,10,34,34]): ts.column_dimensions[chr(66+j)].width = wd

    hs = wb.create_sheet("历史趋势"); hs.sheet_view.showGridLines = False
    hs.cell(2,2,"逐日趋势分（每行一天；首次自2026-01-01补齐，之后增量；当天重跑覆盖）").font = Font(size=13,bold=True)
    piv = hist.pivot_table(index="日期", columns="ETF", values="趋势分").sort_index()
    etfs = list(piv.columns)
    hs.cell(4,2,"日期").font = Font(bold=True,color="FFFFFF"); hs.cell(4,2).fill = fill
    for j,e in enumerate(etfs):
        c = hs.cell(4,3+j,e); c.font = Font(bold=True,color="FFFFFF"); c.fill = fill
    for i,(dt,row) in enumerate(piv.iterrows()):
        hs.cell(5+i,2,dt)
        for j,e in enumerate(etfs):
            v = row[e]
            if pd.notna(v): hs.cell(5+i,3+j,float(v))
    hs.column_dimensions["B"].width = 12
    for j in range(len(etfs)): hs.column_dimensions[chr(67+j)].width = 10
    if len(etfs) and len(piv):
        hs.conditional_formatting.add(f"C5:{chr(66+len(etfs))}{4+len(piv)}", scale)
        ch = LineChart(); ch.title = "ETF趋势分走势"; ch.height = 9; ch.width = 24
        ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 100
        ch.add_data(Reference(hs,min_col=3,min_row=4,max_col=2+len(etfs),max_row=4+len(piv)), titles_from_data=True)
        ch.set_categories(Reference(hs,min_col=2,min_row=5,max_row=4+len(piv)))
        hs.add_chart(ch, f"B{6+len(piv)}")

    cs = wb.create_sheet("持仓变更记录"); cs.sheet_view.showGridLines = False
    cs.cell(2,2,"持仓变更记录（变更日之后按新权重计算；生效日≈检测日，实际以季报截止日为准）").font = Font(size=13,bold=True)
    ccols = ["日期","ETF","季度","变更前","变更后","说明"]
    for j,t in enumerate(ccols):
        c = cs.cell(4,2+j,t); c.font = Font(bold=True,color="FFFFFF"); c.fill = fill
    if changes:
        for i,chg in enumerate(changes):
            for j,k in enumerate(ccols): cs.cell(5+i,2+j,chg.get(k,""))
    else:
        cs.cell(5,2,"截至本次运行未检测到持仓变更（每次运行自动核对）")
    for j,wd in enumerate([12,10,26,55,55,42]): cs.column_dimensions[chr(66+j)].width = wd

    vs = wb.create_sheet("校验信息"); vs.sheet_view.showGridLines = False
    vs.cell(2,2,"数据校验｜停牌复用｜缺口与低覆盖率清单").font = Font(size=13,bold=True)
    for i,line in enumerate(checks): vs.cell(4+i,2,line)
    vs.column_dimensions["B"].width = 130
    wb.save(REPORT)
    log(f"Excel报告已写入 {REPORT}")
# (提示：为了保证代码完整，下方我保留了 main 和 _main_body 的入口，你把上面的函数贴在这两者之间即可)

def main():
    t0 = time.time()
    today_d = datetime.date.today(); today = today_d.strftime("%Y-%m-%d")
    checks = []
    if os.path.exists(LOCK_FILE):
        try:
            age_h = (time.time() - os.path.getmtime(LOCK_FILE)) / 3600
            if age_h < 3:
                log(f"检测到另一实例正在运行（锁已存在{age_h:.1f}小时），本次退出。若确认无其他实例在跑，请删除 {LOCK_FILE}")
                return
            os.remove(LOCK_FILE)
        except Exception: pass
    try:
        with open(LOCK_FILE, "w", encoding="utf-8") as f: f.write(str(os.getpid()))
    except Exception: pass
    try:
        _main_body(t0, today_d, today, checks)
    finally:
        try: os.remove(LOCK_FILE)
        except Exception: pass

def _main_body(t0, today_d, today, checks):
    # === 把你原有的 _main_body 函数内容完整贴在这里 ===
    log("====== ETF趋势监控 V3.9 启动 ======")
    log(f"报告路径：{REPORT}")
    log(f"基准：{BENCHMARK}（修改配置区BENCHMARK即可更换）")
    log("【1/4】核对持仓（前十必取；覆盖<50%自动扩样至最多前100）...")
    archive = load_holdings_archive()
    archive, changes, hold_notes = check_holdings_changes(archive, today)
    save_holdings_archive(archive)

    log("【2/4】更新行情缓存 ...")
    bench = get_benchmark()
    bench_r60 = bench.set_index("date")["close"].pct_change(60)
    bench_dates = bench["date"]
    trade_dates = [d for d in bench["date"] if d >= pd.Timestamp(BACKFILL_START)]
    last_closed = bench["date"].max()
    log(f"基准最新已收盘交易日：{last_closed.date()}")
    # 修正1：月度全量重拉（防前复权漂移）
    meta = _load_meta()
    last_full = meta.get("last_full_refetch", "2000-01-01")
    days_since_full = (today_d - datetime.date.fromisoformat(last_full)).days
    force_full_all = days_since_full >= FULL_REFETCH_INTERVAL_DAY
    if force_full_all:
        log(f"距上次全量重拉已{days_since_full}天（≥{FULL_REFETCH_INTERVAL_DAY}），本次全部全量重拉以防前复权漂移")
    need = set()
    for code in [c for c in ETF_LIST if c in archive]:
        for v in archive[code]["vintages"]:
            for it in v["items"]: need.add(it["code"])
    log(f"  涉及个股 {len(need)} 只")
    bad, mid_warn = [], []
    for i, c6 in enumerate(sorted(need)):
        d, w = update_price_cache(c6, bench_dates, force_full_all)
        mid_warn += w
        if d is None or len(d) < 130: bad.append(c6)
        if (i+1) % 30 == 0: log(f"  进度 {i+1}/{len(need)}")
        time.sleep(0.25)
    if force_full_all and not bad:
        meta["last_full_refetch"] = today; _save_meta(meta)
    checks.append(f"行情校验：共{len(need)}只；历史K不足130天的{len(bad)}只 {bad if bad else '无'}；基准最新收盘 {last_closed.date()}；距上次全量重拉{days_since_full}天")
    checks += mid_warn if mid_warn else ["中段缺口校验：全部通过"]

    log("【3/4】计算趋势分，增量写入历史 ...")
    hist = load_history()
    new_rows, all_cover, all_stock = [], {}, {}
    for code in [c for c in ETF_LIST if c in archive]:
        s, cover_s, cov, sdict = calc_etf_scores_range(code, bench_r60, trade_dates, archive)
        all_cover[code], all_stock[code] = cov, sdict
        if s.dropna().empty:
            log(f"  {code} 无有效分数，跳过"); continue
        tmp = pd.DataFrame({"日期":[d.strftime("%Y-%m-%d") for d in s.index],
                            "ETF":code, "趋势分":s.values,
                            "覆盖权重%":[cover_s.get(d, np.nan) for d in s.index]})
        tmp = tmp.dropna(subset=["趋势分"])
        new_rows.append(tmp)
        log(f"  {code} 计算完成（{len(tmp)}天）")
    if new_rows: hist = pd.concat([hist] + new_rows)
    hist = save_history(hist)
    checks.append(f"历史记录：共{len(hist)}条；当天重复运行按[日期+ETF]覆盖，只保留最后一次")

    # 低覆盖率明细（修正5）
    low_cov = []
    for code, cov in all_cover.items():
        for dt, info in cov.items():
            if info["cover_pct"] < 0.8 and (info["missing"] or info["suspend"]):
                low_cov.append(f"低覆盖：{code} {dt.date()} 覆盖{info['cover_pct']:.0%} 缺失{info['missing']} 停牌复用{info['suspend']}")
    checks.append(f"覆盖率<80%的ETF-日期共{len(low_cov)}条" + ("；" + "；".join(low_cov[:10]) + (" …" if len(low_cov)>10 else "") if low_cov else ""))

    # 快照
    is_today_closed = (last_closed.date() == today_d)
    snap_rows, names_map, cover_map, high_map, low_map = [], {}, {}, {}, {}
    is_intraday = False
    if is_today_closed or not ENABLE_INTRADAY_SNAP:
        snap_date = last_closed if not is_today_closed else pd.Timestamp(today_d)
        snap_str = snap_date.strftime("%Y-%m-%d")
        sub = hist[hist["日期"] == snap_str]
        snap_rows = [{"ETF":r["ETF"],"趋势分":r["趋势分"]} for _,r in sub.iterrows()]
        for code in [c for c in ETF_LIST if c in archive]:
            items = archive[code]["vintages"][-1]["items"]
            names_map[code] = round(sum(i["w"] for i in items),1)
            di = all_cover.get(code,{}).get(snap_date,{})
            cover_map[code] = f"{di.get('cover_pct',0.0):.1%}"
            w_by_code = {i["code"]: float(i["w"]) for i in items}
            sl = [{"name":v["name"],"score":v["score"],"w":w_by_code.get(c6,0.0)} for c6,v in all_stock.get(code,{}).get(snap_date,{}).items()]
            if sl:
                dfs = pd.DataFrame(sl)
                dfs["wscore"] = dfs["score"]*dfs["w"]   # 加权趋势分 = 个股趋势分 × 该股权重
                hi, lo = dfs.loc[dfs["wscore"].idxmax()], dfs.loc[dfs["wscore"].idxmin()]
                high_map[code] = f"{hi['name']}(加权{round(hi['wscore'],1)}/原{hi['score']})"; low_map[code] = f"{lo['name']}(加权{round(lo['wscore'],1)}/原{lo['score']})"
            else: high_map[code] = low_map[code] = "-"
    else:
        is_intraday = True
        snap_str = today
        for code in [c for c in ETF_LIST if c in archive]:
            score, cov, miss, susp, sdict = calc_etf_intraday_point(code, archive, pd.Timestamp(today_d), bench_r60)
            snap_rows.append({"ETF":code,"趋势分":score})
            items = archive[code]["vintages"][-1]["items"]
            names_map[code] = round(sum(i["w"] for i in items),1)
            cover_map[code] = f"{cov:.1%}"
            w_by_code = {i["code"]: float(i["w"]) for i in items}
            sl = [{"name":v["name"],"score":v["score"],"w":w_by_code.get(c6,0.0)} for c6,v in sdict.items()]
            if sl:
                dfs = pd.DataFrame(sl)
                dfs["wscore"] = dfs["score"]*dfs["w"]   # 加权趋势分 = 个股趋势分 × 该股权重
                hi, lo = dfs.loc[dfs["wscore"].idxmax()], dfs.loc[dfs["wscore"].idxmin()]
                high_map[code] = f"{hi['name']}(加权{round(hi['wscore'],1)}/原{hi['score']})"; low_map[code] = f"{lo['name']}(加权{round(lo['wscore'],1)}/原{lo['score']})"
            else: high_map[code] = low_map[code] = "-"
    snap = pd.DataFrame(snap_rows).dropna(subset=["趋势分"])
    snap["信号"] = snap["趋势分"].apply(lambda x:"强势" if x>=60 else ("震荡" if x>=40 else "弱势"))
    snap["覆盖权重%"] = snap["ETF"].map(names_map)
    snap["数据覆盖%"] = snap["ETF"].map(cover_map)
    snap["最高分个股"] = snap["ETF"].map(high_map)
    snap["最低分个股"] = snap["ETF"].map(low_map)
    try:
        # 优先用持仓建档时已取到的ETF名称（季度串形如"红利ETF工银 2026年2季度股票投资明细…"），
        # 只对缺失的代码才补调一次全市场接口，避免无谓请求被断连
        def _name_from_quarter(q):
            m2 = re.match(r"^(.*?)\s*20\d{2}年", q or "")
            return m2.group(1).strip() if m2 else ""
        local_names = {c: _name_from_quarter(archive[c].get("quarter","")) for c in archive}
        missing_nm = [c for c in snap["ETF"] if not local_names.get(c)]
        if missing_nm:
            nm = ak.fund_etf_spot_em()[["代码","名称"]]
            nm_map2 = dict(zip(nm["代码"].astype(str), nm["名称"]))
            for c in missing_nm: local_names[c] = nm_map2.get(c, c)
        snap["名称"] = snap["ETF"].map(lambda c: local_names.get(c) or c)
    except Exception as e:
        log(f"ETF名称获取异常（已用持仓档案名称/代码兜底）:{e}"); snap["名称"] = snap["ETF"]
    snap["日期"] = snap_str

    # ---- V3.8 顶底观察：基于快照日的权重股日K（数据全部来自本地缓存，无新增接口） ----
    log("【3.5/4】计算顶底观察（位置分/动量衰竭/创新高新低结构）...")
    if POS_USE_FIX_DATE:
        log(f"    位置分模式：固定起始日期，锚点={POS_FIX_START_DATE}")
    else:
        log(f"    位置分模式：滚动窗口，回看{POS_LOOKBACK}交易日")
    tb_date = snap_date if not is_intraday else pd.Timestamp(today_d)
    tb_rows = []
    for code in [c for c in ETF_LIST if c in archive]:
        tb = calc_etf_top_bottom(code, archive, tb_date)
        if tb is None:
            log(f"  {code} 顶底观察无有效数据，跳过"); continue
        tb_rows.append({"ETF": code, **tb})
    tb_df = pd.DataFrame(tb_rows)
    if len(tb_df):
        nm_map = dict(zip(snap["ETF"], snap["名称"])) if "名称" in snap else {}
        tb_df["名称"] = tb_df["ETF"].map(nm_map).fillna(tb_df["ETF"])
        tr_map = dict(zip(snap["ETF"], snap["趋势分"]))
        tb_df["趋势分"] = tb_df["ETF"].map(tr_map)
        tb_df = tb_df.dropna(subset=["位置分"])
    for k in ["位置分", "顶底信号", "动量变化", "近20日动量%", "创新低%", "创新高%"]:
        snap[k] = snap["ETF"].map(dict(zip(tb_df["ETF"], tb_df[k]))) if len(tb_df) else np.nan
    snap["顶底信号"] = snap["顶底信号"].fillna("—")
    # ---- V3.9 综合研判列（含触发条件与确认纪律备注） ----
    snap["综合研判"] = snap.apply(lambda r: comprehensive_verdict(
        r["趋势分"], r["顶底信号"], r["位置分"], r["近20日动量%"], r["动量变化"]), axis=1)
    sig_bot = tb_df[tb_df["顶底信号"].str.contains("底部衰竭观察", na=False)]["ETF"].tolist() if len(tb_df) else []
    sig_top = tb_df[tb_df["顶底信号"].str.contains("顶部背离预警", na=False)]["ETF"].tolist() if len(tb_df) else []
    checks.append(f"顶底观察：底部衰竭观察 {len(sig_bot)} 只 {sig_bot if sig_bot else ''}；"
                  f"顶部背离预警 {len(sig_top)} 只 {sig_top if sig_top else ''}；"
                  f"位置分全部在0~100：{bool(((tb_df['位置分']>=0)&(tb_df['位置分']<=100)).all()) if len(tb_df) else '无数据'}")

    snapshot_final = snap[
        ["日期", "ETF", "名称", "趋势分", "信号", "位置分", "动量变化", "顶底信号", "综合研判", "创新低%", "创新高%",
         "覆盖权重%", "数据覆盖%", "最高分个股", "最低分个股"]]
    checks.append(f"快照日期:{snap_str}（{'盘中未收盘，不写历史' if is_intraday else '收盘'}）；分数全部在0~100：{bool(((snapshot_final['趋势分']>=0)&(snapshot_final['趋势分']<=100)).all())}")
    checks.append("持仓口径：" + "；".join(f"{c}:{n}" for c,n in hold_notes.items()))

    log("【4/4】生成Excel报告 ...")
    write_report(hist, snapshot_final, is_intraday, snap_str, changes, checks, tb_df)
    log("-"*60)
    for _, r in snapshot_final.sort_values("趋势分", ascending=False).iterrows():
        log(f"  {r['ETF']} {str(r['名称'])[:14]:<14} {r['趋势分']:>5} | {r['信号']} | 位置:{r['位置分']} | 动量变化:{r['动量变化']:.2f} | {r['顶底信号']} | 研判:{r['综合研判']} | 覆盖:{r['覆盖权重%']}% | 高:{r['最高分个股']} | 低:{r['最低分个股']}")
    if changes: log(f"⚠ 本次检测到 {len(changes)} 笔持仓变更，见报告「持仓变更记录」")
    if sig_bot: log(f"★ 底部衰竭观察：{'、'.join(sig_bot)}（详见「顶底观察」sheet）")
    if sig_top: log(f"★ 顶部背离预警：{'、'.join(sig_top)}（详见「顶底观察」sheet）")
    log(f"完成，耗时 {time.time()-t0:.0f} 秒")
    pass 


# ================= 以下为新增的 GitHub Pages 发布适配代码 =================
def generate_download_html(excel_filename):
    """生成一个带下载按钮的极简手机网页"""
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ETF监控报告下载</title>
    <style>
        body {{ font-family: sans-serif; text-align: center; padding: 50px 20px; background: #f4f4f9; }}
        .card {{ background: white; padding: 30px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        h2 {{ color: #333; }}
        .time {{ color: #888; font-size: 14px; margin-bottom: 30px; }}
        .download-btn {{ 
            display: inline-block; padding: 15px 40px; font-size: 18px; color: white; 
            background: #007bff; text-decoration: none; border-radius: 8px; font-weight: bold;
        }}
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 ETF趋势监控报告</h2>
        <p class="time">数据更新时间：{now_str}</p>
        <a href="{excel_filename}" class="download-btn" download>📥 下载 Excel 报告</a>
    </div>
</body>
</html>"""
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)

if __name__ == "__main__":
    main()
    
    # 1. 将生成的带中文名的Excel，复制一份为固定名称 report.xlsx (方便固定链接下载)
    if os.path.exists(REPORT):
        shutil.copy(REPORT, "report.xlsx")
        print("✅ 报告已复制为 report.xlsx")
        
        # 2. 生成带下载按钮的网页
        generate_download_html("report.xlsx")
        print("✅ 下载页 index.html 已生成")
    else:
        print("❌ 未找到生成的Excel报告，请检查运行日志")