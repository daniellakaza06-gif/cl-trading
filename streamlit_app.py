"""
Trade Intelligence — a Streamlit analytics suite for your weekly trading workbook.

Single file, no extra modules. Reads trading_data.xlsx sitting next to this file
(or any workbook uploaded from the sidebar) and rebuilds every number from the
trade rows themselves, then cross-checks it against your own daily summary tables.
"""
from __future__ import annotations

import datetime as dt
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

APP_DIR = Path(__file__).parent
GREEN, RED, GOLD, BLUE, MUTE = "#57d58b", "#ff7373", "#f5c542", "#6ea8fe", "#91a39a"
LINE, PANEL = "#213029", "#0d1713"
DIVERGING = [[0, RED], [0.5, "#16231d"], [1, GREEN]]
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
DAY_RE = re.compile(r"^(mon|tues?|wed(?:nes)?|thu(?:rs?)?|fri|sat(?:ur)?|sun)(day)?\b", re.I)
ABBR = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}


# ───────────────────────────── PARSER ─────────────────────────────
def norm(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(x).lower())


def num(v) -> float:
    if v is None or isinstance(v, bool):
        return np.nan
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def day_index(cell):
    """0-6 if the cell is a weekday label ('Monday', 'Tuesday 11 aug 2026'), else None."""
    if not isinstance(cell, str) or len(cell.strip()) > 45:
        return None
    m = DAY_RE.match(cell.strip())
    return ABBR[m.group(1)[:3].lower()] if m else None


def label_date(cell):
    """Explicit date inside a day label, e.g. 'Monday 10 aug 2026'."""
    m = re.search(r"(\d{1,2})\s*([A-Za-z]{3,9})\.?\s*(\d{4})?", str(cell))
    if not m:
        return None
    mon = MONTHS.get(m.group(2)[:3].lower())
    if not mon:
        return None
    year = int(m.group(3)) if m.group(3) else dt.date.today().year
    try:
        return dt.date(year, mon, int(m.group(1)))
    except ValueError:
        return None


def week_range(sheet: str):
    """Best-effort start/end date from names like 'Week 1 31 Aug.- 4 Sept.' or 'Week 2(10-14 Aug. 26)'."""
    s = sheet.replace(".", " ")
    m = re.search(r"(\d{1,2})\s*([A-Za-z]{3,9})?\s*[-–]\s*(\d{1,2})\s*([A-Za-z]{3,9})", s)
    if not m:
        return None, None
    d1, m1, d2, m2 = m.groups()
    mon2 = MONTHS.get(m2[:3].lower())
    mon1 = MONTHS.get(m1[:3].lower()) if m1 else mon2
    if not mon1 or not mon2:
        return None, None
    ym = re.search(r"(20\d{2})", s)
    year = int(ym.group(1)) if ym else dt.date.today().year
    try:
        return dt.date(year, mon1, int(d1)), dt.date(year, mon2, int(d2))
    except ValueError:
        return None, None


def canon_col(n: str):
    if n in ("trade", "trades"):
        return "trade"
    if n == "coin":
        return "coin"
    if n.startswith("capitalallocated"):
        return "capital"
    if "stoploss" in n:
        return "stop"
    if "profittarget" in n:
        return "target"
    if n.startswith("actualprofit"):
        return "profit_zar" if "zar" in n else "profit_pct"
    if n.startswith("actualloss"):
        return "loss_zar" if "zar" in n else "loss_pct"
    if n.startswith("total"):
        return "total_zar"
    return None


def canon_metric(n: str):
    if n.startswith("myactualbalance"):
        return "balance"
    if n.startswith("totaltrades") or n.startswith("noofexecutedtrades"):
        return "trades"
    if n.startswith("noofwinning"):
        return "wins"
    if n.startswith("nooflosing"):
        return "losses"
    if n in ("actualprofitlosszar", "actualplzar", "profitafterlosszar"):
        return "net"
    if n in ("capitaltradedtodayzar", "totalcapitaltradedtodayzar", "capitalusedzar"):
        return "capital_used"
    if n.startswith("mytargetscore"):
        return "target_score"
    return None


def to_date(v, monday):
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, dt.date):
        return v
    di = day_index(v)
    if di is not None and monday is not None:
        return monday + dt.timedelta(days=di)
    return None


def parse_summary(rows, monday):
    """Reads the 'Metric | Mon | Tue | …' table at the top of a sheet."""
    out, dates = [], []
    hi = next((i for i, r in enumerate(rows) if norm(r[0]) == "metric"), None)
    if hi is None:
        return out, dates
    colmap = {}
    for j, v in enumerate(rows[hi][1:], 1):
        d = to_date(v, monday)
        if d is not None:
            colmap[j] = d
    dates = [colmap[j] for j in sorted(colmap)]
    for r in rows[hi + 1:]:
        first = r[0]
        if first is None or day_index(first) is not None or norm(first) in ("trade", "trades"):
            break
        key = canon_metric(norm(first))
        if not key:
            continue
        for j, d in colmap.items():
            v = num(r[j]) if j < len(r) else np.nan
            if not np.isnan(v):
                out.append({"Date": d, "Metric": key, "Value": v})
    return out, dates


def parse_sheet_blocks(rows):
    """Finds every 'Day label → trade table' block on a sheet."""
    blocks, cur, i = [], None, 0
    while i < len(rows):
        first = rows[i][0]
        di = day_index(first)
        if di is not None:
            cur = (di, label_date(first))
            i += 1
            continue
        if norm(first) in ("trade", "trades"):
            colmap = {j: canon_col(norm(h)) for j, h in enumerate(rows[i]) if canon_col(norm(h))}
            coin_j = next((j for j, k in colmap.items() if k == "coin"), 1)
            j, blk = i + 1, []
            while j < len(rows):
                f0 = rows[j][0]
                if day_index(f0) is not None or norm(f0) in ("trade", "trades"):
                    break
                if not np.isnan(num(f0)) and rows[j][coin_j] not in (None, ""):
                    blk.append(rows[j])
                j += 1
            blocks.append({"day": cur, "colmap": colmap, "rows": blk})
            i = j
            continue
        i += 1
    return blocks


def parse_workbook(data: bytes):
    xl = pd.ExcelFile(io.BytesIO(data))
    trades, summ, issues, inventory, raw = [], [], [], [], {}

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)
        raw[sheet] = df
        rows = df.astype(object).where(df.notna(), None).values.tolist()
        is_model = "profitab" in sheet.lower()
        start, _ = week_range(sheet)
        monday = (start - dt.timedelta(days=start.weekday())) if start else None

        s_rows, s_dates = parse_summary(rows, monday)
        for r in s_rows:
            r["Sheet"] = sheet
        if not is_model:
            summ += s_rows

        n_before = len(trades)
        if not is_model:
            blocks = parse_sheet_blocks(rows)
            idxs = [b["day"][0] if b["day"] else None for b in blocks]
            in_order = all(x is not None for x in idxs) and all(a < b for a, b in zip(idxs, idxs[1:]))
            by_order = (not in_order) and s_dates and len(s_dates) == len(blocks)
            if not in_order:
                labels = ", ".join(WEEKDAYS[x][:3] if x is not None else "?" for x in idxs)
                issues.append({"Sheet": sheet, "Severity": "Fixed",
                               "Message": f"Day labels out of order or repeated ({labels}). "
                                          + ("Re-sequenced by block order using the summary table dates."
                                             if by_order else "Re-sequenced to the next free weekday.")})
            prev = -1
            for k, b in enumerate(blocks):
                di, ld = b["day"] if b["day"] else (None, None)
                if by_order:
                    date = s_dates[k]
                else:
                    if di is None or di <= prev:
                        di = prev + 1
                    prev = di
                    date = ld or (monday + dt.timedelta(days=di) if monday else None)
                for r in b["rows"]:
                    rec = {k2: r[j] for j, k2 in b["colmap"].items() if j < len(r)}
                    cap = num(rec.get("capital"))
                    p, l = num(rec.get("profit_zar")), num(rec.get("loss_zar"))
                    pp, lp = num(rec.get("profit_pct")), num(rec.get("loss_pct"))
                    if np.isnan(p) and not np.isnan(pp) and not np.isnan(cap):
                        p = pp * cap
                    if np.isnan(l) and not np.isnan(lp) and not np.isnan(cap):
                        l = -abs(lp) * cap
                    p = 0.0 if np.isnan(p) else p
                    l = 0.0 if np.isnan(l) else -abs(l)
                    coin = str(rec.get("coin", "")).strip()
                    m = re.search(r"\(([^)]+)\)", coin)
                    trades.append({
                        "Sheet": sheet, "Date": pd.Timestamp(date) if date else pd.NaT,
                        "Trade": int(num(rec.get("trade"))), "Coin": coin,
                        "Ticker": (m.group(1) if m else coin).upper(),
                        "Capital (ZAR)": cap, "Stop Loss %": num(rec.get("stop")),
                        "Profit Target %": num(rec.get("target")),
                        "Actual Profit %": pp, "Actual Profit (ZAR)": p,
                        "Actual Loss %": -abs(lp) if not np.isnan(lp) else np.nan,
                        "Actual Loss (ZAR)": l, "P/L (ZAR)": p + l,
                        "WeekStart": pd.Timestamp(monday) if monday else pd.NaT,
                    })
        inventory.append({"Sheet": sheet, "Type": "Profitability model" if is_model else "Weekly trades",
                          "Rows": len(df), "Columns": df.shape[1],
                          "Trade rows parsed": len(trades) - n_before,
                          "Summary metrics found": len({r['Metric'] for r in s_rows})})

    tdf = pd.DataFrame(trades)
    sdf = pd.DataFrame(summ)
    if not tdf.empty:
        tdf["Date"] = pd.to_datetime(tdf["Date"])
        tdf["Day"] = tdf["Date"].dt.day_name().fillna("Unspecified")
        tdf["Return %"] = tdf["P/L (ZAR)"] / tdf["Capital (ZAR)"]
        tdf["Outcome"] = np.select([tdf["P/L (ZAR)"] > 0, tdf["P/L (ZAR)"] < 0], ["Win", "Loss"], "Flat")
        starts = sorted(tdf["WeekStart"].dropna().unique())
        names = {s: f"W{i} · {pd.Timestamp(s):%d %b}" for i, s in enumerate(starts, 1)}
        tdf["Week"] = tdf["WeekStart"].map(names).fillna(tdf["Sheet"])
        tdf = tdf.sort_values(["Date", "Trade"]).reset_index(drop=True)
    if not sdf.empty:
        sdf["Date"] = pd.to_datetime(sdf["Date"])
        sdf = sdf.drop_duplicates(["Sheet", "Date", "Metric"])
    return {"trades": tdf, "summary": sdf, "issues": pd.DataFrame(issues),
            "inventory": pd.DataFrame(inventory), "raw": raw, "sheets": xl.sheet_names}


@st.cache_data(show_spinner="Reading workbook…")
def load(data: bytes):
    return parse_workbook(data)


def reconcile(trades: pd.DataFrame, summ: pd.DataFrame) -> pd.DataFrame:
    """Compares what the trade rows add up to with the daily summary table in the same sheet."""
    if trades.empty or summ.empty:
        return pd.DataFrame()
    g = trades.groupby(["Sheet", "Date"]).agg(
        Trades=("Trade", "count"), Wins=("Outcome", lambda x: (x == "Win").sum()),
        Losses=("Outcome", lambda x: (x == "Loss").sum()), Net=("P/L (ZAR)", "sum")).reset_index()
    p = summ.pivot_table(index=["Sheet", "Date"], columns="Metric", values="Value", aggfunc="first").reset_index()
    m = g.merge(p, on=["Sheet", "Date"], how="left")
    out = []
    for _, r in m.iterrows():
        checks, notes = [], []
        for mine, theirs, tol, label in (("Trades", "trades", 0, "trade count"), ("Wins", "wins", 0, "wins"),
                                         ("Losses", "losses", 0, "losses"), ("Net", "net", 1.5, "net P/L")):
            if theirs in m.columns and not pd.isna(r.get(theirs)):
                ok = abs(r[mine] - r[theirs]) <= tol
                checks.append(ok)
                if not ok:
                    notes.append(f"{label}: rows {r[mine]:,.0f} vs summary {r[theirs]:,.0f}")
        status = "No summary" if not checks else ("Match" if all(checks) else "Check")
        out.append({"Sheet": r["Sheet"], "Date": r["Date"], "Rows: trades": r["Trades"],
                    "Summary: trades": r.get("trades", np.nan), "Rows: net (ZAR)": r["Net"],
                    "Summary: net (ZAR)": r.get("net", np.nan), "Status": status, "Detail": "; ".join(notes)})
    return pd.DataFrame(out)


# ───────────────────────────── ANALYTICS ─────────────────────────────
def rz(x, d=0):
    if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
        return "—"
    return f"{'-' if x < 0 else ''}R {abs(x):,.{d}f}"


def streaks(outcomes):
    bw = bl = cw = cl = 0
    for o in outcomes:
        if o == "Win":
            cw, cl = cw + 1, 0
        elif o == "Loss":
            cl, cw = cl + 1, 0
        else:
            cw = cl = 0
        bw, bl = max(bw, cw), max(bl, cl)
    return bw, bl


def daily_table(f: pd.DataFrame) -> pd.DataFrame:
    d = f.dropna(subset=["Date"]).groupby("Date").agg(
        Trades=("Trade", "count"), Wins=("Outcome", lambda x: (x == "Win").sum()),
        Losses=("Outcome", lambda x: (x == "Loss").sum()), Net=("P/L (ZAR)", "sum"),
        Capital=("Capital (ZAR)", "sum")).reset_index().sort_values("Date")
    d["Win rate"] = d["Wins"] / d["Trades"]
    d["Return %"] = d["Net"] / d["Capital"]
    d["Cum"] = d["Net"].cumsum()
    d["Peak"] = np.maximum(d["Cum"].cummax(), 0)
    d["DD"] = d["Cum"] - d["Peak"]
    d["Cum %"] = d["Return %"].cumsum()
    return d


def kpis(f: pd.DataFrame, d: pd.DataFrame) -> dict:
    n = len(f)
    wins, losses = int((f.Outcome == "Win").sum()), int((f.Outcome == "Loss").sum())
    pl = f["P/L (ZAR)"]
    gp, gl = pl[pl > 0].sum(), -pl[pl < 0].sum()
    avg_win, avg_loss = (gp / wins if wins else 0.0), (gl / losses if losses else 0.0)
    bw, bl = streaks(f.sort_values(["Date", "Trade"]).Outcome)
    return {
        "n": n, "wins": wins, "losses": losses, "net": pl.sum(), "gp": gp, "gl": gl,
        "winrate": wins / n if n else 0, "pf": gp / gl if gl else np.inf,
        "avg_win": avg_win, "avg_loss": avg_loss, "expectancy": pl.sum() / n if n else 0,
        "payoff": avg_win / avg_loss if avg_loss else np.inf,
        "breakeven": avg_loss / (avg_win + avg_loss) if (avg_win + avg_loss) else np.nan,
        "maxdd": d["DD"].min() if len(d) else 0, "capital": f["Capital (ZAR)"].sum(),
        "best_streak": bw, "worst_streak": bl, "days": len(d),
    }


def build_insights(f, d, k):
    out = []
    if k["wins"] + k["losses"]:
        gap = (k["winrate"] - k["breakeven"]) * 100
        out.append(("good" if gap > 10 else "warn", "Your edge in one line",
                    f"You win {k['winrate']:.0%} of trades. Average win is {rz(k['avg_win'])}, average loss {rz(k['avg_loss'])}. "
                    f"Break-even win rate is {k['breakeven']:.0%}, so you're running {abs(gap):.0f} pts "
                    f"{'above' if gap >= 0 else 'below'} it."))
    if len(d) >= 4:
        best, worst = d.loc[d.Net.idxmax()], d.loc[d.Net.idxmin()]
        out.append(("info", "Best and worst day",
                    f"Best: {best.Date:%a %d %b} at {rz(best.Net)} ({best.Wins:.0f}/{best.Trades:.0f} won). "
                    f"Worst: {worst.Date:%a %d %b} at {rz(worst.Net)} ({worst.Wins:.0f}/{worst.Trades:.0f} won)."))
    wd = f.dropna(subset=["Date"]).groupby("Day").agg(Net=("P/L (ZAR)", "sum"), WR=("Outcome", lambda x: (x == "Win").mean()),
                                                    N=("Trade", "count"))
    wd = wd[wd.N >= 8]
    if len(wd) >= 3:
        b, w = wd.WR.idxmax(), wd.WR.idxmin()
        out.append(("info", "Weekday pattern",
                    f"{b} is your strongest day by win rate ({wd.WR[b]:.0%}); {w} is the weakest ({wd.WR[w]:.0%}). "
                    "Worth watching, but each weekday is only a handful of sessions."))
    if len(d) >= 8:
        med = d.Trades.median()
        hi, lo = d[d.Trades > med], d[d.Trades <= med]
        if len(hi) and len(lo):
            wr_hi, wr_lo = hi.Wins.sum() / hi.Trades.sum(), lo.Wins.sum() / lo.Trades.sum()
            out.append(("warn" if wr_hi < wr_lo - 0.05 else "info", "Volume vs quality",
                        f"On days with more than {med:.0f} trades your win rate was {wr_hi:.0%}, versus {wr_lo:.0%} on lighter days."))
    ls = f[(f.Outcome == "Loss") & (f["Stop Loss %"] > 0)]
    if len(ls) >= 3:
        over = (-ls["P/L (ZAR)"] / ls["Capital (ZAR)"]) / ls["Stop Loss %"]
        share = (over > 1.5).mean()
        excess = (-ls["P/L (ZAR)"] - ls["Stop Loss %"] * ls["Capital (ZAR)"]).clip(lower=0).sum()
        out.append(("warn" if share > 0.2 else "good", "Stop discipline",
                    f"The median loss is {over.median():.1f}× the planned stop, and {share:.0%} of losses ran beyond 1.5× it. "
                    f"Losses past the planned stop cost an extra {rz(excess)} in total."))
    if k["gp"] > 0 and len(f):
        top1 = f["P/L (ZAR)"].max()
        by_coin = f.groupby("Ticker")["P/L (ZAR)"].sum().sort_values(ascending=False)
        out.append(("info", "Profit concentration",
                    f"Your single best trade contributed {top1 / k['gp']:.0%} of gross profit. The top 3 coins ({', '.join(by_coin.index[:3])}) "
                    f"hold {by_coin.head(3).sum() / k['gp']:.0%}."))
    if len(d) >= 10:
        h = len(d) // 2
        a, b = d.iloc[:h], d.iloc[h:]
        wa, wb = a.Wins.sum() / a.Trades.sum(), b.Wins.sum() / b.Trades.sum()
        out.append(("good" if wb >= wa else "warn", "Trend",
                    f"Win rate moved from {wa:.0%} in the first half of the period to {wb:.0%} in the second; "
                    f"average daily return went from {a.Net.sum() / a.Capital.sum():.2%} to {b.Net.sum() / b.Capital.sum():.2%} of capital deployed."))
    out.append(("info", "Streaks", f"Longest winning run: {k['best_streak']} trades. Longest losing run: {k['worst_streak']} trades."))
    return out


# ───────────────────────────── UI HELPERS ─────────────────────────────
def show(fig, height=None, key=None):
    fig.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, system-ui, sans-serif", color="#c9d6cf", size=12),
        margin=dict(l=10, r=10, t=50, b=10), title_font=dict(size=15),
        legend=dict(orientation="h", y=1.08, x=0, bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="#0b1411", bordercolor=LINE),
    )
    fig.update_xaxes(gridcolor="#16231d", zeroline=False)
    fig.update_yaxes(gridcolor="#16231d", zeroline=False)
    if height:
        fig.update_layout(height=height)
    try:
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key=key)
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=key)


def table(df, **kw):
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kw)
    except TypeError:
        st.dataframe(df, use_container_width=True, hide_index=True, **kw)


def card(label, value, sub="", cls=""):
    st.markdown(f'<div class="kpi"><div class="kpi-l">{label}</div><div class="kpi-v {cls}">{value}</div>'
                f'<div class="kpi-s">{sub}</div></div>', unsafe_allow_html=True)


def insight_html(items):
    icon = {"good": "▲", "warn": "▼", "info": "◆"}
    return "".join(f'<div class="ins {t}"><div class="ins-h"><span>{icon[t]}</span> {h}</div><div class="ins-b">{b}</div></div>'
                   for t, h, b in items)


CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] {font-family: 'Inter', system-ui, sans-serif;}
.stApp{background:#07100c;color:#e6efe9}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp label,.stApp p,.stApp li,[data-testid="stMarkdownContainer"],[data-testid="stCaptionContainer"]{color:#dbe6df}
.stApp [data-testid="stCaptionContainer"]{color:#8fa198}
[data-testid="stHeader"]{background:transparent}
.stTabs [data-baseweb="tab"]{color:#b6c5bd}
.block-container{padding-top:1.2rem;padding-bottom:3rem;max-width:1500px}
[data-testid="stSidebar"]{background:#0b1411;border-right:1px solid #1c2a24}
h1,h2,h3{letter-spacing:-.03em}
.eyebrow{font-size:.7rem;letter-spacing:.16em;text-transform:uppercase;color:#91a39a;font-weight:700}
.hero{padding:1.5rem 1.7rem;border:1px solid #213029;border-radius:20px;margin-bottom:1.1rem;
 background:radial-gradient(1200px 300px at 90% -20%,rgba(87,213,139,.14),transparent),linear-gradient(135deg,#101b17 0%,#09110e 70%)}
.hero h1{font-size:2.3rem;margin:.2rem 0 .35rem;font-weight:800}
.hero p{color:#9fb0a7;margin:0;max-width:820px;font-size:.95rem}
.chips{margin-top:.9rem;display:flex;gap:.5rem;flex-wrap:wrap}
.chip{border:1px solid #213029;background:#0d1713;border-radius:999px;padding:.22rem .75rem;font-size:.74rem;color:#b6c5bd}
.chip b{color:#f5c542;font-weight:700}
.kpi{border:1px solid #213029;background:linear-gradient(180deg,#0f1b16,#0b1410);border-radius:16px;padding:1rem 1.1rem;min-height:116px;
 transition:border-color .2s,transform .2s}
.kpi:hover{border-color:#2f5a45;transform:translateY(-2px)}
.kpi-l{font-size:.68rem;color:#91a39a;text-transform:uppercase;letter-spacing:.12em;font-weight:700}
.kpi-v{font-size:1.65rem;font-weight:800;margin-top:.3rem;letter-spacing:-.02em}
.kpi-s{font-size:.76rem;color:#789086;margin-top:.15rem}
.good{color:#57d58b}.bad{color:#ff7373}.accent{color:#f5c542}.muted{color:#91a39a}
.ins{border:1px solid #213029;border-left:3px solid #6ea8fe;background:#0d1713;border-radius:12px;padding:.7rem .9rem;margin-bottom:.6rem}
.ins.good{border-left-color:#57d58b;color:inherit}.ins.warn{border-left-color:#f5c542;color:inherit}
.ins-h{font-weight:700;font-size:.82rem;color:#e6efe9}.ins-b{font-size:.8rem;color:#9fb0a7;margin-top:.15rem;line-height:1.45}
.ins.good .ins-h span{color:#57d58b}.ins.warn .ins-h span{color:#f5c542}.ins.info .ins-h span{color:#6ea8fe}
.section{font-size:.7rem;letter-spacing:.16em;text-transform:uppercase;color:#91a39a;font-weight:700;margin:1.4rem 0 .3rem}
div[data-testid="stDataFrame"]{border:1px solid #213029;border-radius:14px;overflow:hidden}
</style>
"""


def section(txt):
    st.markdown(f'<div class="section">{txt}</div>', unsafe_allow_html=True)


# ───────────────────────────── PAGES ─────────────────────────────
def page_command(f, d, k, summ, week_order):
    c = st.columns(6)
    with c[0]:
        card("Net P/L", rz(k["net"]), f"{rz(k['net'] / max(k['days'], 1))} per trading day", "good" if k["net"] >= 0 else "bad")
    with c[1]:
        card("Win rate", f"{k['winrate']:.1%}", f"{k['wins']} wins · {k['losses']} losses", "accent")
    with c[2]:
        card("Profit factor", "∞" if np.isinf(k["pf"]) else f"{k['pf']:.2f}", "Gross profit ÷ gross loss")
    with c[3]:
        card("Expectancy / trade", rz(k["expectancy"], 1), f"Payoff ratio {k['payoff']:.2f}×" if not np.isinf(k["payoff"]) else "No losing trades")
    with c[4]:
        card("Max drawdown", rz(k["maxdd"]), "Peak-to-trough on cumulative P/L", "bad" if k["maxdd"] < 0 else "")
    with c[5]:
        card("Trades", f"{k['n']:,}", f"{k['days']} sessions · {f.Ticker.nunique()} coins")

    mode = st.radio("Equity view", ["ZAR", "% of capital deployed"], horizontal=True, label_visibility="collapsed")
    left, right = st.columns([2.1, 1])
    with left:
        col, ddcol, fmt = ("Cum", "DD", "R %{y:,.0f}") if mode == "ZAR" else ("Cum %", None, "%{y:.1%}")
        x = [d.Date.iloc[0] - pd.Timedelta(days=1)] + list(d.Date)
        y = [0] + list(d[col])
        fig = make_subplots(rows=2 if ddcol else 1, cols=1, shared_xaxes=True,
                            row_heights=[0.74, 0.26] if ddcol else [1], vertical_spacing=0.04)
        fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name="Cumulative", line=dict(color=GREEN, width=3, shape="spline"),
                                 marker=dict(size=5), fill="tozeroy", fillcolor="rgba(87,213,139,.10)",
                                 hovertemplate="%{x|%a %d %b}<br>" + fmt + "<extra></extra>"), 1, 1)
        if ddcol:
            fig.add_trace(go.Bar(x=d.Date, y=d.DD, name="Drawdown", marker_color=RED, opacity=.8,
                                 hovertemplate="%{x|%a %d %b}<br>R %{y:,.0f}<extra></extra>"), 2, 1)
        fig.update_layout(title="Equity curve" + (" & drawdown" if ddcol else ""), showlegend=False)
        show(fig, 470)
    with right:
        section("What the data says")
        st.markdown(insight_html(build_insights(f, d, k)[:5]), unsafe_allow_html=True)

    a, b = st.columns([1.7, 1])
    with a:
        piv = f.dropna(subset=["Date"]).assign(Day=lambda x: x.Date.dt.day_name()).pivot_table(
            index="Week", columns="Day", values="P/L (ZAR)", aggfunc="sum")
        cols = [w for w in WEEKDAYS if w in piv.columns]
        piv = piv.reindex(index=[w for w in week_order if w in piv.index], columns=cols)
        txt = piv.map(lambda v: "" if pd.isna(v) else rz(v)) if hasattr(piv, "map") else piv.applymap(lambda v: "" if pd.isna(v) else rz(v))
        fig = go.Figure(go.Heatmap(z=piv.values, x=[c_[:3] for c_ in cols], y=piv.index, text=txt.values, texttemplate="%{text}",
                                   colorscale=DIVERGING, zmid=0, showscale=False, xgap=4, ygap=4,
                                   hovertemplate="%{y} · %{x}<br>%{text}<extra></extra>"))
        fig.update_layout(title="P/L calendar", yaxis=dict(autorange="reversed"))
        show(fig, 130 + 52 * len(piv))
    with b:
        by = f.groupby("Ticker").agg(Net=("P/L (ZAR)", "sum")).sort_values("Net")
        pick = pd.concat([by.head(6), by.tail(6)]).drop_duplicates().sort_values("Net")
        fig = go.Figure(go.Bar(x=pick.Net, y=pick.index, orientation="h", marker_color=[GREEN if v >= 0 else RED for v in pick.Net],
                               hovertemplate="%{y}: R %{x:,.0f}<extra></extra>"))
        fig.update_layout(title="Top contributors and detractors")
        show(fig, 130 + 52 * max(len(piv), 5))


def page_performance(f, d, k, summ, week_order):
    wk = f.groupby("Week", sort=False).agg(
        Dates=("Date", lambda x: f"{x.min():%d %b} – {x.max():%d %b}"), Trades=("Trade", "count"),
        Wins=("Outcome", lambda x: (x == "Win").sum()), Net=("P/L (ZAR)", "sum"), Capital=("Capital (ZAR)", "sum"),
        Sessions=("Date", "nunique")).reset_index()
    wk["Win rate"] = wk.Wins / wk.Trades
    wk["Return on deployed"] = wk.Net / wk.Capital
    gp = f[f["P/L (ZAR)"] > 0].groupby("Week")["P/L (ZAR)"].sum()
    gl = -f[f["P/L (ZAR)"] < 0].groupby("Week")["P/L (ZAR)"].sum()
    wk["Profit factor"] = wk.Week.map(gp / gl.replace(0, np.nan))
    wk["Trades / day"] = wk.Trades / wk.Sessions

    section("Weekly scorecard")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=wk.Week, y=wk.Net, name="Net P/L (ZAR)", marker_color=[GREEN if v >= 0 else RED for v in wk.Net],
                         hovertemplate="%{x}<br>R %{y:,.0f}<extra></extra>"), secondary_y=False)
    fig.add_trace(go.Scatter(x=wk.Week, y=wk["Win rate"], name="Win rate", mode="lines+markers", line=dict(color=GOLD, width=3),
                             hovertemplate="%{y:.0%}<extra></extra>"), secondary_y=True)
    fig.update_yaxes(tickformat=".0%", range=[0, 1.05], secondary_y=True, showgrid=False)
    fig.update_layout(title="Net P/L and win rate by week")
    show(fig, 400)
    wt = wk[["Week", "Dates", "Sessions", "Trades", "Win rate", "Net", "Capital", "Return on deployed", "Profit factor"]].rename(
        columns={"Net": "Net P/L (ZAR)", "Capital": "Capital deployed (ZAR)", "Win rate": "Win rate (%)", "Return on deployed": "Return on deployed (%)"})
    wt["Win rate (%)"] *= 100
    wt["Return on deployed (%)"] *= 100
    table(wt, column_config={"Win rate (%)": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%"),
                             "Net P/L (ZAR)": st.column_config.NumberColumn(format="R %.0f"),
                             "Capital deployed (ZAR)": st.column_config.NumberColumn(format="R %.0f"),
                             "Return on deployed (%)": st.column_config.NumberColumn(format="%.2f%%"),
                             "Profit factor": st.column_config.NumberColumn(format="%.2f")})

    section("Day-level behaviour")
    a, b = st.columns(2)
    with a:
        dd = d.copy()
        fig = px.scatter(dd, x="Trades", y="Win rate", size="Capital", color="Net", color_continuous_scale=DIVERGING,
                         color_continuous_midpoint=0, hover_data={"Date": "|%a %d %b", "Net": ":,.0f", "Capital": ":,.0f"},
                         title="Do more trades mean worse trades?")
        if len(dd) >= 5 and dd.Trades.nunique() > 1:
            z = np.polyfit(dd.Trades, dd["Win rate"], 1)
            xs = np.linspace(dd.Trades.min(), dd.Trades.max(), 20)
            fig.add_trace(go.Scatter(x=xs, y=np.polyval(z, xs), mode="lines", name="Trend", line=dict(color=GOLD, dash="dot")))
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(coloraxis_showscale=False)
        show(fig, 420)
    with b:
        wd = f.dropna(subset=["Date"]).groupby("Day").agg(Net=("P/L (ZAR)", "sum"), WR=("Outcome", lambda x: (x == "Win").mean()),
                                                          N=("Trade", "count")).reindex(WEEKDAYS).dropna()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=[w[:3] for w in wd.index], y=wd.Net, name="Net P/L", marker_color=[GREEN if v >= 0 else RED for v in wd.Net]), secondary_y=False)
        fig.add_trace(go.Scatter(x=[w[:3] for w in wd.index], y=wd.WR, name="Win rate", mode="lines+markers", line=dict(color=GOLD, width=3)), secondary_y=True)
        fig.update_yaxes(tickformat=".0%", range=[0, 1.05], secondary_y=True, showgrid=False)
        fig.update_layout(title="Weekday edge")
        show(fig, 420)

    section("Week vs week")
    g = f.dropna(subset=["Date"]).groupby(["Week", "Date"], sort=False)["P/L (ZAR)"].sum().reset_index()
    g["Session"] = g.groupby("Week").cumcount() + 1
    g["Cumulative"] = g.groupby("Week")["P/L (ZAR)"].cumsum()
    fig = px.line(g, x="Session", y="Cumulative", color="Week", markers=True, title="Cumulative P/L by session, every week overlaid",
                  color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_layout(xaxis=dict(dtick=1, title="Trading day of the week"), yaxis_title="Cumulative P/L (ZAR)")
    show(fig, 400)

    bal = summ[summ.Metric == "balance"].sort_values("Date") if len(summ) else pd.DataFrame()
    if len(bal) >= 2:
        section("Account balance (from your workbook)")
        fig = go.Figure(go.Scatter(x=bal.Date, y=bal.Value, mode="lines+markers", line=dict(color=BLUE, width=3, shape="spline"),
                                   fill="tozeroy", fillcolor="rgba(110,168,254,.08)", hovertemplate="%{x|%a %d %b}<br>R %{y:,.0f}<extra></extra>"))
        fig.update_yaxes(range=[bal.Value.min() * 0.97, bal.Value.max() * 1.02])
        fig.update_layout(title=f"Balance: {rz(bal.Value.iloc[0])} → {rz(bal.Value.iloc[-1])}")
        show(fig, 340)


def page_risk(f, d, k, summ, week_order):
    c = st.columns(5)
    with c[0]:
        card("Average win", rz(k["avg_win"]), "Per winning trade", "good")
    with c[1]:
        card("Average loss", rz(-k["avg_loss"]), "Per losing trade", "bad")
    with c[2]:
        card("Break-even win rate", f"{k['breakeven']:.0%}" if not np.isnan(k["breakeven"]) else "—", f"You're at {k['winrate']:.0%}", "accent")
    with c[3]:
        card("Longest losing run", f"{k['worst_streak']}", "Consecutive losing trades", "bad")
    with c[4]:
        card("Worst session", rz(d.Net.min()) if len(d) else "—", "Single-day net P/L", "bad")

    section("Planned vs realised")
    a, b = st.columns(2)
    ls = f[(f.Outcome == "Loss") & (f["Stop Loss %"] > 0)].copy()
    ws = f[(f.Outcome == "Win") & (f["Profit Target %"] > 0)].copy()
    with a:
        if len(ls):
            ls["x"] = (-ls["P/L (ZAR)"] / ls["Capital (ZAR)"]) / ls["Stop Loss %"]
            fig = px.histogram(ls, x="x", nbins=24, title="Losses vs planned stop (1× = stop honoured)", color_discrete_sequence=[RED])
            fig.add_vline(x=1, line_dash="dot", line_color=GOLD, annotation_text="planned stop")
            fig.update_layout(xaxis_title="Realised loss ÷ planned stop", yaxis_title="Trades", bargap=.05)
            show(fig, 380)
        else:
            st.info("No losing trades with a planned stop in this selection.")
    with b:
        if len(ws):
            ws["x"] = (ws["P/L (ZAR)"] / ws["Capital (ZAR)"]) / ws["Profit Target %"]
            fig = px.histogram(ws, x="x", nbins=24, title="Wins vs planned target (1× = target hit)", color_discrete_sequence=[GREEN])
            fig.add_vline(x=1, line_dash="dot", line_color=GOLD, annotation_text="planned target")
            fig.update_layout(xaxis_title="Realised gain ÷ planned target", yaxis_title="Trades", bargap=.05)
            show(fig, 380)
        else:
            st.info("No winning trades with a planned target in this selection.")

    section("Distribution and extremes")
    a, b = st.columns([1.2, 1])
    with a:
        fig = px.histogram(f, x="Return %", color="Outcome", nbins=40, title="Return per trade (on capital allocated)",
                           color_discrete_map={"Win": GREEN, "Loss": RED, "Flat": MUTE})
        fig.update_xaxes(tickformat=".1%")
        fig.update_layout(bargap=.05, barmode="overlay")
        fig.update_traces(opacity=.8)
        show(fig, 380)
    with b:
        ex = pd.concat([f.nsmallest(5, "P/L (ZAR)"), f.nlargest(5, "P/L (ZAR)")]).sort_values("P/L (ZAR)")
        ex["Label"] = ex.Ticker + " · " + ex.Date.dt.strftime("%d %b")
        fig = go.Figure(go.Bar(x=ex["P/L (ZAR)"], y=ex.Label, orientation="h", marker_color=[GREEN if v >= 0 else RED for v in ex["P/L (ZAR)"]]))
        fig.update_layout(title="Five best and five worst trades")
        show(fig, 380)

    section("Drawdown history")
    fig = go.Figure(go.Scatter(x=d.Date, y=d.DD, fill="tozeroy", line=dict(color=RED), fillcolor="rgba(255,115,115,.15)",
                               hovertemplate="%{x|%a %d %b}<br>R %{y:,.0f}<extra></extra>"))
    fig.update_layout(title=f"Underwater curve (max {rz(k['maxdd'])})")
    show(fig, 300)


def page_coins(f, d, k, summ, week_order):
    mn = st.slider("Minimum trades per coin", 1, int(max(f.groupby("Ticker").size().max(), 2)), 3)
    coin = f.groupby("Ticker").agg(Trades=("Trade", "count"), Wins=("Outcome", lambda x: (x == "Win").sum()),
                                   Net=("P/L (ZAR)", "sum"), Capital=("Capital (ZAR)", "sum"), Best=("P/L (ZAR)", "max"),
                                   Worst=("P/L (ZAR)", "min")).reset_index()
    coin["Win rate"] = coin.Wins / coin.Trades
    coin["Avg P/L"] = coin.Net / coin.Trades
    coin = coin[coin.Trades >= mn].sort_values("Net", ascending=False)
    if coin.empty:
        st.info("No coin meets that minimum.")
        return
    a, b = st.columns([1.2, 1])
    with a:
        fig = px.treemap(coin, path=["Ticker"], values="Capital", color="Net", color_continuous_scale=DIVERGING,
                         color_continuous_midpoint=0, hover_data={"Trades": True, "Win rate": ":.0%", "Net": ":,.0f"},
                         title="Where the capital went (size) and what it earned (colour)")
        fig.update_layout(coloraxis_showscale=False)
        show(fig, 470)
    with b:
        fig = px.scatter(coin, x="Trades", y="Win rate", size="Capital", color="Net", color_continuous_scale=DIVERGING,
                         color_continuous_midpoint=0, hover_name="Ticker", title="Frequency vs reliability")
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(coloraxis_showscale=False)
        show(fig, 470)
    section("Leaderboard")
    coin = coin.assign(**{"Win rate": coin["Win rate"] * 100})
    table(coin, column_config={
        "Win rate": st.column_config.ProgressColumn("Win rate (%)", min_value=0, max_value=100, format="%.0f%%"),
        "Net": st.column_config.NumberColumn("Net P/L", format="R %.0f"), "Capital": st.column_config.NumberColumn(format="R %.0f"),
        "Avg P/L": st.column_config.NumberColumn(format="R %.1f"), "Best": st.column_config.NumberColumn(format="R %.0f"),
        "Worst": st.column_config.NumberColumn(format="R %.0f")})

    section("Coin deep-dive")
    pick = st.selectbox("Coin", coin.Ticker.tolist(), label_visibility="collapsed")
    t = f[f.Ticker == pick]
    c = st.columns(4)
    with c[0]:
        card("Net P/L", rz(t["P/L (ZAR)"].sum()), f"{len(t)} trades", "good" if t["P/L (ZAR)"].sum() >= 0 else "bad")
    with c[1]:
        card("Win rate", f"{(t.Outcome == 'Win').mean():.0%}")
    with c[2]:
        card("Best trade", rz(t["P/L (ZAR)"].max()), "", "good")
    with c[3]:
        card("Worst trade", rz(t["P/L (ZAR)"].min()), "", "bad")
    fig = px.scatter(t, x="Date", y="P/L (ZAR)", color="Outcome", size="Capital (ZAR)", color_discrete_map={"Win": GREEN, "Loss": RED, "Flat": MUTE},
                     hover_data=["Week", "Trade"], title=f"{pick}: every trade")
    show(fig, 340)


def page_projection(f, d, k, summ, raw, sheets):
    st.markdown("Simulates the next trading month by **re-sampling your own real trading days** (so bad days stay bad). "
                "It's an illustration of ranges, not a forecast.")
    bal = summ[summ.Metric == "balance"].sort_values("Date") if len(summ) else pd.DataFrame()
    preset = st.radio("Scale", ["My current scale", "Workbook model (R300k)"], horizontal=True)
    cap0, alloc0 = (float(bal.Value.iloc[-1]), 2000.0) if (preset == "My current scale" and len(bal)) else (300000.0, 30000.0)
    c = st.columns(4)
    cap = c[0].number_input("Trading capital (ZAR)", 1000.0, 5_000_000.0, round(cap0, -2), 1000.0)
    alloc = c[1].number_input("Allocation per trade (ZAR)", 100.0, 500_000.0, alloc0, 100.0)
    tpd = c[2].slider("Trades per day", 1, 40, 10)
    days = c[3].slider("Trading days", 5, 60, 20)
    c = st.columns(4)
    haircut = c[0].slider("Stress: shave winners by", 0, 60, 0, format="%d%%") / 100
    tax = c[1].slider("Tax rate", 0, 45, 30, format="%d%%") / 100
    comm = c[2].number_input("Commission on gross profit (%)", 0.0, 5.0, 0.01, 0.01) / 100
    sims = c[3].select_slider("Simulations", [500, 1000, 2000, 5000], 2000)
    rets = d["Return %"].dropna().values
    if len(rets) < 5:
        st.warning("Need at least 5 trading days in the current filter to simulate.")
        return
    rng = np.random.default_rng(7)
    draw = rng.choice(rets, size=(sims, days))
    draw = np.where(draw > 0, draw * (1 - haircut), draw)
    exposure = min(alloc * tpd, cap)
    paths = np.cumsum(draw * exposure, axis=1)
    paths = np.hstack([np.zeros((sims, 1)), paths])
    final = paths[:, -1]
    net_after = np.where(final > 0, final * (1 - tax - comm), final)
    q = np.percentile(paths, [5, 25, 50, 75, 95], axis=0)
    x = np.arange(days + 1)

    m = st.columns(5)
    with m[0]:
        card("Median outcome", rz(np.median(final)), f"{np.median(final) / cap:.1%} of capital", "good" if np.median(final) >= 0 else "bad")
    with m[1]:
        card("Bad case (5th pct)", rz(q[0, -1]), "1 in 20 runs finish worse", "bad")
    with m[2]:
        card("Good case (95th pct)", rz(q[4, -1]), "1 in 20 runs finish better", "good")
    with m[3]:
        card("Chance of a losing period", f"{(final < 0).mean():.0%}", "Of simulated runs", "accent")
    with m[4]:
        card("Median after tax & fees", rz(np.median(net_after)), f"Tax {tax:.0%} · commission {comm:.2%}")

    a, b = st.columns([1.6, 1])
    with a:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=x, y=q[4], line=dict(width=0), showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=x, y=q[0], fill="tonexty", fillcolor="rgba(87,213,139,.10)", line=dict(width=0), name="5–95%"))
        fig.add_trace(go.Scatter(x=x, y=q[3], line=dict(width=0), showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=x, y=q[1], fill="tonexty", fillcolor="rgba(87,213,139,.22)", line=dict(width=0), name="25–75%"))
        fig.add_trace(go.Scatter(x=x, y=q[2], line=dict(color=GREEN, width=3), name="Median"))
        fig.update_layout(title="Range of cumulative profit across simulations", xaxis_title="Trading day", yaxis_title="Profit (ZAR)")
        show(fig, 420)
    with b:
        fig = px.histogram(x=final, nbins=40, title="Where runs finish", color_discrete_sequence=[GREEN])
        fig.add_vline(x=0, line_dash="dot", line_color=GOLD)
        fig.update_layout(xaxis_title="Final profit (ZAR)", yaxis_title="Runs", bargap=.05)
        show(fig, 420)
    st.caption(f"Built from {len(rets)} real trading days. Assumes that edge persists and that bigger size fills at the same prices; "
               "in smaller coins that often does not hold. Descriptive analytics only, not financial advice.")

    model = [s for s in sheets if "profitab" in s.lower()]
    if model:
        with st.expander("Your original profitability model sheet"):
            table(raw[model[0]].dropna(how="all").dropna(axis=1, how="all").astype(str).replace("nan", ""))


def page_day(f, d, k):
    dates = list(d.Date.sort_values(ascending=False))
    sel = st.selectbox("Trading day", dates, format_func=lambda x: f"{x:%A %d %b %Y}")
    t = f[f.Date == sel].sort_values("Trade")
    row = d[d.Date == sel].iloc[0]
    rank = int((d.Net > row.Net).sum()) + 1
    c = st.columns(5)
    with c[0]:
        card("Net P/L", rz(row.Net), f"Ranked {rank} of {len(d)} sessions", "good" if row.Net >= 0 else "bad")
    with c[1]:
        card("Trades", f"{int(row.Trades)}", f"{int(row.Wins)} wins · {int(row.Losses)} losses")
    with c[2]:
        card("Win rate", f"{row['Win rate']:.0%}", f"Period average {k['winrate']:.0%}", "accent")
    with c[3]:
        card("Return on deployed", f"{row['Return %']:.2%}", f"Capital {rz(row.Capital)}")
    with c[4]:
        card("Vs typical day", rz(row.Net - d.Net.median()), f"Median session {rz(d.Net.median())}", "good" if row.Net >= d.Net.median() else "bad")
    a, b = st.columns([1.7, 1])
    with a:
        w = go.Figure(go.Waterfall(x=t.Ticker, y=t["P/L (ZAR)"], measure=["relative"] * len(t),
                                   increasing=dict(marker=dict(color=GREEN)), decreasing=dict(marker=dict(color=RED)),
                                   connector=dict(line=dict(color="#2b3b33")),
                                   hovertemplate="%{x}<br>R %{y:,.0f}<extra></extra>"))
        w.update_layout(title="How the day built up, trade by trade", showlegend=False)
        show(w, 420)
    with b:
        by = t.sort_values("P/L (ZAR)")
        fig = go.Figure(go.Bar(x=by["P/L (ZAR)"], y=by.Ticker, orientation="h", marker_color=[GREEN if v >= 0 else RED for v in by["P/L (ZAR)"]]))
        fig.update_layout(title="Ranked")
        show(fig, 420)
    disp = t[["Trade", "Coin", "Capital (ZAR)", "Stop Loss %", "Profit Target %", "Actual Profit (ZAR)", "Actual Loss (ZAR)", "P/L (ZAR)", "Outcome"]].copy()
    for c_ in ["Stop Loss %", "Profit Target %"]:
        disp[c_] = disp[c_] * 100
    table(disp, column_config={c_: st.column_config.NumberColumn(format="%.2f%%") for c_ in ["Stop Loss %", "Profit Target %"]})


def page_goals(f, d, k):
    st.markdown("Set a daily target and see how often you hit it, and by how much.")
    c = st.columns(3)
    mode = c[0].radio("Target type", ["% of capital deployed", "Fixed ZAR per day"], horizontal=True)
    if mode == "% of capital deployed":
        pct = c[1].number_input("Daily target (%)", 0.1, 20.0, 2.0, 0.1) / 100
        d = d.assign(Target=d.Capital * pct)
    else:
        fixed = c[1].number_input("Daily target (ZAR)", 50.0, 100000.0, 750.0, 50.0)
        d = d.assign(Target=fixed)
    d = d.assign(Hit=d.Net >= d.Target, Attain=d.Net / d.Target)
    run = best = 0
    for h in d.Hit:
        run = run + 1 if h else 0
        best = max(best, run)
    c = st.columns(4)
    with c[0]:
        card("Days target hit", f"{int(d.Hit.sum())}/{len(d)}", f"{d.Hit.mean():.0%} of sessions", "accent")
    with c[1]:
        card("Average attainment", f"{d.Attain.mean():.0%}", "Net ÷ target, per day")
    with c[2]:
        card("Best run of hits", f"{best}", "Consecutive days")
    with c[3]:
        card("Total vs target", f"{d.Net.sum() / d.Target.sum():.0%}", f"{rz(d.Net.sum())} of {rz(d.Target.sum())}", "good" if d.Net.sum() >= d.Target.sum() else "bad")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=d.Date, y=d.Net, name="Net P/L", marker_color=[GREEN if h else (RED if n < 0 else GOLD) for h, n in zip(d.Hit, d.Net)],
                         hovertemplate="%{x|%a %d %b}<br>R %{y:,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=d.Date, y=d.Target, name="Target", mode="lines", line=dict(color="#e6efe9", dash="dot", width=2)))
    fig.update_layout(title="Daily result vs target (green = hit, gold = positive but short, red = loss)")
    show(fig, 400)
    cum = go.Figure()
    cum.add_trace(go.Scatter(x=d.Date, y=d.Net.cumsum(), name="Actual", line=dict(color=GREEN, width=3)))
    cum.add_trace(go.Scatter(x=d.Date, y=d.Target.cumsum(), name="Target pace", line=dict(color=GOLD, width=2, dash="dash")))
    cum.update_layout(title="Cumulative: actual vs target pace")
    show(cum, 360)


def excel_bytes(f, d):
    wk = f.groupby("Week", sort=False).agg(Trades=("Trade", "count"), Net=("P/L (ZAR)", "sum"), Capital=("Capital (ZAR)", "sum")).reset_index()
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        f.drop(columns=["WeekStart"]).to_excel(xw, sheet_name="Trades", index=False)
        d.to_excel(xw, sheet_name="Daily", index=False)
        wk.to_excel(xw, sheet_name="Weekly", index=False)
    return bio.getvalue()


def page_ledger(f, d=None):
    q = st.text_input("Search coin", placeholder="e.g. ETH, ONDO, ARB")
    view = f[f.Coin.str.contains(q, case=False, na=False)] if q else f
    disp = view.copy()
    for c_ in ["Stop Loss %", "Profit Target %", "Actual Profit %", "Actual Loss %", "Return %"]:
        disp[c_] = disp[c_] * 100
    cols = ["Week", "Date", "Day", "Trade", "Coin", "Capital (ZAR)", "Stop Loss %", "Profit Target %", "Actual Profit %", "Actual Profit (ZAR)",
            "Actual Loss %", "Actual Loss (ZAR)", "P/L (ZAR)", "Return %", "Outcome"]
    pc = {c_: st.column_config.NumberColumn(format="%.2f%%") for c_ in ["Stop Loss %", "Profit Target %", "Actual Profit %", "Actual Loss %", "Return %"]}
    pc["Date"] = st.column_config.DateColumn(format="DD MMM YYYY")
    table(disp[cols], column_config=pc)
    c = st.columns(2)
    c[0].download_button("Download filtered trades (CSV)", view.drop(columns=["WeekStart"]).to_csv(index=False).encode(), "filtered_trades.csv", "text/csv")
    if d is not None:
        c[1].download_button("Download clean workbook (Excel)", excel_bytes(f, d), "trade_intelligence_export.xlsx",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_quality(res, trades, summ):
    rec = reconcile(trades, summ)
    inv, iss = res["inventory"], res["issues"]
    ok = (rec.Status == "Match").sum() if len(rec) else 0
    checked = (rec.Status != "No summary").sum() if len(rec) else 0
    c = st.columns(4)
    with c[0]:
        card("Trade rows parsed", f"{len(trades):,}", f"{trades.Sheet.nunique()} weekly sheets")
    with c[1]:
        card("Days reconciled", f"{ok}/{checked}", "Trade rows agree with your summary table", "good" if ok == checked else "accent")
    with c[2]:
        card("Layout fixes applied", str(len(iss)), "Auto-corrected on read", "accent")
    with c[3]:
        card("Unique coins", str(trades.Ticker.nunique()))

    if len(iss):
        section("Auto-fixes")
        table(iss)
    if len(rec):
        section("Reconciliation: trade rows vs your daily summary")
        r2 = rec.copy()
        r2["Date"] = r2["Date"].dt.strftime("%a %d %b")
        r2 = r2.merge(trades[["Sheet", "Week"]].drop_duplicates(), on="Sheet", how="left")
        table(r2[["Week", "Date", "Rows: trades", "Summary: trades", "Rows: net (ZAR)", "Summary: net (ZAR)", "Status", "Detail"]])
    bal = summ[summ.Metric == "balance"].sort_values("Date") if len(summ) else pd.DataFrame()
    if len(bal) > 2:
        section("Balance bridge")
        net = trades.groupby("Date")["P/L (ZAR)"].sum()
        b = bal[["Date", "Value"]].rename(columns={"Value": "Balance"}).set_index("Date")
        b["Change"] = b.Balance.diff()
        b["Net (same day)"] = net.reindex(b.index)
        b["Net (prior day)"] = net.reindex(b.index - pd.Timedelta(days=1)).values
        b["Reads as"] = np.where(np.isclose(b.Change, b["Net (same day)"], atol=1.5), "End-of-day balance",
                                 np.where(np.isclose(b.Change, b["Net (prior day)"], atol=1.5), "Start-of-day balance", "No match"))
        b.loc[b.Change.isna(), "Reads as"] = ""
        table(b.reset_index().assign(Date=lambda x: x.Date.dt.strftime("%a %d %b")))
        st.caption("A balance that mixes start-of-day and end-of-day readings makes the account curve look off by a day. "
                   "Worth keeping one convention.")
    section("Gaps")
    mons = sorted(trades.WeekStart.dropna().unique())
    if len(mons) > 1:
        full = pd.date_range(mons[0], mons[-1], freq="7D")
        missing = [m for m in full if m not in set(mons)]
        st.write("No trade sheet for the week of: " + (", ".join(f"{m:%d %b}" for m in missing) if missing else "none"))
    section("Sheet inventory")
    table(inv)
    with st.expander("How the parser reads your workbook"):
        st.markdown("- A weekday label (`Monday`, `Tuesday 11 aug 2026`) in column A starts a day.\n"
                    "- A row starting with `Trade` or `Trades` is the header; numbered rows below it are trades.\n"
                    "- Columns are matched by meaning, so `Capital allocated(ZAR)`, `Capital Allocated` and `Stop Loss %` all work.\n"
                    "- The `Metric | day | day` table at the top of a sheet is read for reconciliation and balances.\n"
                    "- Sheets with `profitability` in the name are treated as the model, not trades.")


# ───────────────────────────── APP ─────────────────────────────
def main():
    st.set_page_config(page_title="Trade Intelligence", page_icon="◈", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### ◈ Trade Intelligence")
        st.caption("Workbook-powered analytics")
        up = st.file_uploader("Replace workbook", type=["xlsx"], help="Same layout as trading_data.xlsx")
        st.divider()
        page = st.radio("Workspace", ["Command Center", "Performance", "Day Explorer", "Goals", "Risk Lab", "Coin Intelligence",
                                      "Projection Lab", "Trade Ledger", "Data Quality"], label_visibility="collapsed")
        st.divider()

    if up is not None:
        data = up.getvalue()
    else:
        wb = APP_DIR / "trading_data.xlsx"
        if not wb.exists():
            found = sorted(APP_DIR.glob("*.xlsx")) + sorted((APP_DIR / "data").glob("*.xlsx")) + sorted(Path.cwd().glob("*.xlsx"))
            wb = found[0] if found else None
        if wb is None:
            st.error("No workbook found. Upload trading_data.xlsx next to streamlit_app.py or use the sidebar uploader.")
            st.stop()
        data = wb.read_bytes()
    res = load(data)
    trades, summ = res["trades"], res["summary"]
    if trades.empty:
        st.error("No trade rows could be parsed from the workbook.")
        st.stop()

    week_order = list(dict.fromkeys(trades.sort_values("Date").Week))
    with st.sidebar:
        st.markdown("**Filters**")
        wk = st.multiselect("Weeks", week_order, default=week_order)
        days = [x for x in WEEKDAYS if x in set(trades.Day)]
        dy = st.multiselect("Days", days, default=days)
        oc = st.multiselect("Outcome", ["Win", "Loss", "Flat"], default=["Win", "Loss", "Flat"])
        st.divider()
        st.caption("Built with Streamlit · pandas · Plotly")

    f = trades[trades.Week.isin(wk) & trades.Day.isin(dy) & trades.Outcome.isin(oc)].copy()
    if f.empty:
        st.warning("No trades match the current filters.")
        st.stop()
    d = daily_table(f)
    k = kpis(f, d)
    span = f"{f.Date.min():%d %b} – {f.Date.max():%d %b %Y}"
    st.markdown(
        f'<div class="hero"><div class="eyebrow">Trading performance system</div><h1>Trade Intelligence</h1>'
        f'<p>Every figure is rebuilt from your trade rows, cross-checked against your own summary tables, and sliced for edge, risk and consistency.</p>'
        f'<div class="chips"><span class="chip"><b>{k["n"]:,}</b> trades</span><span class="chip"><b>{k["days"]}</b> sessions</span>'
        f'<span class="chip"><b>{f.Ticker.nunique()}</b> coins</span><span class="chip">{span}</span></div></div>', unsafe_allow_html=True)

    if page == "Command Center":
        page_command(f, d, k, summ, week_order)
    elif page == "Performance":
        page_performance(f, d, k, summ, week_order)
    elif page == "Day Explorer":
        page_day(f, d, k)
    elif page == "Goals":
        page_goals(f, d, k)
    elif page == "Risk Lab":
        page_risk(f, d, k, summ, week_order)
    elif page == "Coin Intelligence":
        page_coins(f, d, k, summ, week_order)
    elif page == "Projection Lab":
        page_projection(f, d, k, summ, res["raw"], res["sheets"])
    elif page == "Trade Ledger":
        page_ledger(f, d)
    else:
        page_quality(res, trades, summ)
    st.divider()
    st.caption("Source: your Excel trading workbook · Analytics are descriptive and do not constitute financial advice.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # show the real error in the app instead of a blank page
        st.error("The dashboard hit an error. Copy the details below if you need help.")
        st.exception(exc)
