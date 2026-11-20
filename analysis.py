"""集計ロジック(画面・保存先に依存しない)。pandasだけで動くのでテストしやすい。"""
from __future__ import annotations

import pandas as pd

from calc import SiteParams, billing_month, billing_period, calc_bill, excel_round

BAND_COLS = ["morning", "noon", "evening", "night"]
BAND_JA = ["朝", "昼", "晩", "夜"]
DAILY_NUM = ["total_kwh", "new_bldg_kwh"] + BAND_COLS
DAILY_COLS = ["site", "date"] + DAILY_NUM + ["source"]
PARAM_NUM = ["contract_kw", "base_unit", "power_factor", "unit_price", "fuel_adj", "renewable",
             "annual_discount", "fixed_credit", "price_morning", "price_noon", "price_evening", "price_night"]
PARAM_COLS = ["site", "month"] + PARAM_NUM
SITE_COLS = ["site", "start_day", "has_new_bldg", "source_name"]


def _num(s: pd.Series) -> pd.Series:
    s = s.astype("string").str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce").astype("float64")


def _ensure(df: pd.DataFrame, cols) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df


def normalize_daily(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure(df, DAILY_COLS)
    df["site"] = df["site"].astype("string").str.strip()
    df["date"] = pd.to_datetime(df["date"].astype("string").str.strip(), format="mixed", errors="coerce").dt.normalize()
    for c in DAILY_NUM:
        df[c] = _num(df[c])
    df["source"] = df["source"].astype("string").fillna("actual")
    df = df.dropna(subset=["site", "date"])
    df = df.drop_duplicates(["site", "date"], keep="last").sort_values(["site", "date"]).reset_index(drop=True)
    return df[DAILY_COLS]


def normalize_params(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure(df, PARAM_COLS)
    df["site"] = df["site"].astype("string").str.strip()
    mk = pd.to_datetime(df["month"].astype("string").str.strip(), format="mixed", errors="coerce")
    df["month_key"] = mk.dt.year * 100 + mk.dt.month
    for c in PARAM_NUM:
        df[c] = _num(df[c])
    df = df.dropna(subset=["site", "month_key"])
    df["month_key"] = df["month_key"].astype(int)
    return df.sort_values(["site", "month_key"]).reset_index(drop=True)


def normalize_sites(df: pd.DataFrame) -> pd.DataFrame:
    df = _ensure(df, SITE_COLS)
    df["site"] = df["site"].astype("string").str.strip()
    df = df.dropna(subset=["site"]).copy()
    df["start_day"] = _num(df["start_day"]).fillna(27).astype(int)
    flag = df["has_new_bldg"].astype("string").str.strip().str.lower()
    df["has_new_bldg"] = flag.isin(["true", "1", "yes", "y", "はい", "○"])
    df["source_name"] = df["source_name"].astype("string").str.strip().fillna(df["site"])
    return df[SITE_COLS].reset_index(drop=True)


def params_for(params: pd.DataFrame, site: str, y: int, m: int) -> SiteParams | None:
    """その月に適用する料金設定。月ごとの行がなければ、直前の月の設定を引き継ぐ。"""
    p = params[(params["site"] == site) & (params["month_key"] <= y * 100 + m)]
    if p.empty:
        return None
    r = p.iloc[-1]

    def g(k, default=0.0):
        return float(r[k]) if pd.notna(r[k]) else default

    return SiteParams(
        contract_kw=g("contract_kw"), base_unit=g("base_unit", 2530), power_factor=g("power_factor", 0.85),
        unit_price=g("unit_price"), fuel_adj=g("fuel_adj"), renewable=g("renewable"),
        annual_discount=g("annual_discount"), fixed_credit=g("fixed_credit"),
        band_prices=(g("price_morning"), g("price_noon"), g("price_evening"), g("price_night")),
    )


def prev_year(d):
    try:
        return d.replace(year=d.year - 1)
    except ValueError:            # 2月29日
        return d.replace(year=d.year - 1, day=28)


def months_available(daily: pd.DataFrame, site: str, start_day: int):
    ds = daily[(daily["site"] == site) & (daily["source"] != "prev_year")]
    return sorted({billing_month(d, start_day) for d in ds["date"].dt.date}, reverse=True)


def month_label(y: int, m: int, start_day: int) -> str:
    a, b = billing_period(y, m, start_day)
    return f"{y}年{m}月分({a.month}/{a.day}〜{b.month}/{b.day})"


def period_frame(daily: pd.DataFrame, site: str, y: int, m: int, start_day: int) -> pd.DataFrame:
    a, b = billing_period(y, m, start_day)
    idx = pd.date_range(a, b)
    s = daily[daily["site"] == site].set_index("date")
    f = pd.DataFrame(index=idx)
    for c in DAILY_NUM:
        f[c] = s[c].reindex(idx)
    f["prev_total"] = s["total_kwh"].reindex([prev_year(d) for d in idx]).to_numpy()
    return f


def month_summary(daily, params, site: str, start_day: int, y: int, m: int) -> dict:
    a, b = billing_period(y, m, start_day)
    f = period_frame(daily, site, y, m, start_day)
    filled = f["total_kwh"].notna()
    days_total, days_filled = len(f), int(filled.sum())
    kwh = float(f["total_kwh"].sum())
    prev_days = int((filled & f["prev_total"].notna()).sum())
    prev_kwh = float(f.loc[filled, "prev_total"].sum()) if prev_days else None
    band_days = int(f.loc[filled, BAND_COLS].notna().all(axis=1).sum())
    bands = tuple(float(f[c].sum()) for c in BAND_COLS)
    p = params_for(params, site, y, m)

    notes, bill = [], None
    if days_filled < days_total:
        notes.append(f"途中経過です({days_filled}/{days_total}日分)。料金は参考値で、基本料金は1か月分で計算しています。")
    if p is None:
        notes.append("この月の料金設定(paramsタブ)がありません。")
    elif days_filled == 0 or band_days == 0:
        notes.append("時間帯別(朝・昼・晩・夜)が入力されていないため、料金を試算できません。")
    else:
        if band_days < days_filled:
            notes.append(f"時間帯別が未入力の日があります({days_filled - band_days}日)。割合は入力済みの日だけで計算します。")
        bill = calc_bill(kwh, bands, p)
    if prev_kwh is not None and prev_days < days_filled:
        notes.append(f"前年の値が一部ありません(前年比は{prev_days}日分で比較)。")
    return dict(site=site, year=y, month=m, start=a, end=b, days_total=days_total, days_filled=days_filled,
                kwh=kwh, prev_kwh=prev_kwh, prev_days=prev_days, band_days=band_days, bands=bands,
                params=p, bill=bill, notes=notes, frame=f)


def monthly_table(daily, params, sites, site_names, d_from, d_to) -> pd.DataFrame:
    rows = []
    for _, cfg in sites[sites["site"].isin(site_names)].iterrows():
        sd = int(cfg["start_day"])
        for (y, m) in sorted(months_available(daily, cfg["site"], sd)):
            a, b = billing_period(y, m, sd)
            if b < d_from or a > d_to:
                continue
            s = month_summary(daily, params, cfg["site"], sd, y, m)
            bill = s["bill"]
            r = {"拠点": cfg["site"], "検針月": f"{y}-{m:02d}", "期間開始": a.isoformat(), "期間終了": b.isoformat(),
                 "入力日数": s["days_filled"], "日数": s["days_total"], "kWh": round(s["kwh"]),
                 "前年同期kWh": None if s["prev_kwh"] is None else round(s["prev_kwh"])}
            names = [("基本料金", "base"), ("電力量料金", "energy"), ("燃料費調整", "fuel"), ("再エネ", "renewable"),
                     ("産蓄熱等", "fixed_credit"), ("単年割", "annual_discount"), ("市場価格", "market"),
                     ("税込み", "tax_included"), ("税抜き", "tax_excluded")]
            for ja, attr in names:
                r[ja] = None if bill is None else int(excel_round(getattr(bill, attr)))
            r["備考"] = " / ".join(s["notes"])
            rows.append(r)
    return pd.DataFrame(rows)


# ---------- 収集プログラムの出力(Excel)の取り込み ----------
COLLECTOR_ROWS = {"morningKwh": "morning", "noonKwh": "noon", "eveningKwh": "evening", "nightKwh": "night"}


def parse_collector_xlsx(file, year: int, month: int, sites: pd.DataFrame):
    """収集プログラムのExcel(拠点ごとのシート、1行目=日、A列=morningKwh等)を読む。
    ファイルに年月は入っていないので、year/month は呼び出し側で指定する。"""
    from datetime import date
    from openpyxl import load_workbook
    wb = load_workbook(file, data_only=True)
    src = {str(r.source_name): r.site for r in sites.itertuples()}
    rows, warns, seen = [], [], set()
    for ws in wb.worksheets:
        site = src.get(ws.title.strip())
        if site is None:
            warns.append(f"シート「{ws.title}」に対応する拠点がありません(sitesタブのsource_nameを確認してください)。")
            continue
        seen.add(site)
        labels = {}
        for r in range(1, ws.max_row + 1):
            lab = ws.cell(r, 1).value
            if isinstance(lab, str) and lab.strip() in COLLECTOR_ROWS:
                labels[COLLECTOR_ROWS[lab.strip()]] = r
        if len(labels) < 4:
            warns.append(f"シート「{ws.title}」に朝・昼・晩・夜の行がそろっていません。")
            continue
        for c in range(2, ws.max_column + 1):
            try:
                d = date(year, month, int(ws.cell(1, c).value))
            except (TypeError, ValueError):
                continue
            vals = {k: ws.cell(r, c).value for k, r in labels.items()}
            if any(isinstance(v, bool) or not isinstance(v, (int, float)) for v in vals.values()):
                warns.append(f"{site} {d.month}/{d.day}: 値が空か数値でないため取り込みません。")
                continue
            rows.append({"site": site, "date": pd.Timestamp(d), **{k: float(v) for k, v in vals.items()}})
    for s in sites["site"]:
        if s not in seen:
            warns.append(f"{s}のシートがファイルにありません。")
    return pd.DataFrame(rows, columns=["site", "date"] + BAND_COLS), warns


def plan_import(daily: pd.DataFrame, new: pd.DataFrame, today, update_total: bool = False) -> pd.DataFrame:
    """取り込み内容と、既存データとの比較(新規/更新/同じ)を作る。
    全体kWh: 既存の行があればそのまま残す(update_total=Trueなら朝〜夜の合計で更新)。新規の日は朝〜夜の合計。"""
    ex = daily.set_index(["site", "date"]) if not daily.empty else None
    out = []
    for r in new.itertuples(index=False):
        d = r.date.date()
        if d > today:
            continue                                   # 未来の日付は取り込まない
        bands = [float(getattr(r, c)) for c in BAND_COLS]
        bsum = sum(bands)
        notes, new_bldg = [], float("nan")
        e = ex.loc[(r.site, r.date)] if ex is not None and (r.site, r.date) in ex.index else None
        if e is None:
            status, total = "新規", bsum
        else:
            old = [e[c] for c in BAND_COLS]
            same = all(pd.notna(o) and o == b for o, b in zip(old, bands))
            total = bsum if update_total else (float(e["total_kwh"]) if pd.notna(e["total_kwh"]) else bsum)
            new_bldg = e["new_bldg_kwh"]
            status = "同じ" if same and (not update_total or total == e["total_kwh"]) else "更新"
            if not same:
                diffs = [f"{n} {o:,.0f}→{b:,.0f}" if pd.notna(o) else f"{n} 未入力→{b:,.0f}"
                         for n, o, b in zip(BAND_JA, old, bands) if not (pd.notna(o) and o == b)]
                notes.append("既存と違う区分: " + "、".join(diffs))
        if total > 0 and abs(total - bsum) > total * 0.02:
            notes.append(f"全体({total:,.0f})と朝〜夜の合計({bsum:,.0f})の差が大きい")
        if d == today:
            notes.append("今日の分です。1日の途中の値かもしれません")
        out.append({"status": status, "site": r.site, "date": r.date, "total_kwh": total, "new_bldg_kwh": new_bldg,
                    **dict(zip(BAND_COLS, bands)), "note": " / ".join(notes)})
    cols = ["status", "site", "date", "total_kwh", "new_bldg_kwh"] + BAND_COLS + ["note"]
    return pd.DataFrame(out, columns=cols).sort_values(["site", "date"]).reset_index(drop=True)


def plan_records(plan: pd.DataFrame) -> list:
    """保存する行(「同じ」は除く)を、upsert用の辞書にする。"""
    recs = []
    for r in plan[plan["status"] != "同じ"].itertuples(index=False):
        recs.append({"site": r.site, "date": r.date.date().isoformat(), "total_kwh": r.total_kwh,
                     "new_bldg_kwh": None if pd.isna(r.new_bldg_kwh) else r.new_bldg_kwh,
                     **{c: getattr(r, c) for c in BAND_COLS}, "source": "actual"})
    return recs


def date_ranges(dates) -> list:
    ds = sorted(set(dates))
    out, i = [], 0
    while i < len(ds):
        j = i
        while j + 1 < len(ds) and (ds[j + 1] - ds[j]).days == 1:
            j += 1
        a, b = ds[i], ds[j]
        out.append(f"{a.month}/{a.day}" if i == j else f"{a.month}/{a.day}〜{b.month}/{b.day}")
        i = j + 1
    return out


def missing_days(daily: pd.DataFrame, plan: pd.DataFrame, site_names, year: int, month: int, today) -> dict:
    """指定した月の今日までで、(既存+今回の取込を合わせても)全体kWhが入っていない日。"""
    from datetime import date, timedelta
    first = date(year, month, 1)
    nxt = date(year + (month == 12), month % 12 + 1, 1)
    last = min(today, nxt - timedelta(days=1))
    have = {(r.site, r.date.date()) for r in daily.itertuples() if pd.notna(r.total_kwh)}
    have |= {(r.site, r.date.date()) for r in plan.itertuples()}
    res = {}
    for s in site_names:
        miss, d = [], first
        while d <= last:
            if (s, d) not in have:
                miss.append(d)
            d += timedelta(days=1)
        if miss:
            res[s] = date_ranges(miss)
    return res
