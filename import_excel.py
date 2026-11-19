"""既存の電気代Excel(HO1/HO2/城南島/大井埠頭)から、日次データと料金パラメータをCSVに取り込む。

出力: data/daily.csv   (site,date,total_kwh,new_bldg_kwh,morning,noon,evening,night,source)
      data/params.csv  (site,month,契約kWなど。月ごとの料金パラメータ)
実行: python import_excel.py <Excelのフォルダ> [--check]
"""
import re, sys, csv, glob, os, warnings
from datetime import date
from openpyxl import load_workbook
from calc import SiteParams, calc_bill, billing_period

warnings.filterwarnings("ignore")
SITES = {"HO1": ("HO1電気代", 27), "HO2": ("HO2電気代", 27), "城南島": ("城南島電気代", 1), "大井埠頭": ("大井埠頭電気代", 27)}
BAND_LABELS = ["朝時間", "昼時間", "晩時間", "夜時間"]
PARAM_LABELS = {"基本料金": "base", "単価": "energy", "燃料費＠": "fuel", "燃料費": "fuel", "再エネ＠": "renewable",
                "再エネ": "renewable", "産蓄熱": "fixed", "単年割": "annual"}


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def sheet_month(name):
    d = re.sub(r"\D", "", name)
    return (int(d[:4]), int(d[4:])) if len(d) >= 5 else None


def column_dates(wv, ym, tc, start_day):
    """データ列 C..(tc-1) の日付を返す {列番号: date}。"""
    y, m = ym
    first = None
    out, prev_day, cur = {}, None, None
    for c in range(3, tc):
        t = wv.cell(5, c).value
        mm = re.match(r"(\d+)", str(t or ""))
        if not mm:
            continue
        day = int(mm.group(1))
        if cur is None:
            cur = (m if start_day <= 1 else (m - 1 or 12), y if (start_day <= 1 or m > 1) else y - 1)
            cur = (cur[1], cur[0]) if False else cur
        if prev_day is not None and day < prev_day:       # 月が変わった(31→1)
            cur = (m, y)
        prev_day = day
        mon, yr = cur
        try:
            out[c] = date(yr, mon, day)
        except ValueError:
            pass
    return out


def find_rows(wv, ym):
    """B列の見出しから 当年/前年 の全体・新棟、および区分(朝昼晩夜)の行番号を探す。"""
    y = ym[0]
    rows = {}
    year_rows = []
    for r in range(6, 25):
        b = wv.cell(r, 2).value
        if not isinstance(b, str):
            continue
        b = b.strip()
        if b in BAND_LABELS:
            rows[b] = r
        elif "全体" in b or "新棟" in b:
            yy = int(re.match(r"(\d{4})", b).group(1))
            kind = "new" if "新棟" in b else "total"
            rows[("cur" if yy == y else "prev", kind)] = r
        elif "年度" in b:
            year_rows.append(r)
    if year_rows:
        rows[("cur", "total")] = year_rows[0]
        if len(year_rows) > 1 and not wv.cell(year_rows[1], 2).value.strip().startswith("差"):
            rows[("prev", "total")] = year_rows[1]
    return rows


def read_params(wf, wv, tc):
    hdr = None
    for r in range(8, wf.max_row + 1):
        for c in range(20, wf.max_column + 1):
            if wf.cell(r, c).value == "月合計按分":
                hdr = r
    if hdr is None:
        return None
    pc = next(c for c in range(20, wf.max_column + 1) if wf.cell(hdr, c).value == "単価")
    prices = [num(wv.cell(hdr + i, pc).value) for i in (1, 2, 3, 4)]
    P, fixed = {}, 0
    for r in range(hdr + 5, wf.max_row + 1):
        for c in range(20, wf.max_column + 1):
            lab = wf.cell(r, c).value
            if isinstance(lab, str) and lab.strip() in PARAM_LABELS:
                k = PARAM_LABELS[lab.strip()]
                if k == "fixed":
                    fixed = num(wv.cell(r, tc).value) or 0
                else:
                    P[k] = [num(wv.cell(r, c + i).value) for i in (1, 2, 3)]
    if "base" not in P or None in prices:
        return None
    first = lambda k: next((x for x in P.get(k, []) if x is not None), 0)
    b = P["base"]
    return dict(contract_kw=b[2], base_unit=b[0], power_factor=b[1], unit_price=first("energy"),
                fuel_adj=first("fuel"), renewable=first("renewable"), annual_discount=first("annual"),
                fixed_credit=fixed, price_morning=prices[0], price_noon=prices[1],
                price_evening=prices[2], price_night=prices[3])


def main(folder, check=False):
    daily, params, excel_bills = {}, [], {}
    for site, (prefix, start_day) in SITES.items():
        path = next(iter(glob.glob(os.path.join(folder, prefix + "*.xlsx"))), None)
        if not path:
            print("見つかりません:", prefix); continue
        wbf, wbv = load_workbook(path), load_workbook(path, data_only=True)
        for ws in wbf.worksheets:
            ym = sheet_month(ws.title)
            wv = wbv[ws.title]
            tc = next((c for c in range(3, ws.max_column + 1) if ws.cell(5, c).value == "月合計"), None)
            if ym is None or tc is None:
                continue
            rows = find_rows(wv, ym)
            if ("cur", "total") not in rows:
                print(f"  スキップ {site}[{ws.title}]: 全体行なし"); continue
            dates = column_dates(wv, ym, tc, start_day)
            for kind in ("cur", "prev"):
                for c, d in dates.items():
                    dd = d if kind == "cur" else (d.replace(year=d.year - 1) if not (d.month == 2 and d.day == 29) else None)
                    if dd is None:
                        continue
                    tot = num(wv.cell(rows[(kind, "total")], c).value) if (kind, "total") in rows else None
                    if not tot:
                        continue
                    rec = dict(site=site, date=dd.isoformat(), total_kwh=tot,
                               new_bldg_kwh=num(wv.cell(rows[(kind, "new")], c).value) if (kind, "new") in rows else None,
                               source="actual" if kind == "cur" else "prev_year")
                    if kind == "cur":
                        for key, lab in zip(("morning", "noon", "evening", "night"), BAND_LABELS):
                            rec[key] = num(wv.cell(rows[lab], c).value) if lab in rows else None
                    old = daily.get((site, rec["date"]))
                    if old is None or (old["source"] == "prev_year" and rec["source"] == "actual"):
                        daily[(site, rec["date"])] = rec
            p = read_params(ws, wv, tc)
            if p:
                params.append(dict(site=site, month=f"{ym[0]}-{ym[1]:02d}", **p))
                tot_cell = next((num(wv.cell(r, tc).value) for r in range(15, ws.max_row + 1)
                                 if any(ws.cell(r, c).value == "税込み" for c in range(20, ws.max_column + 1))), None)
                excel_bills[(site, f"{ym[0]}-{ym[1]:02d}")] = tot_cell
    os.makedirs("data", exist_ok=True)
    cols = ["site", "date", "total_kwh", "new_bldg_kwh", "morning", "noon", "evening", "night", "source"]
    with open("data/daily.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, cols); w.writeheader()
        for k in sorted(daily): w.writerow(daily[k])
    pcols = ["site", "month", "contract_kw", "base_unit", "power_factor", "unit_price", "fuel_adj", "renewable",
             "annual_discount", "fixed_credit", "price_morning", "price_noon", "price_evening", "price_night"]
    with open("data/params.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, pcols); w.writeheader(); w.writerows(params)
    n_act = sum(1 for r in daily.values() if r["source"] == "actual")
    print(f"daily.csv: {len(daily)}行(実績{n_act} / 前年値{len(daily)-n_act}), params.csv: {len(params)}行")
    if check:
        roundtrip(daily, params, excel_bills)


def roundtrip(daily, params, excel_bills):
    """取り込んだCSV相当のデータから料金を再計算し、Excelの税込額と比べる。"""
    print("\n--- 往復チェック(取込データ→calc.py→Excelの税込額) ---")
    for p in params:
        site, month = p["site"], p["month"]
        y, m = map(int, month.split("-"))
        a, b = billing_period(y, m, SITES[site][1])
        rows = [r for (s, d), r in daily.items() if s == site and a.isoformat() <= d <= b.isoformat() and r["source"] == "actual"]
        if not rows:
            continue
        kwh = sum(r["total_kwh"] for r in rows)
        bands = [sum((r[k] or 0) for r in rows) for k in ("morning", "noon", "evening", "night")]
        sp = SiteParams(contract_kw=p["contract_kw"], base_unit=p["base_unit"], power_factor=p["power_factor"],
                        unit_price=p["unit_price"], fuel_adj=p["fuel_adj"], renewable=p["renewable"],
                        annual_discount=p["annual_discount"], fixed_credit=p["fixed_credit"],
                        band_prices=(p["price_morning"], p["price_noon"], p["price_evening"], p["price_night"]))
        mine = calc_bill(kwh, bands, sp).tax_included
        ex = excel_bills.get((site, month))
        flag = "OK" if ex is not None and abs(mine - ex) <= 1 else "差あり"
        print(f"  {site:<5}{month} 日数{len(rows):>2}  Excel={ex if ex is None else format(ex, ',.0f')}  取込計算={mine:,.0f}  {flag}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else ".", "--check" in sys.argv)
