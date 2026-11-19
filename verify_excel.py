"""既存Excelの各月シートについて、calc.py の結果とExcelの保存済み計算結果を突き合わせる。"""
import glob, sys, warnings
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L
from calc import SiteParams, calc_bill

warnings.filterwarnings("ignore")
LABELS = {"基本料金":"base","単価":"energy","燃料費＠":"fuel","燃料費":"fuel","再エネ＠":"renewable","再エネ":"renewable",
          "産蓄熱":"fixed","単年割":"annual","市場":"market","市場価格":"market",
          "税込み":"incl","税抜き":"excl"}

def num(v):
    return v if isinstance(v,(int,float)) else None

def verify_sheet(wf, wv):
    # 月合計列(5行目)
    tc = next((c for c in range(3, wf.max_column+1) if wf.cell(5,c).value == "月合計"), None)
    if tc is None: return "SKIP:月合計列なし"
    # 区分表のヘッダ行(月合計按分)
    hdr = tc2 = None
    for r in range(8, wf.max_row+1):
        for c in range(20, wf.max_column+1):
            if wf.cell(r,c).value == "月合計按分":
                hdr = r
    if hdr is None: return "SKIP:区分表なし"
    price_col = next(c for c in range(20, wf.max_column+1) if wf.cell(hdr,c).value == "単価")
    sum_col  = next(c for c in range(20, wf.max_column+1) if wf.cell(hdr,c).value in ("月合計","先月合計"))
    kwh = num(wv.cell(6,tc).value)
    bands = [num(wv.cell(hdr+i,sum_col).value) for i in (1,2,3,4)]
    prices = [num(wv.cell(hdr+i,price_col).value) for i in (1,2,3,4)]
    if None in bands or None in prices or kwh is None: return "SKIP:数値欠落"
    # 料金ブロック
    P = {}; res = {}
    for r in range(hdr+5, wf.max_row+1):
        for c in range(20, wf.max_column+1):
            lab = wf.cell(r,c).value
            if isinstance(lab,str) and lab.strip() in LABELS:
                k = LABELS[lab.strip()]
                P[k] = [num(wv.cell(r,c+i).value) for i in (1,2,3)]
                res[k] = (num(wv.cell(r,tc).value), wf.cell(r,tc).value)
    if "base" not in P: return "SKIP:料金ブロックなし"
    b = P["base"]
    first = lambda k: next((x for x in P.get(k, []) if x is not None), 0)
    sp = SiteParams(contract_kw=b[2], base_unit=b[0], power_factor=b[1],
                    unit_price=first("energy"), fuel_adj=first("fuel"),
                    renewable=first("renewable"),
                    annual_discount=first("annual"),
                    fixed_credit=(res.get("fixed",(0,))[0] or 0),
                    band_prices=tuple(prices))
    # 市場価格の行が「=月合計セル(kWh)」になっている場合は kWh×1円 として再現
    mf = res.get("market", (None, None))[1]
    if isinstance(mf, str) and mf.replace("$", "") == f"={L(tc)}6":
        sp = SiteParams(**{**sp.__dict__, "market_flat": 1.0})
    bill = calc_bill(kwh, bands, sp)
    mine = {"base":bill.base,"energy":bill.energy,"fuel":bill.fuel,"renewable":bill.renewable,
            "annual":bill.annual_discount,"market":bill.market,"incl":bill.tax_included,"excl":bill.tax_excluded}
    out = []
    for k,(cv,formula) in res.items():
        if k in mine and cv is not None and abs(cv-mine[k]) > 0.5:
            out.append(f"{k}: Excel={cv:,.1f}({formula}) / 計算={mine[k]:,.1f}")
    return out

files = sorted(glob.glob("/mnt/user-data/uploads/*.xlsx"))
for f in files:
    wbf = load_workbook(f); wbv = load_workbook(f, data_only=True)
    print("="*70); print(f.split("/")[-1])
    for ws in wbf.worksheets:
        if ws.title == "年間合算": continue
        r = verify_sheet(ws, wbv[ws.title])
        if isinstance(r,str):  print(f"  [{ws.title}] {r}")
        elif not r:            print(f"  [{ws.title}] OK(全項目一致)")
        else:
            print(f"  [{ws.title}] 差あり")
            for line in r: print("      ", line)
