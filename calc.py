"""電気代計算ロジック(画面・保存先に依存しない純粋関数)。

既存Excel(HO1/HO2/HO3/城南島/大井埠頭)の月次シートの計算式を再現したもの。
  割合       = ROUND(区分kWh / 区分kWh合計, 3)  ※最後の区分は 1 - 他の合計
  按分kWh    = ROUND(割合 × 月合計kWh, 1)
  市場価格   = Σ(按分kWh × 区分単価)
  基本料金   = 基本単価 × 力率 × 契約kW
  電力量料金 = kWh × 単価、 燃料費/再エネ/単年割 も kWh × 円/kWh
  税込み     = 基本 + 単価 + 燃料費 + 再エネ + 産蓄熱 + 単年割 + 市場
  税抜き     = 税込み ÷ 1.1
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Sequence

BANDS = ("朝", "昼", "晩", "夜")


def excel_round(x: float, digits: int = 0) -> float:
    """ExcelのROUND(四捨五入=0から遠い方へ)。PythonのroundはBankers丸めなので使わない。"""
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP))


@dataclass(frozen=True)
class SiteParams:
    """拠点ごとの料金パラメータ(適用開始月ごとに持つ想定)。"""
    contract_kw: float                       # 契約kW
    base_unit: float = 2530                  # 基本料金の単価(円/kW)
    power_factor: float = 0.85               # 力率
    unit_price: float = 17.43                # 電力量料金単価(円/kWh)
    fuel_adj: float = -0.4                   # 燃料費調整(円/kWh)
    renewable: float = 4.18                  # 再エネ賦課金(円/kWh)
    annual_discount: float = 0.0             # 単年割(円/kWh、割引は負)
    fixed_credit: float = 0.0                # 産蓄熱など月額固定(円、割引は負)
    band_prices: tuple = (4.5, 6.16, 6.99, 3.69)  # 朝・昼・晩・夜の市場単価
    # 市場価格を「kWh × 固定円/kWh」で計算する場合に指定(通常はNone=区分別の市場単価で計算)。
    # HO3の2026年9〜11月のExcelは市場価格を `=AH6`(kWh×1)で足しているため、その再現用。
    market_flat: float | None = None
    tax_rate: float = 0.10


@dataclass
class Bill:
    kwh: float
    shares: tuple
    allocated: tuple
    market: float
    base: float
    energy: float
    fuel: float
    renewable: float
    fixed_credit: float
    annual_discount: float
    tax_included: float
    tax_excluded: float
    detail: dict = field(default_factory=dict)


def band_shares(band_kwh: Sequence[float]) -> tuple:
    """区分別kWhから割合(小数3桁)を出す。最後の区分は端数調整で 1 - 他の合計。"""
    total = sum(band_kwh)
    if total <= 0:
        return (0.0, 0.0, 0.0, 0.0)
    first3 = [excel_round(v / total, 3) for v in band_kwh[:3]]
    return (*first3, 1 - sum(first3))


def allocate(kwh: float, shares: Sequence[float]) -> tuple:
    """月合計kWhを区分の割合で按分(小数1桁)。"""
    return tuple(excel_round(s * kwh, 1) for s in shares)


def market_cost(allocated: Sequence[float], prices: Sequence[float]) -> float:
    return sum(a * p for a, p in zip(allocated, prices))


def calc_bill(kwh: float, band_kwh: Sequence[float], p: SiteParams) -> Bill:
    """月合計kWhと区分別kWh(当月の日次合計。HO3は先月分)から料金を試算する。"""
    shares = band_shares(band_kwh)
    allocated = allocate(kwh, shares)
    if p.market_flat is None:
        market = market_cost(allocated, p.band_prices)
    else:
        market = kwh * p.market_flat

    base = p.base_unit * p.power_factor * p.contract_kw
    energy = kwh * p.unit_price
    fuel = kwh * p.fuel_adj
    renewable = kwh * p.renewable
    annual = kwh * p.annual_discount
    incl = base + energy + fuel + renewable + p.fixed_credit + annual + market
    excl = incl / (1 + p.tax_rate)
    return Bill(kwh, shares, allocated, market, base, energy, fuel, renewable,
                p.fixed_credit, annual, incl, excl)


def billing_period(year: int, month: int, start_day: int = 27):
    """検針期間(開始日, 終了日)。start_day=27 なら 前月27日〜当月26日、1 なら 当月1日〜月末。"""
    from datetime import date, timedelta
    if start_day <= 1:
        nxt = date(year + (month == 12), month % 12 + 1, 1)
        return date(year, month, 1), nxt - timedelta(days=1)
    py, pm = (year - 1, 12) if month == 1 else (year, month - 1)
    return date(py, pm, start_day), date(year, month, start_day - 1)


def billing_month(d, start_day: int = 27):
    """日付が属する検針月(年, 月)。start_day=27 なら 27日以降は翌月分、1 なら暦月。"""
    if start_day <= 1 or d.day < start_day:
        return d.year, d.month
    return (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
