"""既存Excelの実績値との突き合わせテスト。  実行: python -m pytest -q  (または python test_calc.py)"""
from calc import SiteParams, calc_bill


def near(a, b, tol=0.5):
    assert abs(a - b) <= tol, f"{a} != {b}"


def test_ho1_2026_07():
    p = SiteParams(contract_kw=1200, fuel_adj=1.47, renewable=4.18, annual_discount=-3,
                   fixed_credit=-1021180, band_prices=(4.13, 6.15, 5.77, 2.99))
    b = calc_bill(628195, (104970, 73414, 146530, 303284), p)
    near(b.market, 2637061.994)
    near(b.tax_included, 16810637.594)
    near(b.tax_excluded, 15282397.813)


def test_ho2_2026_09():  # Excel側は単年割が0固定(要確認)
    p = SiteParams(contract_kw=1520, annual_discount=0.0)
    b = calc_bill(752459, (177569, 105966, 198940, 268996), p)
    near(b.market, 3840497.513)
    near(b.tax_included, 23068912.903)


if __name__ == "__main__":
    test_ho1_2026_07(); test_ho2_2026_09(); print("OK")
