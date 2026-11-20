"""電気代 入力・閲覧アプリ(Streamlit)。スマホ前提の1カラム表示。"""
import hmac
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import analysis as an
from calc import excel_round
from storage import make_store

st.set_page_config(page_title="電気代", page_icon="⚡", layout="centered")


def secret(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


# ---------- 合言葉(パスワード) ----------
def require_password():
    pw = secret("app_password")
    if not pw:
        return
    if st.session_state.get("authed"):
        return
    with st.form("login"):
        typed = st.text_input("パスワード", type="password")
        if st.form_submit_button("開く"):
            if hmac.compare_digest(str(typed).encode(), str(pw).encode()):
                st.session_state["authed"] = True
                st.rerun()
            else:
                st.error("パスワードが違います。")
    st.stop()


require_password()


# ---------- データの読み込み ----------
@st.cache_resource
def get_store():
    info = secret("gcp_service_account")
    return make_store(dict(info) if info else None, secret("spreadsheet_id"))


@st.cache_data(ttl=60, show_spinner="読み込み中…")
def load_all():
    s = get_store()
    return (an.normalize_daily(s.read("daily")),
            an.normalize_params(s.read("params")),
            an.normalize_sites(s.read("sites")))


daily, params, sites = load_all()
if get_store().is_demo:
    st.info("デモモードです(架空のデータを表示しています)。")
if sites.empty:
    st.error("sitesタブに拠点が登録されていません。")
    st.stop()

site_names = list(sites["site"])


def site_cfg(name):
    r = sites[sites["site"] == name].iloc[0]
    return int(r["start_day"]), bool(r["has_new_bldg"])


def yen(x):
    return f"{excel_round(x):,.0f}"


st.title("⚡ 電気代")
tab_in, tab_imp, tab_view, tab_bill, tab_csv = st.tabs(["入力", "取込", "閲覧", "料金", "CSV"])

# ---------- 入力 ----------
with tab_in:
    site = st.selectbox("拠点", site_names, key="in_site")
    start_day, has_new = site_cfg(site)
    ds = daily[daily["site"] == site]
    today = date.today()
    latest = None if ds.empty else ds["date"].max().date()
    default_day = today if latest is None else min(latest + timedelta(days=1), today)
    day = st.date_input("日付", value=default_day, key=f"in_day_{site}")

    hit = ds[ds["date"] == pd.Timestamp(day)]
    ex = None if hit.empty else hit.iloc[0]

    def pre(col):
        return 0.0 if ex is None or pd.isna(ex[col]) else float(ex[col])

    st.caption("この日の値が登録済みです。保存すると上書きします。" if ex is not None else "この日はまだ登録されていません。")
    k = f"{site}_{day}"
    total = st.number_input("全体 (kWh)", min_value=0.0, step=1.0, format="%.0f", value=pre("total_kwh"), key=f"tot_{k}")
    new_bldg = None
    if has_new:
        new_bldg = st.number_input("新棟 (kWh)", min_value=0.0, step=1.0, format="%.0f",
                                   value=pre("new_bldg_kwh"), key=f"new_{k}")
    c1, c2 = st.columns(2)
    bands = [
        c1.number_input("朝時間", min_value=0.0, step=1.0, format="%.0f", value=pre("morning"), key=f"bm_{k}"),
        c2.number_input("昼時間", min_value=0.0, step=1.0, format="%.0f", value=pre("noon"), key=f"bn_{k}"),
        c1.number_input("晩時間", min_value=0.0, step=1.0, format="%.0f", value=pre("evening"), key=f"be_{k}"),
        c2.number_input("夜時間", min_value=0.0, step=1.0, format="%.0f", value=pre("night"), key=f"bi_{k}"),
    ]
    if total > 0 and sum(bands) > 0:
        diff = sum(bands) - total
        if abs(diff) > total * 0.02:
            st.warning(f"朝〜夜の合計が全体より {diff:+,.0f} kWh ずれています。入力を確認してください。")
        else:
            st.caption(f"朝〜夜の合計との差: {diff:+,.0f} kWh")

    if st.button("保存", type="primary", key=f"save_{k}"):
        if total <= 0:
            st.error("全体のkWhが0です。")
        else:
            no_bands = sum(bands) == 0
            rec = {"site": site, "date": day.isoformat(), "total_kwh": total,
                   "new_bldg_kwh": new_bldg if has_new else None,
                   "morning": None if no_bands else bands[0], "noon": None if no_bands else bands[1],
                   "evening": None if no_bands else bands[2], "night": None if no_bands else bands[3],
                   "source": "actual"}
            res = get_store().upsert_daily(rec)
            st.cache_data.clear()
            st.success("上書きしました。" if res == "updated" else "保存しました。")
            daily, params, sites = load_all()

    recent = daily[daily["site"] == site].tail(7).copy()
    if not recent.empty:
        st.markdown("**直近の入力**")
        recent["date"] = recent["date"].dt.strftime("%m/%d")
        recent = recent.rename(columns={"date": "日付", "total_kwh": "全体", "new_bldg_kwh": "新棟", "morning": "朝",
                                        "noon": "昼", "evening": "晩", "night": "夜"})
        cols = ["日付", "全体"] + (["新棟"] if has_new else []) + ["朝", "昼", "晩", "夜"]
        st.dataframe(recent[cols], hide_index=True)

# ---------- 取込(収集プログラムの出力) ----------
with tab_imp:
    st.caption("収集プログラムが作ったExcelを選ぶと、朝・昼・晩・夜を日ごとに取り込みます。複数日分をまとめて取り込めます。")
    up = st.file_uploader("収集プログラムのExcel(.xlsx)", type=["xlsx"], key="imp_file")
    today = date.today()
    ci1, ci2 = st.columns(2)
    imp_year = int(ci1.number_input("年", min_value=2020, max_value=2100, value=today.year, step=1, key="imp_y"))
    imp_month = int(ci2.number_input("月", min_value=1, max_value=12, value=today.month, step=1, key="imp_m"))
    st.caption("ファイルには日(1〜31)しか入っていないので、何年何月分かをここで指定します。")
    upd_total = st.checkbox("登録済みの日の「全体」も、朝〜夜の合計で更新する", value=False, key="imp_total")
    if up is not None:
        try:
            new_rows, warns = an.parse_collector_xlsx(up, imp_year, imp_month, sites)
        except Exception as e:  # 壊れたファイルなど
            st.error(f"ファイルを読み取れませんでした: {e}")
            new_rows, warns = None, []
        for w in warns:
            st.warning(w)
        if new_rows is not None and new_rows.empty:
            st.info("取り込める日がありませんでした。")
        elif new_rows is not None:
            plan = an.plan_import(daily, new_rows, today, update_total=upd_total)
            n_new = int((plan["status"] == "新規").sum())
            n_upd = int((plan["status"] == "更新").sum())
            n_same = int((plan["status"] == "同じ").sum())
            st.markdown(f"**新規 {n_new}日 / 更新 {n_upd}日 / 変更なし {n_same}日**")
            view = plan.copy()
            view["date"] = view["date"].dt.strftime("%m/%d")
            view = view.rename(columns={"status": "状態", "site": "拠点", "date": "日付", "total_kwh": "全体",
                                        "morning": "朝", "noon": "昼", "evening": "晩", "night": "夜", "note": "備考"})
            st.dataframe(view[["状態", "拠点", "日付", "全体", "朝", "昼", "晩", "夜", "備考"]], hide_index=True)
            miss = an.missing_days(daily, plan, site_names, imp_year, imp_month, today)
            if miss:
                st.warning("まだ入っていない日: " + " / ".join(f"{s} {', '.join(r)}" for s, r in miss.items()))
            recs = an.plan_records(plan)
            if st.button(f"この内容を保存({len(recs)}日分)", type="primary", disabled=not recs, key="imp_save"):
                added, updated = get_store().upsert_daily_many(recs)
                st.cache_data.clear()
                st.success(f"保存しました(新規 {added}日、更新 {updated}日)。")
                daily, params, sites = load_all()

# ---------- 閲覧 ----------
with tab_view:
    site = st.selectbox("拠点", site_names, key="v_site")
    start_day, has_new = site_cfg(site)
    months = an.months_available(daily, site, start_day)
    if not months:
        st.info("この拠点のデータがありません。")
    else:
        ym = st.selectbox("検針月", months, key="v_month", format_func=lambda x: an.month_label(x[0], x[1], start_day))
        s = an.month_summary(daily, params, site, start_day, *ym)
        m1, m2, m3 = st.columns(3)
        m1.metric("入力日数", f"{s['days_filled']}/{s['days_total']}")
        m2.metric("合計 kWh", f"{s['kwh']:,.0f}")
        if s["prev_kwh"]:
            pct = (sum(s["frame"].loc[s["frame"]["total_kwh"].notna() & s["frame"]["prev_total"].notna(), "total_kwh"])
                   / s["prev_kwh"] - 1) * 100
            m3.metric("前年同期比", f"{pct:+.1f}%")
        else:
            m3.metric("前年同期比", "-")
        f = s["frame"]
        st.line_chart(f[["total_kwh", "prev_total"]].rename(columns={"total_kwh": "今年", "prev_total": "前年同日"}))
        t = f.reset_index().rename(columns={"index": "日付", "total_kwh": "全体", "new_bldg_kwh": "新棟", "morning": "朝",
                                            "noon": "昼", "evening": "晩", "night": "夜", "prev_total": "前年同日"})
        t["日付"] = t["日付"].dt.strftime("%m/%d")
        cols = ["日付", "全体"] + (["新棟"] if has_new else []) + ["朝", "昼", "晩", "夜", "前年同日"]
        st.dataframe(t[cols], hide_index=True)

# ---------- 料金 ----------
with tab_bill:
    site = st.selectbox("拠点", site_names, key="b_site")
    start_day, _ = site_cfg(site)
    months = an.months_available(daily, site, start_day)
    if not months:
        st.info("この拠点のデータがありません。")
    else:
        ym = st.selectbox("検針月", months, key="b_month", format_func=lambda x: an.month_label(x[0], x[1], start_day))
        s = an.month_summary(daily, params, site, start_day, *ym)
        for n in s["notes"]:
            st.warning(n)
        bill = s["bill"]
        if bill is not None:
            a, b = st.columns(2)
            a.metric("税込み", f"{yen(bill.tax_included)} 円")
            b.metric("税抜き", f"{yen(bill.tax_excluded)} 円")
            items = [("基本料金", bill.base), ("電力量料金", bill.energy), ("燃料費調整", bill.fuel),
                     ("再エネ賦課金", bill.renewable), ("産蓄熱等(固定)", bill.fixed_credit),
                     ("単年割", bill.annual_discount), ("市場価格調整", bill.market)]
            st.dataframe(pd.DataFrame({"項目": [i[0] for i in items], "金額(円)": [yen(i[1]) for i in items]}),
                         hide_index=True)
            p = s["params"]
            band_df = pd.DataFrame({"区分": an.BAND_JA, "割合": [f"{x:.3f}" for x in bill.shares],
                                    "按分kWh": [f"{x:,.1f}" for x in bill.allocated],
                                    "単価": list(p.band_prices),
                                    "金額(円)": [yen(x * y) for x, y in zip(bill.allocated, p.band_prices)]})
            st.markdown("**市場価格調整の内訳**")
            st.dataframe(band_df, hide_index=True)

# ---------- CSV ----------
with tab_csv:
    chosen = st.multiselect("拠点", site_names, default=site_names, key="c_sites")
    if daily.empty:
        st.info("データがありません。")
    else:
        lo, hi = daily["date"].min().date(), daily["date"].max().date()
        rng = st.date_input("期間", value=(max(lo, hi - timedelta(days=89)), hi), key="c_range")
        if not isinstance(rng, (tuple, list)) or len(rng) != 2:
            st.info("期間の開始日と終了日を選んでください。")
        else:
            d_from, d_to = rng
            sel = daily[daily["site"].isin(chosen) & (daily["date"] >= pd.Timestamp(d_from))
                        & (daily["date"] <= pd.Timestamp(d_to))].copy()
            sel["date"] = sel["date"].dt.strftime("%Y-%m-%d")
            sel = sel.rename(columns={"site": "拠点", "date": "日付", "total_kwh": "全体kWh", "new_bldg_kwh": "新棟kWh",
                                      "morning": "朝", "noon": "昼", "evening": "晩", "night": "夜", "source": "区分"})
            st.download_button("日次データをCSVで保存", sel.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"denki_daily_{d_from}_{d_to}.csv", mime="text/csv")
            mt = an.monthly_table(daily, params, sites, chosen, d_from, d_to)
            st.download_button("月ごとの集計をCSVで保存", mt.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"denki_monthly_{d_from}_{d_to}.csv", mime="text/csv", disabled=mt.empty)
            st.caption(f"日次 {len(sel)}行 / 月次 {len(mt)}行")
