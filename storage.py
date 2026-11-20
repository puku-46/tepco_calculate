"""保存先。Googleスプレッドシート(本番)と、手元のCSVフォルダ(デモ・テスト)を同じ使い方で扱う。"""
from __future__ import annotations

import os

import pandas as pd

from analysis import DAILY_COLS, PARAM_COLS, SITE_COLS

TABS = {"daily": DAILY_COLS, "params": PARAM_COLS, "sites": SITE_COLS}


def _clean(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    return v


def _s(v) -> str:
    """CSVに書く文字列にする(1234.0 → 1234)。"""
    v = _clean(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _norm_date(x) -> str:
    d = pd.to_datetime(str(x).strip(), format="mixed", errors="coerce")
    return "" if pd.isna(d) else d.strftime("%Y-%m-%d")


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class CsvStore:
    """手元のCSVフォルダを保存先にする(デモ用)。"""
    is_demo = True

    def __init__(self, folder: str):
        self.folder = folder

    def _path(self, name):
        return os.path.join(self.folder, f"{name}.csv")

    def read(self, name: str) -> pd.DataFrame:
        p = self._path(name)
        if not os.path.exists(p):
            return pd.DataFrame(columns=TABS[name])
        return pd.read_csv(p, dtype=str, encoding="utf-8-sig")

    def upsert_daily(self, rec: dict) -> str:
        df = self.read("daily")
        row = {c: _s(rec.get(c)) for c in DAILY_COLS}
        target = _norm_date(row["date"])
        hit = (df["site"].astype(str).str.strip() == rec["site"]) & (
            pd.to_datetime(df["date"], format="mixed", errors="coerce").dt.strftime("%Y-%m-%d") == target)
        if hit.any():
            for c, v in row.items():
                df.loc[hit, c] = v
            result = "updated"
        else:
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            result = "added"
        os.makedirs(self.folder, exist_ok=True)
        df.to_csv(self._path("daily"), index=False, encoding="utf-8-sig")
        return result


    def upsert_daily_many(self, recs: list) -> tuple:
        """複数日分をまとめて保存する。(追加数, 更新数)を返す。"""
        df = self.read("daily")
        norm = pd.to_datetime(df["date"], format="mixed", errors="coerce").dt.strftime("%Y-%m-%d") if len(df) else pd.Series(dtype=str)
        added = updated = 0
        for rec in recs:
            row = {c: _s(rec.get(c)) for c in DAILY_COLS}
            hit = (df["site"].astype(str).str.strip() == rec["site"]) & (norm == _norm_date(row["date"])) if len(df) else pd.Series(dtype=bool)
            if len(df) and hit.any():
                for c, v in row.items():
                    df.loc[hit, c] = v
                updated += 1
            else:
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                norm = pd.concat([norm, pd.Series([_norm_date(row["date"])])], ignore_index=True)
                added += 1
        os.makedirs(self.folder, exist_ok=True)
        df.to_csv(self._path("daily"), index=False, encoding="utf-8-sig")
        return added, updated


class SheetsStore:
    """Googleスプレッドシート(daily / params / sites の3タブ)を保存先にする。"""
    is_demo = False

    def __init__(self, spreadsheet):
        self.sh = spreadsheet

    @classmethod
    def from_secrets(cls, info, spreadsheet_id: str):
        import gspread
        gc = gspread.service_account_from_dict(dict(info))
        return cls(gc.open_by_key(spreadsheet_id))

    def read(self, name: str) -> pd.DataFrame:
        values = self.sh.worksheet(name).get_all_values()
        if not values:
            return pd.DataFrame(columns=TABS[name])
        header, rows = values[0], values[1:]
        return pd.DataFrame(rows, columns=header).replace({"": pd.NA})

    def upsert_daily(self, rec: dict) -> str:
        ws = self.sh.worksheet("daily")
        values = ws.get_all_values()
        if not values:
            ws.append_row(DAILY_COLS, value_input_option="RAW")
            values = [DAILY_COLS]
        header = values[0]
        row = [_clean(rec.get(c)) for c in header]
        i_site, i_date = header.index("site"), header.index("date")
        target = _norm_date(rec["date"])
        for n, r in enumerate(values[1:], start=2):
            if len(r) > max(i_site, i_date) and r[i_site].strip() == rec["site"] and _norm_date(r[i_date]) == target:
                ws.update(values=[row], range_name=f"A{n}:{_col_letter(len(header))}{n}", value_input_option="RAW")
                return "updated"
        ws.append_row(row, value_input_option="RAW")
        return "added"


    def upsert_daily_many(self, recs: list) -> tuple:
        """複数日分をまとめて保存する(読み取り1回、更新は1回のbatch、追加は1回)。(追加数, 更新数)を返す。"""
        ws = self.sh.worksheet("daily")
        values = ws.get_all_values()
        if not values:
            ws.append_row(DAILY_COLS, value_input_option="RAW")
            values = [DAILY_COLS]
        header = values[0]
        i_site, i_date = header.index("site"), header.index("date")
        index = {}
        for n, r in enumerate(values[1:], start=2):
            if len(r) > max(i_site, i_date):
                index[(r[i_site].strip(), _norm_date(r[i_date]))] = n
        updates, appends = [], []
        for rec in recs:
            row = [_clean(rec.get(c)) for c in header]
            n = index.get((rec["site"], _norm_date(rec["date"])))
            if n:
                updates.append({"range": f"A{n}:{_col_letter(len(header))}{n}", "values": [row]})
            else:
                appends.append(row)
        if updates:
            ws.batch_update(updates, value_input_option="RAW")
        if appends:
            ws.append_rows(appends, value_input_option="RAW")
        return len(appends), len(updates)


def make_store(info, spreadsheet_id, sample_dir: str = "sample_data"):
    """Secretsにスプレッドシートの設定があれば本番、なければデモ(sample_data)。"""
    if info and spreadsheet_id:
        return SheetsStore.from_secrets(info, spreadsheet_id)
    return CsvStore(sample_dir)
