from pathlib import Path
from typing import List, Dict, Any
import pandas as pd


def _smart_read_csv(p: Path) -> pd.DataFrame:
    # Пробуем типичные кодировки и разделители
    for enc in ("utf-8-sig", "cp1251", "utf-8"):
        for sep in (None, ";", ","):
            try:
                return pd.read_csv(p, sep=sep, engine="python", encoding=enc)
            except Exception:
                continue
    # Последняя попытка — пусть pandas поднимет понятную ошибку
    return pd.read_csv(p)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        norm = str(c).strip().lower().replace("ё", "е")
        cols.append(norm)
    out = df.copy()
    out.columns = cols
    return out


def _pick_column(df_norm: pd.DataFrame, candidates) -> str:
    for cand in candidates:
        if cand in df_norm.columns:
            return cand
    for col in df_norm.columns:
        for cand in candidates:
            if cand in col:
                return col
    raise KeyError(f"Не найдена колонка из списка: {candidates}. Доступные: {list(df_norm.columns)}")


def _count_in_df(df: pd.DataFrame) -> Dict[str, int]:
    df_norm = _normalize_columns(df)

    status_col = _pick_column(df_norm, ["статус задания", "статус", "статус_задания"])
    type_col   = _pick_column(df_norm, ["тип адреса", "тип_адреса", "тип точки", "тип_точки", "тип"])

    status_series = (
        df_norm[status_col].astype(str).str.strip().str.lower().str.replace("ё", "е", regex=False)
    )
    done_mask = status_series.eq("выполнено")
    total_completed = int(done_mask.sum())

    type_series = df_norm[type_col].astype(str).str.strip().str.upper()
    postomats = int((done_mask & type_series.eq("П")).sum())
    others = int(total_completed - postomats)

    return {
        "total_completed": total_completed,
        "postomats": postomats,
        "others": others,
    }


def analyze_csvs(paths: List[Path]) -> Dict[str, Any]:
    """
    Принимает список путей к CSV.
    Возвращает словарь с общими итогами и списком ошибок по файлам, если были.
    {
      totals: { total_completed, postomats, others },
      per_file: [{file, total_completed, postomats, others}],
      errors: [{file, error}]
    }
    """
    totals = {"total_completed": 0, "postomats": 0, "others": 0}
    per_file = []
    errors = []

    for p in paths:
        try:
            df = _smart_read_csv(p)
            counts = _count_in_df(df)
            per_file.append({"file": str(p), **counts})
            totals["total_completed"] += counts["total_completed"]
            totals["postomats"] += counts["postomats"]
            totals["others"] += counts["others"]
        except Exception as e:
            errors.append({"file": str(p), "error": str(e)})

    return {"totals": totals, "per_file": per_file, "errors": errors}