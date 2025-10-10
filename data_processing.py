import pandas as pd
from pathlib import Path

def _smart_read_csv(p: Path):
    # Пробуем частые кодировки и разделители (auto/;/,)
    for enc in ("utf-8-sig", "cp1251", "utf-8"):
        for sep in (None, ";", ","):
            try:
                df = pd.read_csv(p, sep=sep, engine="python", encoding=enc)
                return df
            except Exception:
                continue
    # Последняя попытка — пусть pandas сам ругнётся понятной ошибкой
    return pd.read_csv(p)

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        norm = str(c).strip().lower().replace("ё", "е")
        cols.append(norm)
    out = df.copy()
    out.columns = cols
    return out

def _pick_column(df_norm: pd.DataFrame, candidates) -> str | None:
    # точное попадание
    for cand in candidates:
        if cand in df_norm.columns:
            return cand
    # подстрока
    for col in df_norm.columns:
        for cand in candidates:
            if cand in col:
                return col
    return None

def analyze_csv(path: str | Path) -> dict:
    """
    Возвращает словарь:
      total_completed — выполнено всего
      postomats      — из них постоматы (Тип Адреса == 'П')
      others         — остальные
      used_status_col, used_type_col — какие колонки использованы
    """
    p = Path(path)
    df = _smart_read_csv(p)
    df_norm = _normalize_columns(df)

    status_col = _pick_column(df_norm, ["статус задания", "статус", "статус_задания"])
    type_col   = _pick_column(df_norm, ["тип адреса", "тип_адреса", "тип точки", "тип_точки", "тип"])

    if status_col is None:
        raise KeyError(f"Не найден столбец статуса. Доступные колонки: {list(df_norm.columns)}")
    if type_col is None:
        raise KeyError(f"Не найден столбец типа адреса. Доступные колонки: {list(df_norm.columns)}")

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
        "used_status_col": status_col,
        "used_type_col": type_col,
    }