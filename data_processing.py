from pathlib import Path
from typing import List, Dict, Any, Iterable
import pandas as pd


# =========================
# Утилиты для CSV
# =========================
def _smart_read_csv(p: Path) -> pd.DataFrame:
    """Чтение CSV с попытками разных кодировок и разделителей."""
    for enc in ("utf-8-sig", "cp1251", "utf-8"):
        for sep in (None, ";", ","):
            try:
                return pd.read_csv(p, sep=sep, engine="python", encoding=enc)
            except Exception:
                continue
    # Если всё не сработало — пусть pandas поднимет явную ошибку
    return pd.read_csv(p)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Нормализуем заголовки: нижний регистр, обрезка пробелов, замена ё→е."""
    cols = []
    for c in df.columns:
        norm = str(c).strip().lower().replace("ё", "е")
        cols.append(norm)
    out = df.copy()
    out.columns = cols
    return out


def _pick_column(df_norm: pd.DataFrame, candidates: Iterable[str]) -> str:
    """Ищем подходящую колонку: сначала точное совпадение, затем подстрока."""
    for cand in candidates:
        if cand in df_norm.columns:
            return cand
    for col in df_norm.columns:
        for cand in candidates:
            if cand in col:
                return col
    raise KeyError(f"Не найдена колонка из списка: {list(candidates)}. Доступные: {list(df_norm.columns)}")


def _count_in_df(df: pd.DataFrame) -> Dict[str, int]:
    """Подсчёт: выполнено всего, постоматы (П), остальные."""
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
    Принимает список путей (можно смешанные файлы).
    Считает итоги по CSV; Excel игнорируется.
    """
    totals = {"total_completed": 0, "postomats": 0, "others": 0}
    per_file = []
    errors = []

    for p in paths:
        if p.suffix.lower() != ".csv":
            continue
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


# =========================
# «КТ» обработка Excel
# =========================
# Оставляем ТОЛЬКО эти столбцы (точные имена, как в файле).
KEEP_COLUMNS_KT = [
    "Номер заказа",
    "Код офиса местонахождения",
    "Контрольная точка",
    "Дней в КТ",
    "Тип заказа",
]


def _autosize_columns_xlsx(xlsx_path: Path, df_out: pd.DataFrame) -> None:
    """Автоподбор ширины столбцов под содержимое (openpyxl)."""
    import openpyxl
    from openpyxl.utils import get_column_letter

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    for col_idx, col_name in enumerate(df_out.columns, start=1):
        max_len = len(str(col_name)) if col_name is not None else 0
        # учитываем содержимое всех строк в данном столбце
        for val in df_out[col_name].astype(str).fillna("").values:
            if val is None:
                continue
            ln = len(str(val))
            if ln > max_len:
                max_len = ln
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2  # небольшой отступ
    wb.save(xlsx_path)


def process_kt_excels(paths: List[Path], keep_columns: List[str] = None) -> Dict[str, Any]:
    """
    Обрабатывает только Excel (.xlsx/.xls/.xlsm):
      - оставляет только указанные столбцы,
      - сохраняет <name>_KT.xlsx,
      - автоматически подгоняет ширину столбцов под содержимое.
    Возвращает: {"saved": [...], "skipped": [...], "errors": [...]}
    """
    if keep_columns is None:
        keep_columns = KEEP_COLUMNS_KT

    saved, skipped, errors = [], [], []

    for p in paths:
        ext = p.suffix.lower()
        if ext not in (".xlsx", ".xls", ".xlsm"):
            skipped.append({"file": str(p), "reason": "Не Excel"})
            continue
        try:
            # гарантируем наличие openpyxl
            try:
                import openpyxl  # noqa: F401
            except ImportError:
                raise RuntimeError("Требуется пакет 'openpyxl'. Установи: pip install openpyxl")

            df = pd.read_excel(p)

            existing = [c for c in keep_columns if c in df.columns]
            if not existing:
                raise KeyError(
                    "Ни один из требуемых столбцов не найден.\n"
                    f"В файле есть: {list(df.columns)}\n"
                    f"Ожидались: {keep_columns}"
                )

            df_out = df[existing].copy()
            out_path = p.with_name(f"{p.stem}_KT.xlsx")
            df_out.to_excel(out_path, index=False)

            # авто ширина
            _autosize_columns_xlsx(out_path, df_out)

            saved.append({
                "file": str(p),
                "saved_as": str(out_path),
                "kept_columns": existing
            })

        except Exception as e:
            errors.append({"file": str(p), "error": str(e)})

    return {"saved": saved, "skipped": skipped, "errors": errors}