"""Построение таблицы счетов: LEFT JOIN SCPF и S5PF по ключу (SCAB, SCAN, SCAS)."""

from __future__ import annotations

import pandas as pd
from loguru import logger

from src.core.config import (
    ACCOUNT_FIELD_KEY_HEADER,
    ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED,
    SCPF_MERGE_COLUMNS,
    S5PF_MERGE_COLUMNS,
)

_ACCOUNT_COL_ORDER: dict[tuple[str, str, str], int] = {
    t: i for i, t in enumerate(ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED)
}


def build_account_table(
    accounts: set[tuple[tuple[str, str, str], tuple[str, str, str]]],
    sc_rows: list[dict],
    s5_rows: list[dict],
) -> pd.DataFrame:
    """Строит итоговую таблицу счетов через LEFT JOIN SCPF + LEFT JOIN S5PF.

    ТЗ: «По каждому счету с Листа SCPF и S5PF по ключу …
    SCAB,SCAN,SCAS, …, S5BAL, S5AIMD, S5AM1D, -(S5AIMD+S5AM1D)»

    Нормализация ключей: компоненты приводятся к строке и ``.strip()``.

    Args:
        accounts: Уникальные пары ``(тройка имён колонок PF, (SCAB, SCAN, SCAS))``
            из ``extract_all_accounts``.
        sc_rows: Все data-строки SCPF.
        s5_rows: Все data-строки S5PF (могут отсутствовать часть счетов — штатно).

    Returns:
        DataFrame с колонками SC* и S5*; формула ``-(S5AIMD+S5AM1D)`` добавляется
        в ``writer.py``.
    """
    if not accounts:
        return pd.DataFrame(
            columns=[
                ACCOUNT_FIELD_KEY_HEADER,
                "SCAB",
                "SCAN",
                "SCAS",
                "SCACT",
                "SCSAC",
                "SCNANC",
                "SCCCY",
                "SCBAL",
                "SCSUM0",
                "SCSUMD",
                "SCSUMC",
                "SCRBA",
                "S5BAL",
                "S5AIMD",
                "S5AM1D",
            ]
        )

    def _acc_sort_key(
        item: tuple[tuple[str, str, str], tuple[str, str, str]],
    ) -> tuple[int, tuple[str, str, str]]:
        cols, _vals = item
        return (_ACCOUNT_COL_ORDER.get(cols, 10**9), cols)

    rows_sorted = sorted(accounts, key=_acc_sort_key)
    acc_df = pd.DataFrame(
        [
            {
                ACCOUNT_FIELD_KEY_HEADER: f"{a}-{b}-{c}",
                "SCAB": v[0],
                "SCAN": v[1],
                "SCAS": v[2],
            }
            for (a, b, c), v in rows_sorted
        ],
    )
    sc_df = pd.DataFrame(sc_rows)
    s5_df = pd.DataFrame(s5_rows)

    for df, cols in [
        (sc_df, ["SCAB", "SCAN", "SCAS"]),
        (s5_df, ["S5AB", "S5AN", "S5AS"]),
        (acc_df, ["SCAB", "SCAN", "SCAS"]),
    ]:
        for c in cols:
            if c in df.columns:
                df[c] = df[c].map(_cell_str).astype(str).str.strip()

    # дубликаты в справочнике — оставляем первую запись, чтобы merge не размножал строки
    if not sc_df.empty and all(c in sc_df.columns for c in ("SCAB", "SCAN", "SCAS")):
        sc_df = sc_df.drop_duplicates(subset=["SCAB", "SCAN", "SCAS"], keep="first")
    if not s5_df.empty and all(c in s5_df.columns for c in ("S5AB", "S5AN", "S5AS")):
        s5_df = s5_df.drop_duplicates(subset=["S5AB", "S5AN", "S5AS"], keep="first")

    missing_in_sc: set[tuple[str, str, str]] = set()
    if sc_df.empty or not all(c in sc_df.columns for c in ("SCAB", "SCAN", "SCAS")):
        for key in acc_df.itertuples(index=False):
            missing_in_sc.add((str(key.SCAB), str(key.SCAN), str(key.SCAS)))
    else:
        for key in acc_df.itertuples(index=False):
            k = (str(key.SCAB), str(key.SCAN), str(key.SCAS))
            m = (sc_df["SCAB"] == k[0]) & (sc_df["SCAN"] == k[1]) & (sc_df["SCAS"] == k[2])
            if not m.any():
                missing_in_sc.add(k)
    for k in missing_in_sc:
        logger.warning("Счёт {} не найден в SCPF (LEFT JOIN — колонки SC будут пустыми).", k)

    for required in ("SCAB", "SCAN", "SCAS"):
        if required not in sc_df.columns:
            raise ValueError(f"В SCPF отсутствует обязательная колонка {required}.")

    sc_use = sc_df.copy()
    for c in SCPF_MERGE_COLUMNS:
        if c not in sc_use.columns:
            sc_use[c] = pd.NA
    sc_sub = sc_use[SCPF_MERGE_COLUMNS]
    merged = acc_df.merge(sc_sub, on=["SCAB", "SCAN", "SCAS"], how="left")

    s5_needed = S5PF_MERGE_COLUMNS
    if s5_df.empty or not all(c in s5_df.columns for c in ("S5AB", "S5AN", "S5AS")):
        for col in ("S5BAL", "S5AIMD", "S5AM1D"):
            merged[col] = pd.NA
    else:
        s5_use = s5_df.copy()
        for c in s5_needed:
            if c not in s5_use.columns:
                s5_use[c] = pd.NA
        s5_sub = s5_use[s5_needed].rename(
            columns={"S5AB": "SCAB", "S5AN": "SCAN", "S5AS": "SCAS"}
        )
        merged = merged.merge(
            s5_sub,
            on=["SCAB", "SCAN", "SCAS"],
            how="left",
        )

    merged = merged[
        [
            ACCOUNT_FIELD_KEY_HEADER,
            "SCAB",
            "SCAN",
            "SCAS",
            "SCACT",
            "SCSAC",
            "SCNANC",
            "SCCCY",
            "SCBAL",
            "SCSUM0",
            "SCSUMD",
            "SCSUMC",
            "SCRBA",
            "S5BAL",
            "S5AIMD",
            "S5AM1D",
        ]
    ]
    return merged


def _cell_str(v: object) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()
