"""Тесты JOIN SCPF / S5PF."""

from __future__ import annotations

import pandas as pd

from src.core.joiner import build_account_table
from src.core.config import ACCOUNT_FIELD_KEY_HEADER


def test_s5_missing_gives_nan() -> None:
    accounts = {
        (("YW3AB2", "YW3AN2", "YW3AS2"), ("A", "B", "C")),
    }
    sc_rows = [
        {
            "SCAB": "A",
            "SCAN": "B",
            "SCAS": "C",
            "SCACT": "1",
            "SCSAC": "2",
            "SCNANC": "3",
            "SCCCY": "EUR",
            "SCBAL": 0,
            "SCSUM0": 0,
            "SCSUMD": 0,
            "SCSUMC": 0,
            "SCRBA": 0,
        }
    ]
    s5_rows: list[dict] = []
    df = build_account_table(accounts, sc_rows, s5_rows)
    assert df.columns[0] == ACCOUNT_FIELD_KEY_HEADER
    assert df.loc[0, ACCOUNT_FIELD_KEY_HEADER] == "YW3AB2-YW3AN2-YW3AS2"
    assert pd.isna(df.loc[0, "S5BAL"])
    assert pd.isna(df.loc[0, "S5AIMD"])


def test_empty_accounts_empty_frame() -> None:
    df = build_account_table(set(), [{"SCAB": "x", "SCAN": "y", "SCAS": "z"}], [])
    assert df.empty
    assert list(df.columns)[0] == ACCOUNT_FIELD_KEY_HEADER
