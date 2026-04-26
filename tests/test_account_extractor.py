"""Тесты извлечения троек AB/AN/AS, BB/BN/BS."""

from __future__ import annotations

from src.core.account_extractor import (
    extract_account_slots,
    extract_all_accounts,
)


def test_yw3anr_is_not_matched_as_slot() -> None:
    """YW3ANR не формирует тройку (нет пары AB/AS)."""
    headers = ["YW3ANR", "YW3AB2", "YW3AN2", "YW3AS2"]
    slots = extract_account_slots(headers, "YW3")
    slot_ids = [s[0] for s in slots if s[1] == "A"]
    assert "R" not in slot_ids  # ANR+... не даёт валидной тройки с тем же slot


def test_full_triple_gives_slot() -> None:
    h = [
        "YW3AB1",
        "YW3AN1",
        "YW3AS1",
    ]
    slots = extract_account_slots(h, "YW3")
    assert any(s[0] == "1" and s[1] == "A" for s in slots)


def test_ywj1_no_suffix_gives_slot_a() -> None:
    """Колонки YWJ1AB/AN/AS без суффикса после буквенной группы — один слот A."""
    h = ["YWJ1AB", "YWJ1AN", "YWJ1AS"]
    slots = extract_account_slots(h, "YWJ1")
    assert any(s[0] == "" and s[1] == "A" for s in slots)


def test_extract_all_accounts_skips_non_allowed_columns() -> None:
    """В extract_all_accounts попадают только значения из allowlist-колонок."""
    sheets = {
        "YW3PF": [
            {
                "YW3AB1": "1111",
                "YW3AN1": "X",
                "YW3AS1": "001",
                "YW3AB2": "0880",
                "YW3AN2": "A",
                "YW3AS2": "006",
            },
        ]
    }
    acc = extract_all_accounts(sheets)  # type: ignore[arg-type]
    key2 = ("YW3AB2", "YW3AN2", "YW3AS2")
    assert (key2, ("0880", "A", "006")) in acc
    key1 = ("YW3AB1", "YW3AN1", "YW3AS1")
    assert not any(cols == key1 for cols, _ in acc)


def test_extract_merges_rows() -> None:
    sheets = {
        "YW3PF": [
            {"YW3AB2": "0880", "YW3AN2": "A", "YW3AS2": "006", "YW3ANR": "x"},
        ]
    }
    acc = extract_all_accounts(sheets)  # type: ignore[arg-type]
    assert (("YW3AB2", "YW3AN2", "YW3AS2"), ("0880", "A", "006")) in acc


def test_zero_rows_ywj1() -> None:
    sheets = {
        "YW2PF": [],
        "YW3PF": [],
        "YWJ1PF": [],
    }
    acc = extract_all_accounts(sheets)  # type: ignore[arg-type]
    assert acc == set()
