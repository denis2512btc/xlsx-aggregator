"""Общие фикстуры для pytest."""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

import pytest

_FIXTURES = Path(__file__).resolve().parent / "fixtures"


def _first_xlsx() -> Path | None:
    """Первый ``*.xlsx`` в ``fixtures/`` (файлы в git не храним — только локально)."""
    for p in _FIXTURES.glob("*.xlsx"):
        return p
    return None


def _require_example_xlsx() -> Path:
    p = _first_xlsx()
    if p is None:
        pytest.skip(
            f"Нет .xlsx в {_FIXTURES} — положите пример локально для интеграционных тестов."
        )
    return p


@pytest.fixture
def sample_xlsx_path() -> Path:
    """Путь к копии примера (временный файл, чтобы не портить оригинал)."""
    src = _require_example_xlsx()
    fd, name = tempfile.mkstemp(suffix=".xlsx")
    import os

    os.close(fd)
    dst = Path(name)
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def example_workbook_path() -> Path:
    """Ссылка на xlsx в fixtures/ (только чтение)."""
    return _require_example_xlsx()
