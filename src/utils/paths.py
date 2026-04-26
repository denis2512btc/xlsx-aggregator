"""Пути Windows: логи в %LOCALAPPDATA%."""

from __future__ import annotations

import os
from pathlib import Path

APP_DIR_NAME = "xlsx_aggregator"


def get_local_app_data_dir() -> Path:
    """Возвращает каталог данных приложения под `%LOCALAPPDATA%`.

    Returns:
        Например ``C:\\Users\\...\\AppData\\Local\\xlsx_aggregator``.
    """
    base = os.getenv("LOCALAPPDATA")
    if not base:
        return Path.home() / "AppData" / "Local" / APP_DIR_NAME
    return Path(base) / APP_DIR_NAME


def get_logs_dir() -> Path:
    """Каталог файлов логов loguru (создаётся при настройке логгера)."""
    return get_local_app_data_dir() / "logs"
