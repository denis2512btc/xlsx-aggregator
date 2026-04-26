"""Настройка loguru: консоль + ротация в %LOCALAPPDATA%\\xlsx_aggregator\\logs\\."""

from __future__ import annotations

import sys
from pathlib import Path

from loguru import logger

from src.utils.paths import get_logs_dir


def setup_logging() -> None:
    """Снимает дефолтный sink loguru, добавляет stderr и файл с ротацией ~5 МБ."""
    logger.remove()
    logger.add(
        sys.stderr,
        level="INFO",
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <level>{message}</level>",
    )
    log_dir = get_logs_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"
    logger.add(
        str(log_file),
        level="DEBUG",
        rotation="5 MB",
        retention=10,
        encoding="utf-8",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {message}",
    )
    err_file = log_dir / "error.log"
    logger.add(
        str(err_file),
        level="ERROR",
        rotation="5 MB",
        encoding="utf-8",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {message}\n{exception}",
    )
