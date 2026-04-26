# XLSX Aggregator

**Репозиторий:** [github.com/denis2512btc/xlsx-aggregator](https://github.com/denis2512btc/xlsx-aggregator)

Десктопное приложение для Windows, которое консолидирует данные из листов Excel в `YW2PF`, строит таблицу счетов (JOIN `SCPF` + опционально `S5PF`) и безопасно перезаписывает исходный `.xlsx` с резервной копией.

## Требования

- Windows 10/11
- Python 3.11+ ([python.org](https://www.python.org/downloads/)) в `PATH` (команда `python`). На очень новых версиях Python используйте актуальный `pandas` с готовым wheel (см. `requirements.txt`).

## Быстрый старт (разработка)

1. Клонируйте/скопируйте проект в папку, например `C:\Projects\Items`.
2. Дважды запустите `run.bat` (создастся `.venv` и установятся зависимости) или вручную:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m src.main
```

## Сборка `.exe` (PyInstaller, one-folder)

```bat
build.bat
```

Результат: `dist\XLSX Aggregator\XLSX Aggregator.exe`.

**Замечание:** Windows Defender иногда помечает неподписанные сборки PyInstaller как подозрительные. Варианты: подпись code-signing сертификатом, добавление `dist\` в исключения на тестовой машине, или доверие к источнику.

## Логи

Файлы логов: `%LOCALAPPDATA%\xlsx_aggregator\logs\` (например, `C:\Users\<User>\AppData\Local\xlsx_aggregator\logs\app.log`).

## Тесты

Файлы `.xls` / `.xlsx` в репозиторий не входят. Для интеграционных тестов положите локально любой ``*.xlsx`` в [tests/fixtures](tests/fixtures) (копия «Пример файла.xlsx»). Без этого соответствующие тесты будут **пропущены** (skip).

```bat
.venv\Scripts\activate
pytest -v tests
```

## CLI (без GUI)

```bat
python -m src.core.pipeline "C:\path\to\file.xlsx"
```

## Исходные требования

Подробный технический план: [PLAN.md](PLAN.md).
