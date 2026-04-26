# План реализации десктопного приложения «XLSX Aggregator»

> **Назначение документа.** Технический план реализации для Cursor. Документ содержит цель, стек, структуру проекта, UI, алгоритм обработки, предположения и пошаговый план внедрения. Всё рассчитано на то, чтобы разработчик (или AI-агент в Cursor) мог пройти по шагам и собрать готовое приложение.

---

## 1. Цель и контекст

Нужно десктопное приложение для автоматизации консолидации данных в Excel-файле (многолистовом, формат — см. `Пример_файла.xlsx`). Пользователь выбирает файл кнопкой и запускает обработку. Приложение:

1. Дописывает в лист **`YW2PF`** (ниже существующих данных, через пустую строку) блоки данных с других листов;
2. Собирает список счетов из листов `YW2PF`, `YW3PF`, `YWJ1PF` по маске полей `AB*/AN*/AS*` и `BB*/BN*/BS*`, **но только для троек колонок из allowlist** (`ALLOWED_ACCOUNT_FIELD_TRIPLES` в `config.py`; порядок вывода задаётся `ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED`);
3. Делает LEFT JOIN значений счёта с листами `SCPF` (всегда) и `S5PF` (если строка найдена), добавляет таблицу со счетами на `YW2PF`: **первая колонка — «Ключ полей PF»** (три имени колонок листа через дефис, напр. `YW3AB2-YW3AN2-YW3AS2`), далее `SCAB`…`S5AM1D` и вычисляемый столбец; накладывает автофильтр;
4. **Перезаписывает исходный файл** обработанной версией. Перед записью — автоматически создаёт резервную копию `<имя>.backup_<timestamp>.xlsx` рядом с оригиналом (см. раздел 6.8 «Безопасная перезапись»).

**Формат листов (везде одинаковый):**
- Строка 1: служебная ссылка «Go to Set Sheet» — игнорируется;
- Строка 2: имена полей (заголовки), напр. `YW2PT`, `SCAB`, `S5BAL`;
- Строка 3 и далее: данные.

---

## 2. Рекомендуемый стек

**Целевая платформа: Windows 10/11.** Всё кроссплатформенно, но выбор и инструкции оптимизированы под Windows.

| Слой | Технология | Обоснование |
|---|---|---|
| Язык | **Python 3.11+** (официальный installer с python.org) | Зрелые библиотеки для XLSX, совпадает с текущим стеком Дениса. Использовать `py -3.11` launcher. |
| GUI | **customtkinter 5.x** | Лёгкий (~50 МБ после PyInstaller), современный вид, идеально для UI из 2 кнопок. Работает на Windows без дополнительных зависимостей. |
| Работа с XLSX | **openpyxl 3.1+** | Сохраняет формулы, форматирование, автофильтр; поддержка `data_only` для чтения вычисленных значений |
| Табличная логика | **pandas 2.x** | Быстрый JOIN `SCPF`/`S5PF` по ключу (AB, AN, AS) |
| Логи | **loguru** | Однострочная настройка, ротация в файл |
| Сборка | **PyInstaller 6.x** | `.exe` (one-folder или one-file) одной командой |
| Тесты | **pytest** | Юнит-тесты на извлечение счетов и join |

**Альтернатива GUI:** PySide6 — мощнее, но тяжелее (бандл ~250 МБ, медленнее сборка). Для заявленного UI не нужен.

**Примечания по Windows:**
- Логи пишем в `%LOCALAPPDATA%\xlsx_aggregator\logs\` (`os.getenv("LOCALAPPDATA")`), **не** в домашнюю директорию.
- Открытие папки результата — `os.startfile(path)`, не `subprocess.run(["open", ...])`.
- Иконки — `.ico` (в `assets/icon.ico`).
- `os.replace(tmp, original)` атомарен на Windows **только в пределах одного тома**. Поэтому `tmp`-файл создаём в той же директории, что и оригинал (что мы и так делаем, см. 6.8).
- Windows Defender иногда помечает `.exe`, собранный PyInstaller, как подозрительный (false positive). Решение: либо подписать `.exe` code-signing сертификатом, либо один раз добавить в исключения. Упомянуть в README.

---

## 3. Структура проекта

```
xlsx_aggregator/
├── README.md
├── requirements.txt
├── pyproject.toml
├── build.bat                       # PyInstaller-команда для .exe (Windows)
├── run.bat                         # запуск из исходников в dev-режиме
├── assets/
│   └── icon.ico
├── src/
│   ├── __init__.py
│   ├── main.py                     # точка входа, запускает GUI
│   ├── gui/
│   │   ├── __init__.py
│   │   └── app.py                  # CTk-окно, 2 кнопки, прогресс-бар, лог
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py               # константы: имена листов, поля SC/S5
│   │   ├── sheet_reader.py         # чтение листа → list[dict] по заголовкам
│   │   ├── account_extractor.py    # слоты AB/AN/AS и BB/BN/BS; allowlist; пары (имена колонок, значения)
│   │   ├── joiner.py               # pandas-JOIN SCPF + S5PF
│   │   ├── writer.py               # дозапись блоков в YW2PF + автофильтр
│   │   └── pipeline.py             # оркестрация всего процесса + safe_overwrite_save
│   └── utils/
│       ├── __init__.py
│       ├── logging_setup.py        # loguru → %LOCALAPPDATA%\xlsx_aggregator\logs\
│       └── paths.py                # получение Windows-специфичных путей
└── tests/
    ├── conftest.py
    ├── fixtures/
    │   └── Пример_файла.xlsx       # реальный пример для интеграционных тестов
    ├── test_account_extractor.py
    ├── test_joiner.py
    ├── test_writer_idempotency.py  # повторный запуск даёт тот же результат
    ├── test_variable_records.py    # YWJ1PF/YW3PF с 0/1/N data-строками
    └── test_pipeline.py
```

---

## 4. Пользовательский интерфейс (customtkinter)

**Макет окна (500×400):**

```
┌──────────────────────────────────────────────┐
│   XLSX Aggregator                            │
│                                              │
│   Файл: [/path/to/file.xlsx        ] [Выбрать] │
│                                              │
│         ┌────────────────────────┐           │
│         │      Обработать        │           │
│         └────────────────────────┘           │
│                                              │
│   [===========░░░░░░░░░] 45%                 │
│                                              │
│   Лог:                                       │
│   ┌────────────────────────────────────┐    │
│   │ [12:03:01] Файл загружен            │   │
│   │ [12:03:02] Извлечено 17 счетов      │   │
│   │ ...                                 │   │
│   └────────────────────────────────────┘    │
│                                              │
│   Статус: Готов                              │
└──────────────────────────────────────────────┘
```

**Поведение:**
- **[Выбрать]** — `filedialog.askopenfilename` с фильтром `*.xlsx`. Путь отображается в поле (read-only).
- **[Обработать]** — disabled, пока не выбран файл. По клику:
  1. Показывает `CTkMessagebox` с подтверждением: «Файл **будет перезаписан**. Бэкап сохранится как `...backup_YYYYMMDD_HHMMSS.xlsx`. Продолжить?»
  2. При подтверждении запускает пайплайн в отдельном потоке (`threading.Thread`), чтобы окно не фризилось. Прогресс-бар обновляется из пайплайна через `queue.Queue` + `after(100)`-поллер.
- **Лог** — `CTkTextbox`, пишет события пайплайна в реальном времени.
- **По завершении** — диалог с сообщением «Готово. Бэкап: `...backup_....xlsx`» и кнопкой «Открыть папку» (`os.startfile(folder)` — штатно открывает Explorer на Windows).
- **При ошибке** — `CTkMessagebox` с текстом ошибки; полный traceback — в лог-файл `%LOCALAPPDATA%\xlsx_aggregator\logs\`.

---

## 5. Модель данных и ключевые константы (`src/core/config.py`)

```python
"""Центральная точка всех бизнес-констант из ТЗ.

Если бизнес-заказчик меняет название поля-триггера, добавляет новый лист
или новое поле в таблицу счетов — правится ТОЛЬКО этот файл.
"""

# =============================================================================
# Имена листов в исходной книге Excel
# =============================================================================

TARGET_SHEET = "YW2PF"
"""Лист, в который дописываем результат обработки.
ТЗ: «в существующем файле на Лист YW2PF … добавились данные».
"""

ALWAYS_APPEND = ["YW3PF", "YWJ1PF"]
"""Листы, данные с которых копируются на YW2PF безусловно.
ТЗ: «с Листа YW3PF / с Листа YWJ1PF».
ВАЖНО: количество data-строк на этих листах — переменное (YWJ1PF особенно
может содержать от 0 до N записей).
"""

CONDITIONAL_APPEND = [
    # ТЗ: «Если поле YW2PRZ2 <> '', то добавить данные с Листа AN6PF»
    {"sheet": "AN6PF", "trigger_field": "YW2PRZ2"},
    # ТЗ: «Если поле YW2PR5 <> '', то добавить данные с Листа AN9PF»
    # Уточнено с заказчиком: в ТЗ опечатка, корректное имя поля — YW2PRZ5.
    {"sheet": "AN9PF", "trigger_field": "YW2PRZ5"},
]
"""Листы, копируемые условно — только если значение указанного поля
из первой data-строки YW2PF непустое (`not _is_blank`)."""

ACCOUNT_SOURCE_SHEETS = ["YW2PF", "YW3PF", "YWJ1PF"]
"""Листы, с которых извлекаются ключи счетов.
ТЗ: «На Листах YW2PF, YW3PF и YWJ1PF указаны счета».
"""

SC_SHEET = "SCPF"   # обязательный источник для таблицы счетов
S5_SHEET = "S5PF"   # опциональный источник (ТЗ: «не все счета есть в S5PF»)

# Префиксы имён полей по имени листа (строка 2 — заголовки)
SHEET_FIELD_PREFIX = {"YW2PF": "YW2", "YW3PF": "YW3", "YWJ1PF": "YWJ1"}

# Разрешённые тройки имён колонок (AB/AN/AS или BB/BN/BS) и их порядок в таблице на выходе
ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED = (
    ("YW3AB2", "YW3AN2", "YW3AS2"),
    # … полный перечень в реальном config.py …
)
ALLOWED_ACCOUNT_FIELD_TRIPLES = frozenset(ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED)

ACCOUNT_FIELD_KEY_HEADER = "Ключ полей PF"  # первая колонка блока ACCOUNTS

# =============================================================================
# Метаданные листов (общая структура для всех PF-листов)
# =============================================================================

HEADER_ROW = 2
"""Строка 2 — имена полей (YW2PT, SCAB, S5BAL и т.д.).
Строка 1 — служебная ссылка «Go to Set Sheet», игнорируется полностью."""

DATA_START_ROW = 3
"""Первая строка с данными. Их может быть произвольное количество."""

# =============================================================================
# Зазоры и маркеры блоков
# =============================================================================

BLOCK_GAP = 1
"""Количество пустых строк ПЕРЕД каждым новым блоком (уточнено заказчиком)."""

BLOCK_MARKER_PREFIX = "[XA:"
BLOCK_MARKER_SUFFIX = "]"
"""Маркер в колонке A в виде `[XA:YW3PF]`, `[XA:ACCOUNTS]` и т.д.
Нужен для идемпотентности (см. `_strip_previous_run` в writer.py):
при повторном запуске все строки от первого такого маркера вниз чистятся,
и запись делается заново."""

# =============================================================================
# Сохранение
# =============================================================================

OVERWRITE_SOURCE = True     # уточнено: перезаписывать исходный файл
MAKE_BACKUP = True          # но обязательно делать бэкап до записи
BACKUP_SUFFIX_FMT = ".backup_%Y%m%d_%H%M%S"  # итог: file.backup_20260422_140533.xlsx

# =============================================================================
# Маски полей счетов
# =============================================================================

ACCOUNT_PREFIXES_GROUP_A = ("AB", "AN", "AS")
"""Тройка A: BIC + имя учётного узла + код счёта (группа A).
ТЗ: «Ключ счета равен маске поля AB*,AN*,AS*»."""

ACCOUNT_PREFIXES_GROUP_B = ("BB", "BN", "BS")
"""Тройка B: аналогичная тройка полей с префиксом B*.
ТЗ: «или BB*,BN*,BS*»."""

# =============================================================================
# Таблица счетов на выходе (колонки в порядке их появления в итоговой таблице)
# =============================================================================
# Фактически joiner добавляет ПЕРЕД этим списком колонку ACCOUNT_FIELD_KEY_HEADER.
# Формат элемента: (имя_колонки_на_выходе, источник)
# Источники: "SCPF" — из листа SCPF, матчится по (SCAB, SCAN, SCAS)
#            "S5PF" — из листа S5PF, матчится по (S5AB, S5AN, S5AS) == (SCAB, SCAN, SCAS)
#            "COMPUTED" — вычисляемое поле (Excel-формула, не значение)

ACCOUNT_TABLE_COLUMNS = [
    # -------- Источник SCPF (обязательный LEFT JOIN) ----------
    ("SCAB",   "SCPF"),  # ТЗ: часть ключа — компонент BIC
    ("SCAN",   "SCPF"),  # ТЗ: часть ключа — имя учётного узла
    ("SCAS",   "SCPF"),  # ТЗ: часть ключа — код счёта
    ("SCACT",  "SCPF"),  # ТЗ
    ("SCSAC",  "SCPF"),  # ТЗ
    ("SCNANC", "SCPF"),  # ТЗ
    ("SCCCY",  "SCPF"),  # ТЗ: валюта счёта
    ("SCBAL",  "SCPF"),  # ТЗ: остаток
    ("SCSUM0", "SCPF"),  # ТЗ
    ("SCSUMD", "SCPF"),  # ТЗ: сумма дебет
    ("SCSUMC", "SCPF"),  # ТЗ: сумма кредит
    ("SCRBA",  "SCPF"),  # ТЗ

    # -------- Источник S5PF (опциональный LEFT JOIN) -----------
    # ТЗ: «(не все счета есть в S5PF)» — эти три колонки могут быть NaN
    ("S5BAL",  "S5PF"),
    ("S5AIMD", "S5PF"),
    ("S5AM1D", "S5PF"),

    # -------- Вычисляемое поле --------
    # ТЗ: «-(S5AIMD+S5AM1D)» — пишется как Excel-формула
    ("S5_NEG_SUM", "COMPUTED"),
]
```

---

### 5.1. Требования к документации кода (обязательно)

> **Критично для maintenance.** ТЗ живёт в голове бизнес-заказчика, код — в репозитории. Без подробных комментариев через полгода никто не вспомнит, почему `SCAB` матчится с `YW3AB2`, а не с `YW3ABQ`, и что такое `YW2PRZ5`.

**Каждая функция и класс, работающие с полями/таблицами из ТЗ, ДОЛЖНЫ иметь:**

1. **Docstring (Google-style) с обязательными секциями:**
   - Краткое описание (1 строка);
   - Секция `ТЗ:` — дословная цитата соответствующего пункта из ТЗ;
   - `Args:` — с описанием формата данных (не просто тип, но и пример значения);
   - `Returns:` — аналогично;
   - `Example:` — для публичных функций.

2. **Inline-комментарии** рядом с:
   - любой магической константой (`HEADER_ROW = 2  # строка заголовков в листах PF`);
   - каждым полем из списка ТЗ (почему оно здесь и откуда берётся);
   - нетривиальной бизнес-логикой (особенно JOIN-ключами, условными блоками, извлечением троек AB/AN/AS).

3. **Таблицы сопоставления полей** вынесены в `config.py` как константы с комментариями, а не хардкодятся в логике.

**Пример эталонной документации функции:**

```python
def extract_all_accounts(
    sheets: dict[str, list[dict]],
) -> set[tuple[tuple[str, str, str], tuple[str, str, str]]]:
    """Извлекает уникальные счета со всех sheet-источников.

    ТЗ: «На Листах YW2PF, YW3PF и YWJ1PF указаны счета…» Плюс прикладное правило:
    учитываются **только** тройки имён колонок из ``ALLOWED_ACCOUNT_FIELD_TRIPLES``.

    Алгоритм:
        1. Для каждого листа из ``ACCOUNT_SOURCE_SHEETS`` по ``SHEET_FIELD_PREFIX``
           находим слоты с полными тройками AB/AN/AS или BB/BN/BS (regex: суффикс
           слота после буквенной группы допускает быть пустым — для колонок вида ``YWJ1AB``).
        2. Слот без полной тройки отбрасывается (защита от полей вроде ``YW3ANR``).
        3. Если ``(c_AB, c_AN, c_AS)`` не входит в allowlist — слот пропускается.
        4. Из ячеек читаются значения; в множество добавляется пара
           ``((c_AB, c_AN, c_AS), (SCAB, SCAN, SCAS))`` при непустой тройке значений.
        5. Группы A и B сливаются в одно множество: JOIN с SCPF/S5PF всегда по значениям SC*.

    Returns:
        Множество ``((имена колонок), (значения для JOIN))``. Пример элемента:
        ``(("YW3AB2", "YW3AN2", "YW3AS2"), ("0880", "AUDE75", "006"))``.

    Example:
        >>> sheets = {"YW3PF": [{"YW3AB2": "0880", "YW3AN2": "AUDE75",
        ...                      "YW3AS2": "006", "YW3ANR": "ignored"}]}
        >>> extract_all_accounts(sheets)
        {(('YW3AB2', 'YW3AN2', 'YW3AS2'), ('0880', 'AUDE75', '006'))}
    """
    ...
```

**Запрещено:** «говорящие» имена функций/переменных без docstring. Имя `extract_all_accounts` не заменяет объяснения логики извлечения.

---

## 6. Алгоритм обработки (пайплайн)

### 6.1. Высокоуровневый порядок

```
load_workbook(path)
 └─► read_all_sheets() ──► dict[sheet_name, list[dict]]  # строки как dict по заголовкам
       │
       ├─► extract_all_accounts(sheets=[YW2PF, YW3PF, YWJ1PF])
       │      ──► set[((c_AB,c_AN,c_AS), (SCAB,SCAN,SCAS))]   # только allowlist
       │
       ├─► determine_conditional_sheets(yw2pf_row)         ──► list[sheet_name]
       │
       ├─► build_account_table(accounts, SCPF, S5PF)        ──► pandas.DataFrame
       │      (сортировка строк по ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED;
       │       колонка «Ключ полей PF» = имена полей через «-»)
       │
       └─► write_to_yw2pf(wb, blocks, account_table, …)
             ├─► append YW3PF (header+data)
             ├─► append YWJ1PF (header+data)
             ├─► append AN6PF  (conditional)
             ├─► append AN9PF  (conditional)
             ├─► append account table  (header + rows + formula column)
             └─► apply AutoFilter to account table range

save_workbook(new_path)
```

### 6.2. Чтение листа (`sheet_reader.py`)

> **Важно: количество data-строк на листах-источниках — переменное.**
> Для **`YWJ1PF` особенно** в примере их 2, но в реальных файлах может быть 0, 1, 3 или больше.
> То же справедливо для `YW3PF`, `AN6PF`, `AN9PF`. Код **НЕ** должен опираться на конкретное число строк — только на маркер конца данных (пустая строка или `ws.max_row`).

```python
def read_sheet_as_dicts(wb, sheet_name: str) -> list[dict]:
    """Возвращает все data-строки листа в виде списка словарей.

    ТЗ: листы-источники могут содержать произвольное число записей.
    В частности, на YWJ1PF их может быть 1, 2 или больше. Функция
    не делает предположений о количестве — читает всё, что есть.

    Формат листа (общий для всех листов PF в файле):
        строка 1: служебная ссылка «Go to Set Sheet» — игнорируется;
        строка 2: имена полей (заголовки);
        строка 3..N: данные.

    Args:
        wb: openpyxl Workbook.
        sheet_name: Имя листа, напр. "YWJ1PF".

    Returns:
        Список словарей вида [{field_name: value, ...}, ...], по одному
        на каждую непустую data-строку. Пустая строка (все ячейки None
        или пустые) трактуется как конец данных и пропускается (не
        прерывает чтение — просто фильтруется).

    Example:
        >>> read_sheet_as_dicts(wb, "YWJ1PF")
        [{"YWJ1ANR": "F0ICRG20S210", "YWJ1OTP": "1", ...},
         {"YWJ1ANR": "F0ICRG20S210", "YWJ1OTP": "2", ...}]
    """
    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[HEADER_ROW]]
    rows = []
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        # Пропуск полностью пустых строк (но НЕ обрыв чтения — возможны
        # «островки» пустых строк в середине листа, хотя в примере их нет)
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        rows.append(dict(zip(headers, row)))
    return rows
```

> Всегда открываем книгу через `load_workbook(path, data_only=False)`, чтобы сохранять формулы в ранее записанных местах. Для **чтения вычисленных значений** (например, если `YW2PR2` — это формула) используем вторую временную копию с `data_only=True`.

### 6.3. Извлечение счетов (`account_extractor.py`)

**Ключевая логика:**

1. Найти слоты, где на листе присутствуют **все три** поля тройки (AB+AN+AS или BB+BN+BS) — иначе поля вроде `YW3ANR` не образуют ложный счёт.
2. Регулярное выражение для заголовка: `^{prefix}(AB|AN|AS|BB|BN|BS)(.*)$` — суффикс после буквенной группы может быть **пустым** (колонки `YWJ1AB`, `YWJ1AN`, `YWJ1AS`).
3. Учитывать только тройки имён колонок из **`ALLOWED_ACCOUNT_FIELD_TRIPLES`** (множество строится из **`ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED`** в `config.py`).
4. В результат попадают пары **(тройка имён колонок, тройка значений ячеек)**; пустые тройки значений отбрасываются.

```python
import re

def extract_account_slots(headers, sheet_prefix: str):
    slots = {}
    pat = re.compile(rf"^{re.escape(sheet_prefix)}(AB|AN|AS|BB|BN|BS)(.*)$")
    for h in headers:
        if h is None:
            continue
        m = pat.match(str(h))
        if not m:
            continue
        kind, slot = m.group(1), m.group(2)
        slots.setdefault(slot, {})[kind] = h
    result = []
    for slot, kinds in slots.items():
        if {"AB", "AN", "AS"} <= kinds.keys():
            result.append((slot, "A", (kinds["AB"], kinds["AN"], kinds["AS"])))
        if {"BB", "BN", "BS"} <= kinds.keys():
            result.append((slot, "B", (kinds["BB"], kinds["BN"], kinds["BS"])))
    return result


def extract_all_accounts(sheets):
    """Итерация по ACCOUNT_SOURCE_SHEETS и SHEET_FIELD_PREFIX — см. реальный модуль."""
    accounts = set()
    for sheet_name in ACCOUNT_SOURCE_SHEETS:
        # ...
        for row in rows:
            for _slot_id, _group, (c1, c2, c3) in slots:
                if (c1, c2, c3) not in ALLOWED_ACCOUNT_FIELD_TRIPLES:
                    continue
                triple = (_norm(row.get(c1)), _norm(row.get(c2)), _norm(row.get(c3)))
                if any(triple):
                    accounts.add(((c1, c2, c3), triple))
    return accounts
```

**Важно.** Значения из групп A и B сопоставляются с `SCPF`/`S5PF` по одному и тому же ключу `SCAB/SCAN/SCAS`. Разные тройки **имён** колонок с одинаковыми **значениями** дают **разные строки** в таблице на выходе (разный «Ключ полей PF»).

### 6.4. Условное добавление AN6PF / AN9PF

Читаем значение полей `YW2PR2` и `YW2PR5` из **первой** data-строки `YW2PF` (`row index = 3`). Если поле непустое (`not _is_blank`), помечаем соответствующий лист к копированию.

```python
def _is_blank(v) -> bool:
    return v is None or str(v).strip() == ""
```

### 6.5. Построение таблицы счетов (`joiner.py`)

```python
import pandas as pd

def build_account_table(
    accounts: set[tuple[tuple[str, str, str], tuple[str, str, str]]],
    sc_rows: list[dict],
    s5_rows: list[dict],
) -> pd.DataFrame:
    """LEFT JOIN SCPF + LEFT JOIN S5PF по (SCAB, SCAN, SCAS).

    Отличия от базового ТЗ в реализации:
        - Вход: ``accounts`` — множество ``((c_AB, c_AN, c_AS), (SCAB, SCAN, SCAS))``.
        - Строки сортируются по индексу тройки колонок в
          ``ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED``.
        - В ``acc_df`` добавляется колонка ``ACCOUNT_FIELD_KEY_HEADER`` («Ключ полей PF»):
          ``f"{c1}-{c2}-{c3}"`` для наглядности в Excel.
        - Итоговый порядок колонок: «Ключ полей PF», SCAB…S5AM1D (как в коде).
        - Колонка ``-(S5AIMD+S5AM1D)`` по-прежнему пишется в ``writer.py`` формулой.

    Алгоритм (сжато):
        1. Отсортировать ``accounts`` по порядку в ``ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED``.
        2. Построить ``acc_df`` с колонками ключ + SCAB, SCAN, SCAS.
        3. Нормализовать ключи (строка, ``.strip()``) в sc_df, s5_df, acc_df.
        4. ``drop_duplicates`` по ключу в справочниках; предупреждение, если счёта нет в SCPF.
        5. Два последовательных ``merge(..., how="left")``; финальная перестановка колонок.
    """
    ...
```

Реальные имена колонок SCPF/S5PF и список полей merge — в ``SCPF_MERGE_COLUMNS`` / ``S5PF_MERGE_COLUMNS`` (`config.py`).

### 6.6. Запись в YW2PF (`writer.py`)

**Принципы:**
- Не трогаем существующие строки (формулы, формат сохраняются).
- **Идемпотентность (критично при перезаписи исходника):** перед записью вызываем `_strip_previous_run(ws)` — удаляем все строки от первого `[XA:*]` маркера вниз. Это позволяет безопасно запускать обработку повторно: результаты прошлого запуска вычищаются, накатывается актуальный.
- Находим `last_data_row` через перебор с конца до первой непустой строки (после очистки).
- Каждый блок: 1 пустая строка → маркер `[XA:YW3PF]` жирным в колонке A → строка заголовков → data-строки.
- Для столбца `-(S5AIMD+S5AM1D)` пишем **Excel-формулу**, не посчитанное значение (ссылки на ячейки той же строки). Индексы колонок `S5AIMD` и `S5AM1D` берутся из `account_df.columns` (после добавления «Ключ полей PF» сдвиг относительно фиксированных 14/15).
- К диапазону таблицы счетов применяем автофильтр от первой колонки до последней **включая** столбец с формулой (в коде — `get_column_letter` по фактическому числу колонок).

```python
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

def _strip_previous_run(ws) -> None:
    """Удаляет все строки от первого маркера [XA:*] в колонке A вниз.
    Гарантирует идемпотентность при повторной обработке того же файла."""
    first_marker_row = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if isinstance(cell.value, str) and cell.value.startswith(BLOCK_MARKER_PREFIX):
            first_marker_row = cell.row
            break
    if first_marker_row is None:
        return
    # удаляем строки пачкой (от first_marker_row-1 — сначала пустую-отступ, потом всё ниже)
    ws.delete_rows(first_marker_row - 1, ws.max_row - first_marker_row + 2)
    # снимаем старый автофильтр, если был
    ws.auto_filter.ref = None


def _find_last_nonempty_row(ws) -> int:
    for r in range(ws.max_row, 0, -1):
        if any(ws.cell(row=r, column=c).value not in (None, "")
               for c in range(1, ws.max_column + 1)):
            return r
    return 0


def write_blocks_to_yw2pf(wb, ordered_blocks: list, account_df):
    """
    ordered_blocks = [
        ("YW3PF",  headers_list, rows_list_of_dicts),
        ("YWJ1PF", headers_list, rows_list_of_dicts),
        # ... условные
    ]
    """
    ws = wb[TARGET_SHEET]

    # 1) чистим хвост прошлого запуска (идемпотентность)
    _strip_previous_run(ws)

    cursor = _find_last_nonempty_row(ws) + 1 + BLOCK_GAP   # пустая строка-отступ

    for block_name, headers, rows in ordered_blocks:
        # маркер блока — по нему при повторном запуске находится точка очистки
        ws.cell(row=cursor, column=1,
                value=f"{BLOCK_MARKER_PREFIX}{block_name}{BLOCK_MARKER_SUFFIX}").font = Font(bold=True)
        cursor += 1
        # заголовки (имена полей)
        for ci, h in enumerate(headers, start=1):
            ws.cell(row=cursor, column=ci, value=h).font = Font(bold=True)
        cursor += 1
        # строки данных
        for row_dict in rows:
            for ci, h in enumerate(headers, start=1):
                ws.cell(row=cursor, column=ci, value=row_dict.get(h))
            cursor += 1
        cursor += BLOCK_GAP   # зазор перед следующим блоком

    # Таблица счетов
    ws.cell(row=cursor, column=1,
            value=f"{BLOCK_MARKER_PREFIX}ACCOUNTS{BLOCK_MARKER_SUFFIX}").font = Font(bold=True)
    cursor += 1

    acc_headers = list(account_df.columns) + ["-(S5AIMD+S5AM1D)"]
    header_row = cursor
    for ci, h in enumerate(acc_headers, start=1):
        ws.cell(row=cursor, column=ci, value=h).font = Font(bold=True)
    cursor += 1

    data_start = cursor
    # Индексы S5AIMD / S5AM1D — по имени в account_df (не хардкод 14/15).
    s5aimd_col_letter = get_column_letter(list(account_df.columns).index("S5AIMD") + 1)
    s5am1d_col_letter = get_column_letter(list(account_df.columns).index("S5AM1D") + 1)
    neg_col_idx = len(account_df.columns) + 1

    for _, rec in account_df.iterrows():
        for ci, col in enumerate(account_df.columns, start=1):
            val = rec[col]
            ws.cell(row=cursor, column=ci, value=val if pd.notna(val) else None)
        ws.cell(row=cursor, column=neg_col_idx,
                value=f"=-({s5aimd_col_letter}{cursor}+{s5am1d_col_letter}{cursor})")
        cursor += 1
    data_end = cursor - 1

    # Автофильтр на диапазон заголовков + данных таблицы счетов
    start_col = get_column_letter(1)
    end_col   = get_column_letter(neg_col_idx)
    ws.auto_filter.ref = f"{start_col}{header_row}:{end_col}{data_end}"
```

### 6.7. Пересчёт формул

После `wb.save(path)` добавленная формула `=-(N5+O5)` будет записана как строка, но без кэшированного значения. **Решение:**
- Либо при открытии файла в Excel пересчёт сработает автоматически (флаг `wb.calculation.calcMode = 'auto'`);
- Либо запускаем LibreOffice в headless-режиме для пересчёта (скрипт `scripts/recalc.py` из xlsx-skill).

Для конечного пользователя (открывает в Excel/Numbers) достаточно первого варианта. Зафиксировать: `wb.properties.calcMode = 'auto'` и `wb.calculation.fullCalcOnLoad = True`.

### 6.8. Безопасная перезапись исходника (`pipeline.py`)

Перезапись оригинального файла — быстрый, но рисковый режим. Минимизируем потери тремя слоями защиты:

**1. Бэкап до любых операций.** В начале pipeline копируем исходник рядом с ним:
```
<name>.xlsx  →  <name>.backup_20260422_140533.xlsx
```

**2. Атомарная запись через временный файл.** Сохраняем обработанный workbook не поверх оригинала, а в соседний `.tmp`, и только после успешного `wb.save()` делаем `os.replace(tmp, original)`. Это защищает от «битого» файла при падении процесса посреди записи (электричество, kill).

**3. Проверка на lock перед стартом.** На Windows Excel держит открытый файл в эксклюзивном режиме. Пробуем открыть `original` на запись в начале — при `PermissionError` показываем диалог «Закройте файл в Excel и повторите» и не трогаем ничего.

```python
import shutil, os, tempfile
from datetime import datetime
from pathlib import Path

def safe_overwrite_save(wb, original_path: str) -> dict:
    """Сохраняет wb поверх original_path с бэкапом и атомарной заменой.
    Возвращает dict с путями: {'result': ..., 'backup': ...}."""
    orig = Path(original_path)

    # 1) Проверка на lock
    try:
        with open(orig, "ab"):  # открытие в режиме append не меняет содержимое
            pass
    except PermissionError:
        raise RuntimeError(
            f"Файл занят другим процессом (вероятно открыт в Excel): {orig}. "
            f"Закройте файл и повторите."
        )

    # 2) Бэкап
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = orig.with_name(f"{orig.stem}.backup_{ts}{orig.suffix}")
    shutil.copy2(orig, backup)

    # 3) Запись в .tmp в той же директории (важно для os.replace на одной FS)
    with tempfile.NamedTemporaryFile(
        dir=orig.parent, prefix=f".{orig.stem}.", suffix=".tmp", delete=False
    ) as tf:
        tmp_path = Path(tf.name)
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, orig)  # атомарная замена
    except Exception:
        # откат: временный файл удаляем, оригинал не тронут (бэкап всё равно остался)
        if tmp_path.exists():
            tmp_path.unlink()
        raise

    return {"result": str(orig), "backup": str(backup)}
```

**Поведение UI по завершении:**
- Диалог «Готово. Файл обработан. Бэкап: `...backup_20260422_140533.xlsx`» + кнопка «Открыть папку».
- Политика хранения бэкапов: по умолчанию храним все. Опционально (настройка-галка в UI на будущее) — автоудаление бэкапов старше N дней.

---

## 7. Обработка ошибок и логирование

| Ситуация | Поведение |
|---|---|
| Файл не `.xlsx` | Диалог ошибки, выход |
| Отсутствует `YW2PF` / `SCPF` | Ошибка, список недостающих листов в диалоге, исходник не тронут |
| Отсутствует `S5PF` | Пропускаем S5-колонки, пишем NaN, лог WARNING |
| Отсутствует `YW3PF`/`YWJ1PF`/`AN6PF`/`AN9PF` | Лог WARNING, блок пропускается |
| Нет ни одного счёта | Блок `[XA:ACCOUNTS]` не пишется, лог INFO |
| **Исходный файл открыт в Excel** | `PermissionError` ловим **до любых операций** (см. 6.8) → диалог «Файл занят — закройте его в Excel». Никакой бэкап не создаётся. |
| Ошибка посреди записи | Временный `.tmp` удаляется. Оригинал не тронут (т.к. `os.replace` ещё не отработал). Бэкап уже существует как дополнительная страховка. |
| Любая другая ошибка | Полный traceback → `%LOCALAPPDATA%\xlsx_aggregator\logs\error.log`, в UI — короткое сообщение + кнопка «Открыть лог» + путь к бэкапу (если уже создан) |

Логгер — `loguru`, 2 синка: stderr (для dev) + файл с ротацией 5 МБ.

---

## 8. Зафиксированные решения и оставшиеся вопросы

### 8.1. Зафиксировано по итогам уточнений

| № | Решение |
|---|---|
| 1 | Триггер AN9PF = поле `YW2PRZ5` (в ТЗ была опечатка `YW2PR5`). |
| 2 | **Одна пустая строка перед каждым** новым блоком. Константа `BLOCK_GAP = 1`. |
| 3 | **Перезапись исходника**. Перед записью автоматический бэкап `*.backup_<timestamp>.xlsx` + атомарная запись через `.tmp` + `os.replace` (см. 6.8). |
| 4 | **Идемпотентность при повторных запусках:** блоки маркируются `[XA:<имя>]` в колонке A. Перед записью все строки от первого маркера и ниже чистятся (см. `_strip_previous_run` в 6.6). |
| 5 | **Счета с PF-листов:** извлекаются только тройки колонок из `ALLOWED_ACCOUNT_FIELD_TRIPLES`; порядок строк в таблице ACCOUNTS = `ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED`; первая колонка таблицы — «Ключ полей PF» (`имя1-имя2-имя3`). |

### 8.2. Мелкие вопросы (не блокируют старт — закладываю дефолт, скажи если не подходит)

1. **Маркеры блоков `[XA:YW3PF]`, `[XA:ACCOUNTS]` жирным в колонке A.** Нужны для идемпотентности и ориентира. Убрать визуально (оставить, но белым цветом) — могу, просто скажи.
2. **Дубликаты.** Одна и та же пара `(тройка имён колонок, тройка значений)` со всех строк и листов сворачивается в один элемент множества → одна строка в таблице. Совпадающие **значения** при разных **именах** колонок (две строки allowlist) → две строки с разным «Ключ полей PF». OK?
3. **Сравнение ключей — строгое, строковое, с `.strip()`.** Ведущие нули (`'001'`, `'002'`) сохраняются. Регистр важен. OK?
4. **Автофильтр** накладываю **только** на таблицу `[XA:ACCOUNTS]`. В блоках YW3/YWJ1/AN6/AN9 по 1–4 строки — фильтр не нужен.
5. **Если блок счетов пуст** (ни одного счёта не извлеклось) — пропускаю, записываю только остальные блоки. Лог INFO.
6. **Язык UI** — русский.

---

## 9. `requirements.txt`

```
openpyxl==3.1.5
pandas==2.2.3
customtkinter==5.2.2
loguru==0.7.2
pyinstaller==6.11.1
pytest==8.3.3
```

---

## 10. Команды запуска и сборки (Windows)

### Запуск из исходников (dev)

`run.bat`:
```bat
@echo off
if not exist .venv (
    py -3.11 -m venv .venv
    call .venv\Scripts\activate
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate
)
python -m src.main
```

### Сборка `.exe` (`build.bat`)

```bat
@echo off
call .venv\Scripts\activate
pyinstaller --noconfirm --windowed --name "XLSX Aggregator" ^
  --icon assets\icon.ico ^
  --add-data "src\core\config.py;core" ^
  --collect-all customtkinter ^
  src\main.py
echo.
echo Result: dist\XLSX Aggregator\XLSX Aggregator.exe
```

**Рекомендуемый режим сборки:** one-folder (как выше), **не** one-file.
One-file запускается медленнее (распаковка во временную папку), и Windows Defender чаще реагирует на него как на подозрительный.

**Что делать с false-positive от антивируса:**
- На dev-машине: добавить папку `dist\` в исключения Windows Defender.
- Для раздачи коллегам: либо подписать `.exe` code-signing сертификатом, либо приложить к артефакту инструкцию «Добавьте в исключения».

### Тесты

```bat
call .venv\Scripts\activate
pytest -v tests/
```

### Путь логов

```
%LOCALAPPDATA%\xlsx_aggregator\logs\app.log
```

Развернётся примерно в `C:\Users\<User>\AppData\Local\xlsx_aggregator\logs\app.log`.

---

## 11. Пошаговый план для Cursor

Разбил на атомарные задачи. Каждая — отдельный коммит.

| # | Шаг | Файл(ы) | Что готово по итогу |
|---|---|---|---|
| 1 | Скелет проекта + `requirements.txt` + `README.md` | все корневые | `pip install -r requirements.txt` проходит |
| 2 | `config.py` — все константы (имена листов, поля, префиксы) | `src/core/config.py` | Константы отделены от логики |
| 3 | `logging_setup.py` — loguru с файловым синком | `src/utils/logging_setup.py` | `logger.info(...)` пишет и в stderr, и в файл |
| 4 | `sheet_reader.py` + тест | `src/core/sheet_reader.py`, `tests/test_sheet_reader.py` | На тестовом файле возвращает корректный `list[dict]` |
| 5 | `account_extractor.py` + тесты (включая `YW3ANR`, allowlist, `YWJ1AB` без суффикса; YWJ1PF с 0/1/N data-строками) | `src/core/account_extractor.py` | Учитываются только allowlist-тройки; лишние колонки не попадают в счета |
| 6 | `joiner.py` + тест на LEFT JOIN (S5 отсутствует — NaN), колонка «Ключ полей PF», сортировка | `src/core/joiner.py` | DataFrame: ключ + SC*…S5* в порядке `ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED` |
| 7 | `writer.py` — запись блоков с маркерами `[XA:*]`, формулой и автофильтром + `_strip_previous_run` для идемпотентности | `src/core/writer.py` | Повторный запуск на уже обработанном файле даёт тот же результат, а не удваивает блоки |
| 8 | `pipeline.py` — оркестрация + `safe_overwrite_save` (бэкап + lock-check + атомарная запись через `.tmp`) + колбэк прогресса (%) | `src/core/pipeline.py` | Из CLI (`python -m src.core.pipeline file.xlsx`) работает без GUI; при открытом в Excel файле даёт понятную ошибку, файл не трогается |
| 9 | `gui/app.py` — customtkinter, 2 кнопки, прогресс, лог-вьюшка, поток, диалог подтверждения перезаписи | `src/gui/app.py` | GUI запускает pipeline без фриза окна; по завершении показывает путь к бэкапу |
| 10 | `main.py` — точка входа, подъём GUI | `src/main.py` | `python -m src.main` открывает окно |
| 11 | Интеграционный тест на примере `Пример_файла.xlsx` + синтетический кейс с YWJ1PF, содержащим 0 и 5 записей | `tests/test_pipeline.py` | End-to-end проверка: все блоки на месте, автофильтр применён, переменное число записей обработано корректно |
| 12 | Сборка `.exe` через PyInstaller (`build.bat`, one-folder) | `build.bat` | `dist\XLSX Aggregator\XLSX Aggregator.exe` запускается двойным кликом на чистой Windows без установленного Python |
| 13 | README.md: скриншоты, how-to, примечание про Windows Defender и путь логов `%LOCALAPPDATA%\xlsx_aggregator\logs\` | `README.md` | Новый пользователь может всё поставить и собрать |

---

## 12. Чек-лист приёмки

- [ ] Пользователь выбирает файл и жмёт «Обработать» — через ≤ 10 с получает обработанный файл (перезаписанный).
- [ ] Перед записью автоматически создан бэкап `<имя>.backup_<timestamp>.xlsx` рядом с исходником.
- [ ] На `YW2PF` ниже исходных данных (через 1 пустую строку перед каждым блоком) появились блоки `YW3PF`, `YWJ1PF`, опционально `AN6PF` (если `YW2PR2 ≠ ""`), опционально `AN9PF` (если `YW2PRZ5 ≠ ""`), и таблица счетов.
- [ ] Каждый блок помечен маркером `[XA:<имя>]` в колонке A.
- [ ] Таблица счетов начинается с колонки «Ключ полей PF»; далее колонки в порядке ТЗ; строки в порядке `ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED`; вычисляемый столбец `-(S5AIMD+S5AM1D)` — Excel-формула.
- [ ] На диапазон таблицы счетов наложен автофильтр.
- [ ] **Идемпотентность:** повторный запуск на том же файле даёт тот же результат (блоки не удваиваются).
- [ ] Если файл открыт в Excel — приложение выдаёт понятную ошибку и НЕ портит файл и НЕ создаёт бэкап.
- [ ] Все формулы в получившемся файле без ошибок (`#REF!`, `#VALUE!` и т. п.).
- [ ] Приложение собирается в `.exe` (PyInstaller one-folder) и запускается на чистой Windows без установленного Python.
- [ ] Логи сохраняются в `%LOCALAPPDATA%\xlsx_aggregator\logs\`.
- [ ] Обработан кейс с переменным числом data-строк на YWJ1PF (проверено тестом с 0, 1 и N записями).
