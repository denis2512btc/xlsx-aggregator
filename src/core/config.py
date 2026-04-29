"""Центральная точка всех бизнес-констант из ТЗ.

Если бизнес-заказчик меняет название поля-триггера, добавляет новый лист
или новое поле в таблицу счетов — правится ТОЛЬКО этот файл.
"""

from __future__ import annotations

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

SC_SHEET = "SCPF"  # обязательный источник для таблицы счетов
S5_SHEET = "S5PF"  # опциональный источник (ТЗ: «не все счета есть в S5PF»)

# Префиксы имён полей по имени листа (строка 2 — заголовки)
SHEET_FIELD_PREFIX: dict[str, str] = {
    "YW2PF": "YW2",
    "YW3PF": "YW3",
    "YWJ1PF": "YWJ1",
}

ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED: tuple[tuple[str, str, str], ...] = (
    ("YW3AB2", "YW3AN2", "YW3AS2"),
    ("YWJ1AB", "YWJ1AN", "YWJ1AS"),
    ("YW3AB4", "YW3AN4", "YW3AS4"),
    ("YW3AB3", "YW3AN3", "YW3AS3"),
    ("YW3AB5", "YW3AN5", "YW3AS5"),
    ("YWJ1AB2", "YWJ1AN2", "YWJ1AS2"),
    ("YW3AB8", "YW3AN8", "YW3AS8"),
    ("YW3ABL", "YW3ANL", "YW3ASL"),
    ("YW3ABO", "YW3ANO", "YW3ASO"),
    ("YW3BB4", "YW3BN4", "YW3BS4"),
    ("YW3BB5", "YW3BN5", "YW3BS5"),
    ("YW3BB6", "YW3BN6", "YW3BS6"),
)
"""Порядок троек полей: строки таблицы счетов на выходе сортируются так же."""

ALLOWED_ACCOUNT_FIELD_TRIPLES: frozenset[tuple[str, str, str]] = frozenset(
    ALLOWED_ACCOUNT_FIELD_TRIPLES_ORDERED
)
"""Только эти тройки полей участвуют в извлечении ключей счетов с PF-листов."""

ACCOUNT_FIELD_KEY_HEADER = "Ключ полей PF"
"""Заголовок колонки: имена полей AB/AN/AS (или BB/BN/BS) через дефис, напр. ``YW3AB2-YW3AN2-YW3AS2``."""

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

ACCOUNTS_BLOCK_NAME = "ACCOUNTS"

# =============================================================================
# Сохранение
# =============================================================================

OVERWRITE_SOURCE = True  # уточнено: перезаписывать исходный файл
MAKE_BACKUP = True  # но обязательно делать бэкап до записи
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
# Формат элемента: (имя_колонки_на_выходе, источник)
# Источники: "SCPF" — из листа SCPF, матчится по (SCAB, SCAN, SCAS)
#            "S5PF" — из листа S5PF, матчится по (S5AB, S5AN, S5AS) == (SCAB, SCAN, SCAS)
#            "COMPUTED" — вычисляемое поле (Excel-формула, не значение)

ACCOUNT_TABLE_COLUMNS = [
    # -------- Источник SCPF (обязательный LEFT JOIN) ----------
    ("SCAB", "SCPF"),  # ТЗ: часть ключа — компонент BIC
    ("SCAN", "SCPF"),  # ТЗ: часть ключа — имя учётного узла
    ("SCAS", "SCPF"),  # ТЗ: часть ключа — код счёта
    ("SCACT", "SCPF"),  # ТЗ
    ("SCSAC", "SCPF"),  # ТЗ
    ("SCNANC", "SCPF"),  # ТЗ
    ("SCCCY", "SCPF"),  # ТЗ: валюта счёта
    ("SCBAL", "SCPF"),  # ТЗ: остаток
    ("SCSUM0", "SCPF"),  # ТЗ
    ("SCSUMD", "SCPF"),  # ТЗ: сумма дебет
    ("SCSUMC", "SCPF"),  # ТЗ: сумма кредит
    ("SCRBA", "SCPF"),  # ТЗ
    # -------- Источник S5PF (опциональный LEFT JOIN) -----------
    # ТЗ: «(не все счета есть в S5PF)» — эти три колонки могут быть NaN
    ("S5BAL", "S5PF"),
    ("S5AIMD", "S5PF"),
    ("S5AM1D", "S5PF"),
    # -------- Вычисляемое поле --------
    # ТЗ: «-(S5AIMD+S5AM1D)» — пишется как Excel-формула
    ("S5_NEG_SUM", "COMPUTED"),
]

# Имя последнего столбца в Excel (заголовок), формула в writer
ACCOUNT_COMPUTED_HEADER = "-(S5AIMD+S5AM1D)"

# Колонки SCPF, участвующие в JOIN и переносе (ключ + данные)
SCPF_MERGE_COLUMNS = [
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
]

# Колонки S5PF для merge (ключ переименовывается в joiner)
S5PF_MERGE_COLUMNS = ["S5AB", "S5AN", "S5AS", "S5BAL", "S5AIMD", "S5AM1D"]

# =============================================================================
# YQ-режим: константы
# =============================================================================

TARGET_SHEET_YQ = "YQ2PF"

ALWAYS_APPEND_YQ = ["YQ3PF"]
"""YWJ1PF в YQ-режиме не используется."""

ACCOUNT_SOURCE_SHEETS_YQ = ["YQ2PF", "YQ3PF"]
SHEET_FIELD_PREFIX_YQ: dict[str, str] = {"YQ2PF": "YQ2", "YQ3PF": "YQ3"}

ALLOWED_ACCOUNT_FIELD_TRIPLES_YQ_ORDERED: tuple[tuple[str, str, str], ...] = (
    ("YQ3AB2", "YQ3AN2", "YQ3AS2"),
    ("YQ3AB3", "YQ3AN3", "YQ3AS3"),
    ("YQ3AB4", "YQ3AN4", "YQ3AS4"),
    ("YQ3AB5", "YQ3AN5", "YQ3AS5"),
    ("YQ3AB6", "YQ3AN6", "YQ3AS6"),
    ("YQ3AB7", "YQ3AN7", "YQ3AS7"),
    ("YQ3AB8", "YQ3AN8", "YQ3AS8"),
    ("YQ3ABG", "YQ3ANG", "YQ3ASG"),
    ("YQ3ABP", "YQ3ANP", "YQ3ASP"),
    ("YQ3ABN", "YQ3ANN", "YQ3ASN"),
    ("YQ3ABK", "YQ3ANK", "YQ3ASK"),
    ("YQ3ABL", "YQ3ANL", "YQ3ASL"),
    ("YQ3ABM", "YQ3ANM", "YQ3ASM"),
    ("YQ3ABO", "YQ3ANO", "YQ3ASO"),
    ("YQ3BB2", "YQ3BN2", "YQ3BS2"),
    ("YQ3BB3", "YQ3BN3", "YQ3BS3"),
    ("YQ2AB1", "YQ2AN1", "YQ2AS1"),
)
"""Порядок троек YQ-режима: строки ACCOUNTS сортируются так же."""

ALLOWED_ACCOUNT_FIELD_TRIPLES_YQ: frozenset[tuple[str, str, str]] = frozenset(
    ALLOWED_ACCOUNT_FIELD_TRIPLES_YQ_ORDERED
)
"""Allowlist для YQ-режима (аналог ALLOWED_ACCOUNT_FIELD_TRIPLES для YW)."""

YQ_CONDITIONAL_TRIGGER = "YQ2PR2"
"""ТЗ: если значение цифровое → AN4PF; если символьное → AN6PF; пусто → никакой лист."""

YQ_CONDITIONAL_NUMERIC_SHEET = "AN4PF"
YQ_CONDITIONAL_STRING_SHEET = "AN6PF"

YQJDATA_BASE_SHEET = "YQJPF"
YQJDATA_OPT_SHEET1 = "YQJOPF"   # LEFT JOIN по (YQJOANR=YQJANR, YQJOSQN=YQJSQN)
YQJDATA_OPT_SHEET2 = "AN41PF"   # LEFT JOIN по (AN41ANR=YQJANR, AN41SQN=YQJSQN)

YQJDATA_BASE_ANR_COL = "YQJANR"
YQJDATA_BASE_SQN_COL = "YQJSQN"
YQJDATA_OPT1_ANR_COL = "YQJOANR"
YQJDATA_OPT1_SQN_COL = "YQJOSQN"
YQJDATA_OPT2_ANR_COL = "AN41ANR"
YQJDATA_OPT2_SQN_COL = "AN41SQN"

YQJDATA_FILTER_FIELD = "YQJSTS"
YQJDATA_FILTER_VALUE = "А"  # Кириллическая А (не латинская)
