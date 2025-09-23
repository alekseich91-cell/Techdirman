"""
Назначение:
    Общие константы и утилиты UI для TechDirRentMan.

Как работает:
    - Определяет пути проекта (APP_ROOT/DATA_DIR/ASSETS_DIR).
    - Словари соответствий классов RU↔EN.
    - Константы внешнего вида (порог переноса).
    - Форматирование/парсинг чисел (fmt_num, fmt_sign, to_float).
    - Настройка автоширин таблиц Qt (setup_auto_col_resize, setup_priority_name, apply_auto_col_resize).

Стиль:
    - Нумерованные секции, краткие комментарии в ключевых местах.
"""

# 1. Импорт библиотек
from pathlib import Path  # пути проекта
from typing import Any     # типы для аннотаций
from PySide6 import QtWidgets  # для настроек таблиц

# 2. Пути проекта
APP_ROOT = Path(__file__).resolve().parents[2]  # корень приложения (папка TechDirRentMan)
DATA_DIR = APP_ROOT / "data"                    # папка данных проекта
ASSETS_DIR = APP_ROOT / "assets"                # папка ассетов (картинки проекта)

# 3. Классы RU↔EN и константы внешнего вида
CLASS_RU2EN = {
    "Оборудование": "equipment",
    "Персонал": "personnel",
    "Логистика": "logistic",
    "Расходник": "consumable",
}
CLASS_EN2RU = {v: k for k, v in CLASS_RU2EN.items()}

WRAP_THRESHOLD = 40  # порог символов для переноса текста в «Наименовании»

# 4. Утилиты форматирования чисел
def fmt_num(value: Any, decimals: int = 2) -> str:
    """Строка числа без лишних нулей: '10', '10,5', '10,25'. Десятичный разделитель — запятая."""
    try:
        v = float(value)
    except Exception:
        return "0"
    s = f"{v:.{decimals}f}".replace('.', ',')
    s = s.rstrip('0').rstrip(',')
    if s == "-0":
        s = "0"
    return s

def fmt_sign(value: Any, decimals: int = 2) -> str:
    """То же, но со знаком для положительных значений."""
    try:
        v = float(value)
    except Exception:
        return "0"
    sign = "+" if v > 0 else ""
    return f"{sign}{fmt_num(v, decimals)}"

def to_float(text: Any, default: float = 0.0) -> float:
    """Надёжный парсер чисел из строк с пробелами и запятыми."""
    if isinstance(text, (int, float)):
        return float(text)
    try:
        return float(str(text).strip().replace(" ", "").replace(",", "."))
    except Exception:
        return float(default)

# 5. Нормализация строк
def normalize_case(text: Any) -> str:
    """Нормализует регистр строковых значений.

    Эта функция предназначена для приведения наименований,
    подрядчиков, отделов и зон к единому виду. Она удаляет
    начальные/конечные пробелы и приводит каждое слово к
    формату с заглавной буквой, что помогает избегать дублирования
    записей, отличающихся только регистром.

    :param text: входная строка или объект, который может быть приведён к строке;
    :return: нормализованная строка (пример: «подрядчик а» → «Подрядчик А»)
    """
    try:
        s = str(text).strip()
    except Exception:
        return ""
    # Для пустой строки ничего не меняем
    if not s:
        return ""
    # Для строк с несколькими словами приводим каждое слово к Title-case
    return " ".join(part.capitalize() for part in s.split())

# 6. «Умные» ключи поиска с учётом регистра и хомоглифов
#
# Для поиска по названию/подрядчику/отделу/зоне требуется игнорировать
# регистр и трактовать похожие кириллические и латинские символы
# как одинаковые (например, 'c' == 'с', 'm' == 'м').
# Эти функции позволяют генерировать канонический ключ для поиска
# и проверять вхождение подстроки.

# 6.1 Карта замены: кириллица → латиница для похожих букв
_HOMO_CYR_TO_LAT = {
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x",
    "к": "k", "м": "m", "т": "t", "в": "b", "н": "h", "ё": "e", "і": "i",
    "ї": "i", "й": "i", "ґ": "g",
    "А": "a", "Е": "e", "О": "o", "Р": "p", "С": "c", "У": "y", "Х": "x",
    "К": "k", "М": "m", "Т": "t", "В": "b", "Н": "h", "Ё": "e", "І": "i",
    "Ї": "i", "Й": "i", "Ґ": "g",
}

# 6.2 Удаляемые символы: диакритики и апострофы, которые мешают поиску
_STRIP_CHARS = {
    "’": "", "ʼ": "", "ʹ": "", "ʾ": "", "ʿ": "", "ˈ": "", "ˌ": "",
    "̀": "", "́": "", "̂": "", "̈": "", "̃": "",
}

def make_search_key(s: Any) -> str:
    """Возвращает канонический ключ для поиска:

    * Приводит к строке, Unicode NFKD, затем casefold (для нечувствительности к регистру).
    * Удаляет комбинируемые диакритики (Mn). Затем заменяет кириллические хомоглифы
      на латиницу (например, 'с' → 'c', 'м' → 'm').
    * Удаляет апострофы и пробелы‑паразиты, схлопывает последовательности пробелов.

    :param s: исходная строка (или объект, который можно привести к строке)
    :return: каноническая строка для поиска
    """
    if not s:
        return ""
    try:
        import unicodedata
        t = unicodedata.normalize("NFKD", str(s)).casefold()
        # Удаляем диакритики
        t = "".join(ch for ch in t if unicodedata.category(ch) != "Mn")
        # Замена кириллицы
        t = "".join(_HOMO_CYR_TO_LAT.get(ch, ch) for ch in t)
        # Удаляем апострофы и прочие
        for bad, rep in _STRIP_CHARS.items():
            t = t.replace(bad, rep)
        # Схлопываем пробелы
        t = " ".join(t.split())
        return t
    except Exception:
        # В случае ошибок возвращаем нижний регистр исходного текста
        try:
            return str(s).casefold()
        except Exception:
            return ""

def contains_search(haystack: Any, needle: Any) -> bool:
    """Проверяет, входит ли канон needle в канон haystack."""
    hk = make_search_key(haystack)
    nk = make_search_key(needle)
    if not nk:
        return True
    return hk.find(nk) >= 0

# 5. Автоширины столбцов и приоритет «Наименования»
def setup_auto_col_resize(table: QtWidgets.QTableWidget):
    """Включить авто-подгон ширины столбцов по содержимому."""
    hdr = table.horizontalHeader()
    hdr.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
    hdr.setStretchLastSection(True)

def setup_priority_name(table: QtWidgets.QTableWidget, name_col: int = 0):
    """Сделать колонку 'Наименование' растягиваемой, остальные — по содержимому."""
    hdr = table.horizontalHeader()
    for c in range(table.columnCount()):
        hdr.setSectionResizeMode(c, QtWidgets.QHeaderView.ResizeToContents)
    hdr.setSectionResizeMode(name_col, QtWidgets.QHeaderView.Stretch)
    hdr.setStretchLastSection(False)

def apply_auto_col_resize(table: QtWidgets.QTableWidget):
    """Применить пересчёт ширин (без ошибок при отсутствии данных)."""
    try:
        table.resizeColumnsToContents()
    except Exception:
        pass
