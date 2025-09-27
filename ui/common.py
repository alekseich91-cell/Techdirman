"""
Назначение:
    Общие константы и утилиты UI для TechDirRentMan. Этот модуль определяет
    пути проекта, конвертацию классов RU↔EN, параметры внешнего вида и
    функции вспомогательного форматирования и нормализации данных. В начале
    файла приведено краткое описание назначения и принципа работы.

Как работает:
    - Определяет корневые директории приложения (APP_ROOT, DATA_DIR,
      ASSETS_DIR).
    - Содержит словари соответствий классов на русском и английском.
    - Определяет порог переноса текста в ячейках таблиц.
    - Предоставляет функции для форматирования чисел и преобразования
      строк в числовой тип.
    - Реализует функцию normalize_case для приведения строк к каноническому
      виду: убирает ведущие/конечные пробелы (в том числе неразрывные
      пробелы), схлопывает последовательности пробелов, приводит каждое
      слово к Title‑case и устраняет различия в регистре. Для диагностики
      функция записывает ошибки в лог через logging.
    - Предоставляет функции для генерации ключей поиска с учётом кириллических
      хомоглифов и диакритических символов.
    - Содержит утилиты настройки ширины колонок Qt таблиц.

Стиль:
    - Код разбит на пронумерованные секции с краткими комментариями,
      поясняющими назначение каждого блока.
"""

# 1. Импорт библиотек
from pathlib import Path  # пути проекта
from typing import Any    # типы для аннотаций
from PySide6 import QtWidgets  # для настроек таблиц
import logging  # для вывода информационных и ошибочных сообщений

# Создаём логгер для модуля. Основная конфигурация задаётся в utils.init_logging().
logger = logging.getLogger(__name__)

# 2. Пути проекта
APP_ROOT = Path(__file__).resolve().parents[2]  # корень приложения (папка TechDirRentMan)
DATA_DIR = APP_ROOT / "data"                   # папка данных проекта
ASSETS_DIR = APP_ROOT / "assets"               # папка ассетов (картинки проекта)

# 3. Классы RU↔EN и константы внешнего вида
CLASS_RU2EN = {
    "Оборудование": "equipment",
    "Персонал": "personnel",
    "Логистика": "logistic",
    "Расходник": "consumable",
}
CLASS_EN2RU = {v: k for k, v in CLASS_RU2EN.items()}

# Порог символов для переноса текста в колонке «Наименование»
WRAP_THRESHOLD = 40

# 4. Утилиты форматирования чисел
def fmt_num(value: Any, decimals: int = 2) -> str:
    """Строка числа без лишних нулей: '10', '10,5', '10,25'.
    Десятичный разделитель — запятая.

    :param value: исходное число или строка
    :param decimals: количество знаков после запятой
    :return: строковое представление числа с учётом формата
    """
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
    """То же, что fmt_num, но со знаком для положительных значений.

    :param value: исходное число или строка
    :param decimals: количество знаков после запятой
    :return: строковое представление числа со знаком
    """
    try:
        v = float(value)
    except Exception:
        return "0"
    sign = "+" if v > 0 else ""
    return f"{sign}{fmt_num(v, decimals)}"


def to_float(text: Any, default: float = 0.0) -> float:
    """Надёжный парсер чисел из строк.

    Заменяет пробелы и запятые на точки, затем приводит к float. В случае
    ошибок возвращает значение по умолчанию.

    :param text: строка, число или другой объект
    :param default: значение по умолчанию при ошибке
    :return: число float
    """
    if isinstance(text, (int, float)):
        return float(text)
    try:
        return float(str(text).strip().replace(" ", "").replace(",", "."))
    except Exception:
        return float(default)


# 5. Нормализация строк
def normalize_case(text: Any) -> str:
    """Нормализует регистр и пробелы строковых значений.

    Функция приводит входной текст к каноническому виду:
        * преобразует объект к строке;
        * заменяет специальные пробельные символы (в т.ч. неразрывные пробелы
          U+00A0 и тонкие U+202F) на обычные пробелы;
        * удаляет пробелы в начале и конце строки;
        * схлопывает последовательности пробелов внутри строки;
        * приводит каждое слово к Title‑case (первая буква заглавная).
    В случае ошибок функция пишет сообщение в лог и возвращает пустую строку.

    :param text: входная строка или объект, который может быть приведён к строке
    :return: нормализованная строка
    """
    try:
        s = str(text)
    except Exception as ex:
        # Логируем ошибку преобразования в строку
        logger.error("normalize_case: не удалось привести к строке: %s", ex, exc_info=True)
        return ""
    # Заменяем особые пробелы на обычные, чтобы strip() и split() корректно
    # обрабатывали строки с неразрывными пробелами
    try:
        s = s.replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
    except Exception as ex:
        logger.error("normalize_case: ошибка замены пробелов: %s", ex, exc_info=True)
        # продолжаем с исходной строкой
    # Удаляем начальные и конечные пробелы
    s = s.strip()
    # Если строка пуста после обрезки — возвращаем пустую строку
    if not s:
        return ""
    try:
        # split() без аргументов делит по любым пробельным символам, включая
        # табуляции и множественные пробелы, но не захватывает неразрывные
        # пробелы — они уже заменены на обычные. Склеиваем слова с одним
        # пробелом и приводим каждое к Title‑case.
        return " ".join(part.capitalize() for part in s.split())
    except Exception as ex:
        # Логируем ошибку нормализации и возвращаем исходную строку
        logger.error("normalize_case: ошибка нормализации '%s': %s", text, ex, exc_info=True)
        return s


# 5.1. Очистка пробелов
def clean_start(text: Any) -> str:
    """Удаляет ведущие пробельные символы из значения.

    Эта утилита расширяет :func:`str.lstrip`, предварительно заменяя
    неразрывные пробелы (U+00A0), тонкие пробелы (U+202F) и табличные пробелы
    (U+2007) на обычные, чтобы затем корректно удалить их. Функция приводит
    значение к строке, в случае ошибки записывает сообщение в лог и
    возвращает пустую строку.

    :param text: значение, из которого удаляются лидирующие пробелы
    :return: строка без пробельных символов в начале
    """
    try:
        s = str(text)
    except Exception as ex:
        logger.error("clean_start: не удалось привести к строке '%s': %s", text, ex, exc_info=True)
        return ""
    try:
        # Заменяем особые пробелы на обычные
        s = s.replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
    except Exception as ex:
        logger.error("clean_start: ошибка замены пробелов: %s", ex, exc_info=True)
    return s.lstrip()


def clean_edges(text: Any) -> str:
    """Удаляет пробельные символы по краям значения.

    В отличие от стандартного :func:`str.strip`, эта функция сначала
    нормализует необычные пробелы (неразрывные, тонкие и табличные) к
    обычным. Это позволяет гарантировать, что "invisible" пробелы также
    будут удалены. При ошибках функция логирует исключение и возвращает
    пустую строку.

    :param text: значение, которое требуется обрезать
    :return: обрезанная строка без лидирующих и завершающих пробелов
    """
    try:
        s = str(text)
    except Exception as ex:
        logger.error("clean_edges: не удалось привести к строке '%s': %s", text, ex, exc_info=True)
        return ""
    try:
        s = s.replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
    except Exception as ex:
        logger.error("clean_edges: ошибка замены пробелов: %s", ex, exc_info=True)
    return s.strip()


# 6. «Умные» ключи поиска с учётом регистра и хомоглифов

# Для поиска по названию/подрядчику/отделу/зоне требуется игнорировать
# регистр и трактовать похожие кириллические и латинские символы как
# одинаковые (например, 'c' == 'с', 'm' == 'м'). Эти функции позволяют
# генерировать канонический ключ для поиска и проверять вхождение
# подстроки.

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
    """Возвращает канонический ключ для поиска.

    Алгоритм:
        * приводит значение к строке, нормализует в форму Unicode NFKD и
          приводит к lower‑case (casefold) для нечувствительности к регистру;
        * удаляет комбинируемые диакритические символы;
        * заменяет кириллические хомоглифы на латиницу;
        * удаляет апострофы и другие символы из _STRIP_CHARS;
        * схлопывает последовательности пробелов.

    :param s: исходная строка или объект, который можно привести к строке
    :return: нормализованная строка для поиска
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
        # Удаляем апострофы и прочие символы
        for bad, rep in _STRIP_CHARS.items():
            t = t.replace(bad, rep)
        # Схлопываем пробелы
        t = " ".join(t.split())
        return t
    except Exception as ex:
        # В случае ошибок возвращаем нижний регистр исходного текста и логируем
        logger.error("make_search_key: ошибка нормализации '%s': %s", s, ex, exc_info=True)
        try:
            return str(s).casefold()
        except Exception:
            return ""


def contains_search(haystack: Any, needle: Any) -> bool:
    """Проверяет, входит ли канон needle в канон haystack.

    :param haystack: строка, в которой производится поиск
    :param needle: подстрока для поиска
    :return: True, если подстрока найдена, иначе False
    """
    hk = make_search_key(haystack)
    nk = make_search_key(needle)
    if not nk:
        return True
    return hk.find(nk) >= 0


# 7. Автоширины столбцов и приоритет «Наименования»
def setup_auto_col_resize(table: QtWidgets.QTableWidget) -> None:
    """Включить авто-подгон ширины столбцов по содержимому.

    :param table: таблица Qt, в которой требуется настроить режим автоподгонки
    """
    hdr = table.horizontalHeader()
    hdr.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
    hdr.setStretchLastSection(True)


def setup_priority_name(table: QtWidgets.QTableWidget, name_col: int = 0) -> None:
    """Сделать колонку 'Наименование' растягиваемой, остальные — по содержимому.

    :param table: таблица Qt
    :param name_col: индекс колонки "Наименование", которую нужно растянуть
    """
    hdr = table.horizontalHeader()
    for c in range(table.columnCount()):
        hdr.setSectionResizeMode(c, QtWidgets.QHeaderView.ResizeToContents)
    hdr.setSectionResizeMode(name_col, QtWidgets.QHeaderView.Stretch)
    hdr.setStretchLastSection(False)


def apply_auto_col_resize(table: QtWidgets.QTableWidget) -> None:
    """Применить пересчёт ширин без ошибок при отсутствии данных.

    :param table: таблица Qt
    """
    try:
        table.resizeColumnsToContents()
    except Exception:
        # Игнорируем ошибки, например, если таблица пуста
        pass
