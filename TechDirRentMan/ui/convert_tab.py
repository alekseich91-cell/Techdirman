"""
Назначение
==========

Этот модуль реализует вкладку «Конвертация» для приложения TechDirRentMan.
Он предоставляет пользователю drag‑and‑drop интерфейс для загрузки PDF‑файлов
и конвертации их в Excel. В рамках вкладки доступны два режима работы:
— **Rentman VSG** и **Rentman Jamteck**. Первый режим предназначен
для коммерческих предложений, сформированных системой Rentman/VSG, где
табличные данные представлены без линий, но следуют строгому порядку:
номер позиции, название, цена за единицу, количество, коэффициент, сумма.

Модуль разбит на пронумерованные секции с краткими заголовками. В
комментариях поясняется назначение и логика работы каждой функции и
блоков кода. Все операции протоколируются в лог файл ``convert_tab.log``,
что позволяет отслеживать успешные действия и ошибки.

Дополнительная логика фильтрации нулевых строк
---------------------------------------------

В некоторых коммерческих предложениях встречаются строки, которые
выглядят пустыми: текст в ячейках скрыт с помощью белого или
прозрачного шрифта. Обычно у таких строк итоговая сумма равна нулю.
Чтобы исключить подобные позиции из результирующей таблицы, в
алгоритм добавлены проверки на нулевую стоимость, количество и
коэффициент. Если хотя бы один из параметров равен нулю, строка
пропускается.
"""

# 1. Импорт стандартных библиотек
import os
import logging
from pathlib import Path
from typing import Any, Optional

from PySide6 import QtWidgets, QtCore, QtGui

# Импорт общих путей (можно использовать для хранения временных файлов)
from .common import ASSETS_DIR, DATA_DIR

# 0. Настройка логирования
# Создаём директорию для логов (если не существует)
LOG_DIR = os.path.join(os.path.dirname(__file__), "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "convert_tab.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("convert_tab")


# 2. Вспомогательные функции: конвертация PDF → Excel
def convert_pdf_to_excel(
    pdf_path: Path,
    dest_path: Path,
    engine: str = "pdfplumber",
    *,
    manual_bounds: Optional[list[float]] = None,
) -> None:
    """
    Конвертирует PDF‑файл в XLSX, собирая все строки таблиц на одном листе.

    Несмотря на наличие параметров ``engine`` и ``manual_bounds`` (оставленных
    для обратной совместимости), функция всегда использует встроенный режим
    Rentman VSG. Этот режим анализирует каждую строку, извлечённую из
    PDF с помощью PyMuPDF, и ищет две суммы (значения с символом «₽»).
    Если строка относится к разделу аренды, она должна содержать как минимум
    количество и коэффициент; строки, в которых числовые поля отсутствуют,
    не попадают в таблицу. В разделе расходной части коэффициент отсутствует
    и принимается равным 1. Заголовочные и сводные строки пропускаются.

    Также добавлены проверки на нулевые значения: если итоговая сумма,
    цена, количество или коэффициент равны нулю, такая строка
    пропускается. Эти проверки помогают отфильтровывать строки с
    скрытыми значениями, в которых сумма указана как ``0 ₽``.

    :param pdf_path: путь к исходному PDF.
    :param dest_path: путь к итоговому XLSX. Папка будет создана при необходимости.
    :param engine: устаревший параметр, игнорируется.
    :param manual_bounds: устаревший параметр, игнорируется.

    :raises RuntimeError: при ошибке чтения или записи.
    """
    # 2.1 Импорт зависимостей
    try:
        import pandas as pd  # type: ignore
    except ImportError as ex:
        msg = (
            "Библиотека pandas не установлена. Добавьте её в requirements.txt или "
            "установите вручную."
        )
        logger.error(msg)
        raise RuntimeError(msg) from ex

    # 2.2 Функция очистки и распознавания чисел
    def _to_number(val: Any) -> Any:
        """
        Преобразует строку в число: удаляет пробелы (включая неразрывные),
        запятые заменяет на точки, убирает символы валют и букв. Если
        преобразование не удаётся, возвращает исходное значение.
        """
        if isinstance(val, str):
            s = val.strip()
            # убираем различные пробелы и символы валюты
            for ch in [" ", "\u00A0", "\u202F", "₽", "р", "Р"]:
                s = s.replace(ch, "")
            s = s.replace(",", ".")
            # если после удаления остались только цифры и точка — это число
            try:
                num = float(s)
                return int(num) if num.is_integer() else num
            except Exception:
                return val
        return val

    # 2.3 Нормализация извлечённой таблицы
    def _normalize_dataframe(df: 'pd.DataFrame') -> list['pd.DataFrame']:
        """
        Разбивает таблицу на набор нормализованных DataFrame. Каждая
        строка исходной таблицы анализируется: если содержит 3–4
        числовых значения, интерпретируется как запись с колонками
        (Оборудование, Кол‑во, Цена за ед., Коэфф., Сумма). Строки с
        меньшим количеством чисел остаются как текст. Строки, в
        которых встречается более четырёх числовых фрагментов (из-за
        разбиения больших чисел на части), также трактуются как
        текстовые записи, чтобы избежать некорректного разделения.

        Заголовочные строки распознаются по ключевым словам (например,
        «оборудование», «кол‑во», «коэфф.», «сумма», «цена» и т.д.),
        но пропускаются только в том случае, если в строке нет
        числовых значений. Это позволяет не отбрасывать сводные
        строки, содержащие числа (например, «Итого сумма»), которые
        необходимо сохранить.

        Возвращает список DataFrame: одни со столбцами таблицы, другие
        одно‑колоночные с текстом. В случае ошибки строка
        конвертируется в текстовую запись.
        """
        frames: list['pd.DataFrame'] = []
        for _, row in df.iterrows():
            try:
                # Представляем все ячейки как строки, удаляя NaN
                vals = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
                nums: list[str] = []  # числовые фрагменты строки
                texts: list[str] = []  # текстовые фрагменты строки
                for cell in vals:
                    if not cell:
                        continue
                    raw = cell
                    tmp = raw
                    # Убираем разделители и валютные символы для проверки
                    for ch in [" ", "\u00A0", "\u202F", "₽", "р", "Р"]:
                        tmp = tmp.replace(ch, "")
                    tmp = tmp.replace(",", ".")
                    # Проверяем, является ли значение числом
                    try:
                        float(tmp)
                        nums.append(tmp)
                        continue
                    except Exception:
                        pass
                    texts.append(raw)
                # Собираем строку из текстовых ячеек и приводим к нижнему регистру
                header_joined = " ".join(texts).lower()

                # 2.3.1 Списки ключевых слов для распознавания заголовков и сводных строк
                header_keywords = [
                    # Ключевые слова для распознавания строк‑заголовков таблицы.
                    # Исключаем слово "оборудование", чтобы не пропускать категории
                    "кол-",
                    "коэфф",
                    "сум",
                    "наименование",
                    "цена",
                ]
                # Сводные строки (итоговые суммы, скидки, налоги и т.п.)
                summary_keywords = [
                    "итого",
                    "скидк",  # скидка, скидки
                    "налог",
                    "ставка",
                    "сумма",
                    "прокат",
                    "персонал",
                    "транспорт",
                    "налогов",
                    "наклад",
                ]
                # Если строка похожа на заголовок (есть ключевые слова) и нет числовых значений — пропускаем её
                if not nums and any(k in header_joined for k in header_keywords):
                    continue
                # Если строка содержит слова из summary_keywords, считаем её сводной
                is_summary = any(k in header_joined for k in summary_keywords)
                # Для строк с 3–4 числовыми значениями и без признаков сводной строки формируем табличную запись
                if 3 <= len(nums) <= 4 and not is_summary:
                    name = " ".join(texts).strip() or ""
                    # Распределяем значения: при 4 числах порядок (price, qty, coeff, total)
                    # При 3 числах коэффициент отсутствует и принимается равным 1
                    if len(nums) == 4:
                        price_str, qty_str, coeff_str, total_str = nums
                    else:
                        price_str, qty_str, total_str = nums
                        coeff_str = "1"
                    # Преобразуем строковые значения в числа, если это возможно
                    try:
                        price_val = _to_number(price_str)
                        qty_val = _to_number(qty_str)
                        coeff_val = _to_number(coeff_str)
                        total_val = _to_number(total_str)
                    except Exception:
                        price_val = price_str
                        qty_val = qty_str
                        coeff_val = coeff_str
                        total_val = total_str
                    frames.append(
                        pd.DataFrame(
                            [
                                {
                                    "Оборудование": name,
                                    "Кол-во": qty_val,
                                    "Цена за ед.": price_val,
                                    "Коэфф.": coeff_val,
                                    "Сумма": total_val,
                                }
                            ]
                        )
                    )
                else:
                    # Для остальных случаев — формируем текстовую запись, сохраняя строку полностью
                    full_text = " ".join([v for v in vals if v])
                    if full_text:
                        frames.append(pd.DataFrame({"text": [full_text]}))
            except Exception:
                # В случае ошибки формируем текстовую запись из всей строки
                full_text = " ".join([str(v) for v in row.tolist() if str(v)])
                if full_text:
                    frames.append(pd.DataFrame({"text": [full_text]}))
        return frames

    # 2.A Специализированный режим Rentman VSG
    #
    # Алгоритм ниже исполняется для всех PDF-файлов, игнорируя параметр
    # ``engine``. Он предназначен для коммерческих предложений, где
    # каждая строка с оборудованием содержит две суммы: цену за единицу и
    # итоговую сумму. Строки, не соответствующие этому формату, не
    # обрабатываются. После завершения алгоритма функция возвращает,
    # предотвращая выполнение общего кода для pdfplumber/pymupdf/regex.
    try:
        import fitz  # type: ignore
    except ImportError as ex:
        msg = (
            "Библиотека PyMuPDF (fitz) не установлена. Добавьте её в requirements.txt или "
            "установите вручную."
        )
        logger.error(msg)
        raise RuntimeError(msg) from ex

    # Ключевые слова для сводных строк, которые нужно пропускать
    summary_keywords = [
        "итого",
        "скидк",
        "налог",
        "ставка",
        "сумма",
        "прокат",
        "персонал",
        # оставляем только слова, однозначно встречающиеся в заголовках.
        # Убираем слово "транспорт", чтобы не пропускать строки типа
        # "Транспортировка по Москве", которые являются строками таблицы.
        "налогов",
        "наклад",
    ]

    # 2.A.1 Состояние: находимся ли мы в разделе «Расходная часть»
    # На первых страницах идут разделы аренды (основные позиции),
    # затем встречается заголовок с ключевым словом «расход», после чего
    # строки имеют иную структуру (нет коэффициента, возможно только цена и сумма).
    expenses_section = False

    # Проверка, является ли токен числовым (удаляются пробелы и запятые)
    def _is_numeric_token(token: str) -> bool:
        cleaned = (
            token.replace(" ", "")
            .replace("\u00A0", "")
            .replace("\u202F", "")
            .replace(",", "")
        )
        return cleaned.isdigit()

    # Список записей (словарей), каждая представляет строку таблицы.
    records: list[dict] = []
    try:
        doc = fitz.open(str(pdf_path))
        for page_index, page in enumerate(doc, start=1):
            try:
                words = page.get_text("words")  # type: ignore[attr-defined]
            except Exception as ex:
                logger.error(
                    "Ошибка извлечения слов на странице %s: %s", page_index, ex, exc_info=True
                )
                continue
            if not words:
                continue
            # Группировка слов в строки по координате Y
            heights = [w[3] - w[1] for w in words]
            avg_height = sum(heights) / len(heights) if heights else 0.0
            y_threshold = avg_height * 0.6 if avg_height else 2.0
            words_sorted = sorted(words, key=lambda w: w[1])
            rows: list[list] = []
            current_row: list = []
            current_y: Optional[float] = None
            for w in words_sorted:
                y0 = w[1]
                if current_y is None or abs(y0 - current_y) <= y_threshold:
                    current_row.append(w)
                    current_y = y0 if current_y is None else (current_y + y0) / 2
                else:
                    if current_row:
                        current_row.sort(key=lambda x: x[0])
                        rows.append(current_row)
                    current_row = [w]
                    current_y = y0
            if current_row:
                current_row.sort(key=lambda x: x[0])
                rows.append(current_row)
            # Обработка каждой строки
            for row_words in rows:
                if not row_words:
                    continue
                tokens = [w[4] for w in row_words]
                # Проверяем на сводные строки
                row_text_lower = " ".join(tokens).lower()
                # Детектируем переход в раздел «расходная часть». Используем
                # более точную проверку: ищем словосочетание "расходная"
                # (например, "Расходная часть"). Это позволяет избежать
                # ложных срабатываний на слова типа "расходники".
                if "расходная" in row_text_lower:
                    expenses_section = True
                    continue
                # Пропускаем строки, содержащие сводные слова (итоги, скидки, налоги)
                if any(k in row_text_lower for k in summary_keywords):
                    continue
                # Убираем первый токен, если это номер позиции
                if tokens and tokens[0].strip().isdigit():
                    tokens = tokens[1:]
                if not tokens:
                    continue
                # Индексы валютных токенов
                currency_indices = [i for i, t in enumerate(tokens) if "₽" in t]
                if len(currency_indices) < 2:
                    continue
                last_idx = currency_indices[-1]
                second_last_idx = currency_indices[-2]
                # Сбор цены за единицу: все подряд идущие числовые токены непосредственно
                # перед предпоследним знаком валюты. Эти токены составляют цену.
                price_tokens_list: list[str] = []
                idx_p = second_last_idx - 1
                while idx_p >= 0 and _is_numeric_token(tokens[idx_p]):
                    price_tokens_list.insert(0, tokens[idx_p])
                    idx_p -= 1
                if not price_tokens_list:
                    continue
                # В некоторых строках перед ценой встречаются лишние числовые токены (например,
                # номер или счётчик типа «2»), которые не являются частью цены. Типичная цена
                # в коммерческих предложениях состоит из двух токенов (тысячи и сотни) либо
                # трёх токенов, если присутствуют копейки (пример: «16 000 00» -> «16 000,00»).
                # Если извлечено более двух токенов и последний токен не является двухзначным
                # (т.е. это не копейки), считаем, что лишние начальные токены относятся к
                # наименованию, а не к цене. Сдвигаем их влево.
                _initial_len = len(price_tokens_list)
                # Если извлечено более двух токенов и последний токен не двухзначный (нет копеек),
                # это может быть длинная цена (например «1 500 000»), но иногда первый токен —
                # лишний счётчик ("2"), который должен попасть в название. Будем считать,
                # что токен относится к имени, если всего три токена и второй токен имеет
                # длину ≤ 2 символов (обычно «25»), что характерно для цен вида «25 000».
                if _initial_len > 2 and len(price_tokens_list[-1]) != 2:
                    # проверяем длину второго токена: если <=2, убираем лишние первые токены
                    second_len = len(price_tokens_list[1]) if _initial_len >= 2 else 0
                    if second_len <= 2:
                        shift = _initial_len - 2
                        price_tokens_list = price_tokens_list[shift:]
                        price_tokens_shift = shift
                    else:
                        price_tokens_shift = 0
                else:
                    price_tokens_shift = 0

                # Промежуточные числовые токены между ценой и итоговой суммой. Они могут
                # включать количество, коэффициент и части суммы (если она разбивается на несколько токенов).
                mid_numeric: list[str] = [
                    t
                    for t in tokens[second_last_idx + 1 : last_idx]
                    if _is_numeric_token(t)
                ]

                # Объединение числовых токенов в строку. Если последняя часть имеет длину две
                # цифры и токенов три и более, считаем её дробной частью (пример: 16 000 00 -> 16 000,00).
                def _join_number_tokens(num_tokens: list[str]) -> str:
                    if not num_tokens:
                        return ""
                    if len(num_tokens) >= 3 and len(num_tokens[-1]) == 2:
                        ints = num_tokens[:-1]
                        decimals = num_tokens[-1]
                        return " ".join(ints) + "," + decimals
                    return " ".join(num_tokens)

                # Числовое значение цены. Если цена равна нулю, строка скрыта – пропускаем.
                try:
                    price_val_tmp = _to_number(_join_number_tokens(price_tokens_list))
                except Exception:
                    price_val_tmp = None
                if isinstance(price_val_tmp, (int, float)) and price_val_tmp == 0:
                    continue

                # Определяем количество (qty), коэффициент (coeff) и элементы суммы (sum_tokens_list)
                qty_str: Optional[str] = None
                coeff_str: Optional[str] = None
                sum_tokens_list: list[str] = []

                if not expenses_section:
                    # В разделе аренды: первый токен – количество, второй – коэффициент,
                    # оставшиеся – части суммы. Если коэффициента нет, считаем его равным 1.
                    if not mid_numeric:
                        continue
                    if len(mid_numeric) >= 2:
                        qty_str = mid_numeric[0]
                        coeff_str = mid_numeric[1]
                        sum_tokens_list = mid_numeric[2:]
                    else:
                        qty_str = mid_numeric[0]
                        coeff_str = "1"
                        sum_tokens_list = []
                else:
                    # В разделе расходов: коэффициент всегда 1, количество – первый токен, остальные – части суммы.
                    if not mid_numeric:
                        continue
                    qty_str = mid_numeric[0]
                    coeff_str = "1"
                    sum_tokens_list = mid_numeric[1:]

                # Пробуем извлечь явную сумму из токенов суммы. Если сумма указана и равна нулю – пропускаем строку.
                sum_val_extracted: Optional[float] = None
                if sum_tokens_list:
                    try:
                        s = _to_number(_join_number_tokens(sum_tokens_list))
                        if isinstance(s, (int, float)):
                            sum_val_extracted = s
                    except Exception:
                        sum_val_extracted = None
                if sum_val_extracted == 0:
                    continue

                # Имя позиции – все токены до начала цены. Если из цены были удалены
                # лишние числовые токены (price_tokens_shift), учтём это смещение, чтобы
                # перенести их в название. price_start_index вычисляется из исходной
                # длины price_tokens_list и сдвига.
                price_start_index = (second_last_idx - (_initial_len)) + price_tokens_shift
                name_tokens = tokens[:price_start_index]
                name = " ".join(name_tokens).strip()
                if not name:
                    continue

                # Преобразуем цену, количество и коэффициент в числа. При ошибке пропускаем строку.
                try:
                    price_val = _to_number(_join_number_tokens(price_tokens_list))
                    qty_val = _to_number(qty_str) if qty_str is not None else None
                    coeff_val = _to_number(coeff_str) if coeff_str is not None else None
                except Exception:
                    continue
                # Проверяем, что цена, количество и коэффициент – числа
                if not (
                    isinstance(price_val, (int, float))
                    and isinstance(qty_val, (int, float))
                    and isinstance(coeff_val, (int, float))
                ):
                    continue
                # Пропускаем строки с нулевыми значениями
                if price_val == 0 or qty_val == 0 or coeff_val == 0:
                    continue
                # Вычисляем итоговую сумму как произведение цены, количества и коэффициента
                total_val = price_val * qty_val * coeff_val
                if total_val == 0:
                    continue
                # Добавляем запись в общий список
                records.append(
                    {
                        "Наименование": name,
                        "Кол-во": qty_val,
                        "Цена за ед.": price_val,
                        "Коэфф.": coeff_val,
                        "Сумма": total_val,
                    }
                )
    except Exception as ex:
        logger.error("Ошибка при чтении PDF: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации: {ex}")

    # Запись в Excel и завершение. После этого выполняется return, чтобы избежать
    # работы старых режимов конвертации.
    try:
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        # Формируем итоговый DataFrame из записей. Заголовок будет записан один раз.
        df_result = pd.DataFrame(
            records,
            columns=["Наименование", "Кол-во", "Цена за ед.", "Коэфф.", "Сумма"],
        )
        with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="Sheet1", index=False)
    except Exception as ex:
        logger.error("Ошибка записи в Excel: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации: {ex}")
    return

    # 2.B Вспомогательные функции для режима Rentman Jamteck

    # NOTE: код Jamteck вынесен в отдельный раздел, чтобы сохранить
    # читаемость. Этот режим использует библиотеку pdfplumber для
    # извлечения текста и регулярные выражения для парсинга строк.

# 2.B.1 Преобразование строки цены в число
def _jamteck_parse_price(text: str) -> float:
    """
    Преобразует строку стоимости, например «1 590,00», в число float.

    В PDF использованы пробелы для разделения тысяч и запятая
    для десятичной точки. Функция удаляет пробельные и валютные
    символы и возвращает число. В случае ошибки возвращает 0.0.

    :param text: строка с числом, без символа валюты.
    :return: вещественное значение стоимости.
    """
    cleaned = text.replace("₽", "").replace("\xa0", "").replace(" ", "")
    cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except Exception:
        logger.error("Jamteck: не удалось преобразовать цену '%s'", text)
        return 0.0


# 2.B.2 Разбор PDF в формате Jamteck
def _jamteck_parse_pdf(pdf_path: Path) -> dict[str, list[dict]]:
    """
    Извлекает данные из PDF коммерческого предложения Jamteck.

    В файле могут присутствовать несколько суб‑проектов. Каждому
    суб‑проекту соответствует ключ словаря, значение которого –
    список словарей с колонками:

      * ``Наименование`` – имя позиции
      * ``Количество`` – целое число
      * ``Цена за единицу`` – float
      * ``Коэффициент`` – целое число
      * ``Сумма`` – float

    Строки с нулевой суммой пропускаются. Строки с некорректной
    структурой игнорируются.

    :param pdf_path: путь к исходному PDF
    :return: словарь суб‑проектов
    """
    try:
        import pdfplumber  # type: ignore
    except ImportError as ex:
        msg = (
            "Библиотека pdfplumber не установлена. Добавьте её в requirements.txt или "
            "установите вручную."
        )
        logger.error(msg)
        raise RuntimeError(msg) from ex
    logging.info("Jamteck: Парсинг PDF %s", pdf_path)
    import re
    # регулярки для суб‑проекта и хвостов колонок
    subproject_re = re.compile(r"^Суб-проект: (.+)$")
    # Паттерны для хвостов строк.
    # Количество (1–2 цифры) обязательно должно быть отделено пробелом; это исключает цифры внутри названия (например, G500).
    equip_tail_re = re.compile(r"(?<=\s)(\d+)\s+([\d\s]+,\d{2})\s+(\d+)\s+([\d\s]+,\d{2})$")
    personnel_tail_re = re.compile(r"(?<=\s)(\d+)\s+([\d\s]+,\d{2})$")
    skip_prefixes = ["Итого", "Подытог", "Цена"]
    skip_contains = ["Налоги", "Сумма", "УСН", "НДС", "Доп.", "Подтверждение"]
    projects: dict[str, list[dict]] = {}
    current_proj: str | None = None
    capturing = False
    # читаем весь текст
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as ex:
        logger.error("Jamteck: ошибка чтения PDF: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка чтения PDF: {ex}")
    lines = text.split("\n")
    for raw_line in lines:
        line = (
            raw_line.replace("₽", "")
            .replace("\xa0", " ")
            .replace("\u202f", " ")
            .strip()
        )
        m = subproject_re.match(line)
        if m:
            current_proj = m.group(1).strip()
            projects[current_proj] = []
            capturing = True
            logger.info("Jamteck: найден суб-проект %s", current_proj)
            continue
        if capturing and line.startswith("Цена:"):
            capturing = False
            continue
        if not capturing or current_proj is None or not line:
            continue
        if any(line.startswith(p) for p in skip_prefixes) or any(word in line for word in skip_contains):
            continue
        eq_match = equip_tail_re.search(line)
        if eq_match:
            qty_str, unit_str, coeff_str, total_str = eq_match.groups()
            name = line[: eq_match.start()].strip()
            try:
                quantity = int(qty_str)
                coeff = int(coeff_str)
                unit_val = _jamteck_parse_price(unit_str)
                total_val = _jamteck_parse_price(total_str)
            except Exception:
                continue
            if total_val == 0:
                continue
            projects[current_proj].append(
                {
                    "Наименование": name,
                    "Количество": quantity,
                    "Цена за единицу": unit_val,
                    "Коэффициент": coeff,
                    "Сумма": total_val,
                }
            )
            continue
        pr_match = personnel_tail_re.search(line)
        if pr_match:
            qty_str, total_str = pr_match.groups()
            name = line[: pr_match.start()].strip()
            try:
                quantity = int(qty_str)
                total_val = _jamteck_parse_price(total_str)
            except Exception:
                continue
            if total_val == 0:
                continue
            unit_val = total_val / quantity if quantity else 0.0
            projects[current_proj].append(
                {
                    "Наименование": name,
                    "Количество": quantity,
                    "Цена за единицу": unit_val,
                    "Коэффициент": 1,
                    "Сумма": total_val,
                }
            )
            continue
        # прочие строки опускаем
        continue
    logger.info("Jamteck: завершён парсинг, найдено %d суб-проектов", len(projects))
    return projects


# 2.B.3 Запись данных Jamteck в Excel
def convert_pdf_to_excel_jamteck(pdf_path: Path, dest_path: Path) -> None:
    """
    Конвертирует PDF Jamteck в Excel.

    Создаёт отдельный лист для каждого суб‑проекта. Если данные не
    найдены, генерируется исключение. Использует pandas и openpyxl для
    записи.

    :param pdf_path: путь к исходному PDF
    :param dest_path: путь, куда будет сохранён XLSX
    """
    try:
        import pandas as pd  # type: ignore
    except ImportError as ex:
        msg = (
            "Jamteck: библиотека pandas не установлена. Добавьте её в requirements.txt или "
            "установите вручную."
        )
        logger.error(msg)
        raise RuntimeError(msg) from ex
    projects = _jamteck_parse_pdf(pdf_path)
    if not projects:
        raise RuntimeError("Jamteck: в выбранном файле не найдено таблиц для обработки.")
    try:
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
            for name, rows in projects.items():
                df = pd.DataFrame(rows)
                # Приводим числовые колонки к корректным типам
                for col in ["Количество", "Коэффициент"]:
                    if col in df.columns:
                        df[col] = df[col].astype(int)
                for col in ["Цена за единицу", "Сумма"]:
                    if col in df.columns:
                        df[col] = df[col].astype(float)
                sheet_name = name[:31] if name else "Sheet1"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as ex:
        logger.error("Jamteck: ошибка записи в Excel: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации Jamteck: {ex}")

    # 2.4 Извлечение данных
    aggregated: list['pd.DataFrame'] = []
    # 2.4.1 Режим ручной разметки (manual)
    if engine.lower() == "manual":
        try:
            import fitz  # type: ignore
        except ImportError as ex:
            msg = (
                "Библиотека PyMuPDF (fitz) не установлена. Добавьте её в requirements.txt или "
                "установите вручную."
            )
            logger.error(msg)
            raise RuntimeError(msg) from ex
        try:
            import pandas as _pd  # type: ignore
        except ImportError as ex:
            msg = (
                "Библиотека pandas не установлена. Добавьте её в requirements.txt или "
                "установите вручную."
            )
            logger.error(msg)
            raise RuntimeError(msg) from ex
        # Значения по умолчанию для границ колонок (5 колонок: Name, Qty, Price per unit, Coeff., Sum)
        # Границы заданы как относительные доли ширины страницы: от 0 до 1.0.
        # Пользователь может переопределить этот список через параметр manual_bounds.
        bounds = manual_bounds
        if not bounds:
            bounds = [0.0, 0.50, 0.60, 0.75, 0.85, 1.0]
        # Проверяем, что список отсортирован и значения находятся в [0,1]
        try:
            bounds = [float(x) for x in bounds]
            if sorted(bounds) != bounds or bounds[0] != 0.0 or bounds[-1] != 1.0:
                raise ValueError
        except Exception:
            logger.error(
                "Некорректный список границ столбцов для режима manual: %s", bounds
            )
            raise RuntimeError("Некорректный список границ столбцов")
        try:
            doc = fitz.open(str(pdf_path))
            for page_index, page in enumerate(doc, start=1):
                try:
                    words = page.get_text("words")  # type: ignore[attr-defined]
                except Exception as ex:
                    logger.error(
                        "Ошибка извлечения слов на странице %s: %s", page_index, ex, exc_info=True
                    )
                    continue
                if not words:
                    continue
                # PyMuPDF возвращает список списков: [x0, y0, x1, y1, text, block_no, line_no, word_no]
                # Сортируем слова по координате Y (относительно верхнего края)
                # Вычисляем среднюю высоту слова для определения порога кластеризации строк
                heights = [w[3] - w[1] for w in words]
                avg_height = sum(heights) / len(heights) if heights else 0.0
                y_threshold = avg_height * 0.6 if avg_height else 2.0
                # Сортируем по y0
                words_sorted = sorted(words, key=lambda w: w[1])
                rows: list[list] = []
                current_row: list = []
                current_y: Optional[float] = None
                for w in words_sorted:
                    y0 = w[1]
                    if current_y is None or abs(y0 - current_y) <= y_threshold:
                        current_row.append(w)
                        current_y = y0 if current_y is None else (current_y + y0) / 2
                    else:
                        # Сохраняем завершённую строку
                        if current_row:
                            current_row.sort(key=lambda x: x[0])
                            rows.append(current_row)
                        current_row = [w]
                        current_y = y0
                if current_row:
                    current_row.sort(key=lambda x: x[0])
                    rows.append(current_row)
                page_width = page.rect.width
                for row_words in rows:
                    # Объединяем слова в столбцы согласно границам
                    col_texts = ["" for _ in range(len(bounds) - 1)]
                    for w in row_words:
                        x_center = (w[0] + w[2]) / 2.0
                        rel_x = x_center / page_width
                        # находим индекс колонки
                        for i in range(len(bounds) - 1):
                            if bounds[i] <= rel_x < bounds[i + 1]:
                                col_texts[i] = (col_texts[i] + " " + w[4]).strip()
                                break
                    # Подсчёт числовых столбцов (начиная со второй колонки) для определения типа строки
                    numeric_count = 0
                    for text in col_texts[1:]:
                        tmp = text
                        for ch in [" ", "\u00A0", "\u202F", "₽", "р", "Р"]:
                            tmp = tmp.replace(ch, "")
                        tmp = tmp.replace(",", ".")
                        try:
                            float(tmp)
                            numeric_count += 1
                        except Exception:
                            pass
                    if numeric_count >= 2:
                        # Считаем, что это строка таблицы
                        # Сопоставляем значения колонкам: Оборудование, Кол‑во, Цена за ед., Коэфф., Сумма
                        # Если границ меньше, оставшиеся поля будут пустыми
                        name_val = col_texts[0] if len(col_texts) > 0 else ""
                        qty_val = col_texts[1] if len(col_texts) > 1 else ""
                        price_val = col_texts[2] if len(col_texts) > 2 else ""
                        coeff_val = col_texts[3] if len(col_texts) > 3 else ""
                        total_val = col_texts[4] if len(col_texts) > 4 else ""
                        try:
                            frame = _pd.DataFrame(
                                [
                                    {
                                        "Оборудование": name_val,
                                        "Кол-во": _to_number(qty_val),
                                        "Цена за ед.": _to_number(price_val),
                                        "Коэфф.": _to_number(coeff_val),
                                        "Сумма": _to_number(total_val),
                                    }
                                ]
                            )
                        except Exception:
                            # fallback: сохраняем без преобразования
                            frame = _pd.DataFrame(
                                [
                                    {
                                        "Оборудование": name_val,
                                        "Кол-во": qty_val,
                                        "Цена за ед.": price_val,
                                        "Коэфф.": coeff_val,
                                        "Сумма": total_val,
                                    }
                                ]
                            )
                        aggregated.append(frame)
                    else:
                        # Строка рассматривается как текст
                        full_text = " ".join([txt for txt in col_texts if txt])
                        if full_text:
                            aggregated.append(_pd.DataFrame({"text": [full_text]}))
        except Exception as ex:
            logger.error("Ошибка при извлечении (manual): %s", ex, exc_info=True)
            raise RuntimeError(f"Ошибка конвертации: {ex}")

    elif engine.lower() == "pymupdf":
        # Используем PyMuPDF для более точного извлечения таблиц без линий
        try:
            import fitz  # type: ignore
        except ImportError as ex:
            msg = (
                "Библиотека PyMuPDF (fitz) не установлена. Добавьте её в requirements.txt или "
                "установите вручную."
            )
            logger.error(msg)
            raise RuntimeError(msg) from ex
        try:
            doc = fitz.open(str(pdf_path))
            for page_index, page in enumerate(doc, start=1):
                try:
                    try:
                        tables = page.find_tables(
                            horizontal_strategy="text", vertical_strategy="text"
                        )
                    except Exception:
                        tables = page.find_tables()  # type: ignore[attr-defined]
                    if tables:
                        for tbl in tables:
                            try:
                                df = tbl.to_pandas()
                                norm = _normalize_dataframe(df)
                                aggregated.extend(norm)
                            except Exception as ex:
                                logger.error(
                                    "Ошибка обработки таблицы на странице %s: %s",
                                    page_index,
                                    ex,
                                    exc_info=True,
                                )
                    else:
                        try:
                            text = page.get_text("text")  # type: ignore[attr-defined]
                        except Exception:
                            text = ""
                        if text:
                            aggregated.append(pd.DataFrame({"text": [text]}))
                except Exception as ex:
                    logger.error(
                        "Ошибка обработки страницы %s: %s", page_index, ex, exc_info=True
                    )
                    try:
                        text = page.get_text("text")  # type: ignore[attr-defined]
                    except Exception:
                        text = ""
                    if text:
                        aggregated.append(pd.DataFrame({"text": [text]}))
        except Exception as ex:
            logger.error("Ошибка при чтении PDF (PyMuPDF): %s", ex, exc_info=True)
            raise RuntimeError(f"Ошибка конвертации: {ex}")
    elif engine.lower() == "regex":
        # Используем PyMuPDF и регулярные выражения для извлечения таблицы из строк без линий.
        # Каждая строка собирается из слов на странице; два последних значения с символом «₽»
        # определяются как «Цена за ед.» и «Сумма», остальные числа интерпретируются как
        # «Кол-во» и «Коэфф.». Строки без двух валютных значений пропускаются.
        try:
            import fitz  # type: ignore
        except ImportError as ex:
            msg = (
                "Библиотека PyMuPDF (fitz) не установлена. Добавьте её в requirements.txt или "
                "установите вручную."
            )
            logger.error(msg)
            raise RuntimeError(msg) from ex
        import re  # Импортируем стандартную библиотеку для регулярных выражений
        try:
            doc = fitz.open(str(pdf_path))
            # Паттерн для поиска сумм: число (с пробелами внутри и, возможно, десятичной частью) перед символом ₽
            currency_pattern = re.compile(r"(\d[\d\s]*\d(?:,\d{2})?)\s*₽")
            for page_index, page in enumerate(doc, start=1):
                try:
                    words = page.get_text("words")  # type: ignore[attr-defined]
                except Exception as ex:
                    logger.error(
                        "Ошибка извлечения слов на странице %s: %s", page_index, ex, exc_info=True
                    )
                    continue
                if not words:
                    continue
                # Определяем порог для группировки слов в строки
                heights = [w[3] - w[1] for w in words]
                avg_height = sum(heights) / len(heights) if heights else 0.0
                y_threshold = avg_height * 0.6 if avg_height else 2.0
                # Сортируем слова по вертикальной координате
                words_sorted = sorted(words, key=lambda w: w[1])
                rows: list[list] = []
                current_row: list = []
                current_y: Optional[float] = None
                for w in words_sorted:
                    y0 = w[1]
                    if current_y is None or abs(y0 - current_y) <= y_threshold:
                        current_row.append(w)
                        current_y = y0 if current_y is None else (current_y + y0) / 2
                    else:
                        if current_row:
                            current_row.sort(key=lambda x: x[0])
                            rows.append(current_row)
                        current_row = [w]
                        current_y = y0
                if current_row:
                    current_row.sort(key=lambda x: x[0])
                    rows.append(current_row)
                # Обрабатываем каждую строку
                for row_words in rows:
                    if not row_words:
                        continue
                    # Копируем слова и удаляем первый элемент, если это номер позиции
                    row_words_sorted = row_words[:]
                    if row_words_sorted and row_words_sorted[0][4].strip().isdigit():
                        row_words_sorted = row_words_sorted[1:]
                    if not row_words_sorted:
                        continue
                    # Собираем текст строки
                    row_text = " ".join([w[4] for w in row_words_sorted]).strip()
                    if not row_text:
                        continue
                    currency_matches = list(currency_pattern.finditer(row_text))
                    if len(currency_matches) < 2:
                        # Если меньше двух сумм (значений с ₽), пропускаем строку
                        continue
                    # Последние две суммы: предпоследняя – цена за единицу, последняя – итоговая сумма
                    unit_price_str = currency_matches[-2].group(1)
                    total_str = currency_matches[-1].group(1)
                    # Удаляем суммы из текста, чтобы не мешали поиску остальных чисел
                    without_currency = currency_pattern.sub("", row_text)
                    # Извлекаем числовые токены без знаков и пробелов
                    num_tokens = re.findall(r"\d+", without_currency)
                    # Если исходная строка имела числовой индекс, он уже удалён; подстрахуемся
                    if num_tokens and row_words and row_words[0][4].strip().isdigit():
                        num_tokens = num_tokens[1:]
                    # Определяем количество (qty) и коэффициент
                    if len(num_tokens) >= 2:
                        qty_str = num_tokens[-2]
                        coeff_str = num_tokens[-1]
                    elif len(num_tokens) == 1:
                        qty_str = num_tokens[0]
                        coeff_str = "1"
                    else:
                        # нет данных о количестве – пропускаем
                        continue
                    # Собираем название, исключая числовые элементы и символ ₽
                    name_tokens: list[str] = []
                    for w in row_words_sorted:
                        txt = w[4]
                        # пропускаем символы ₽ и явно валютные фрагменты
                        if "₽" in txt:
                            continue
                        # убираем пробелы и неразрывные пробелы для проверки числа
                        cleaned = (
                            txt.replace(" ", "")
                            .replace("\u00A0", "")
                            .replace("\u202F", "")
                        )
                        # пропускаем, если это полностью число
                        if cleaned.isdigit():
                            continue
                        # пропускаем, если это дробное число (например, 000,00)
                        try:
                            float(cleaned.replace(",", "."))
                            continue
                        except Exception:
                            pass
                        name_tokens.append(txt)
                    name = " ".join(name_tokens).strip()
                    if not name:
                        continue
                    try:
                        price_val = _to_number(unit_price_str)
                        qty_val = _to_number(qty_str)
                        coeff_val = _to_number(coeff_str)
                        total_val = _to_number(total_str)
                    except Exception:
                        price_val, qty_val, coeff_val, total_val = (
                            unit_price_str,
                            qty_str,
                            coeff_str,
                            total_str,
                        )
                    aggregated.append(
                        pd.DataFrame(
                            [
                                {
                                    "Оборудование": name,
                                    "Кол-во": qty_val,
                                    "Цена за ед.": price_val,
                                    "Коэфф.": coeff_val,
                                    "Сумма": total_val,
                                }
                            ]
                        )
                    )
        except Exception as ex:
            logger.error("Ошибка при чтении PDF (regex): %s", ex, exc_info=True)
            raise RuntimeError(f"Ошибка конвертации: {ex}")
    else:
        # Используем pdfplumber
        try:
            import pdfplumber  # type: ignore
        except ImportError as ex:
            msg = (
                "Библиотека pdfplumber не установлена. Добавьте её в requirements.txt или "
                "установите вручную."
            )
            logger.error(msg)
            raise RuntimeError(msg) from ex
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_index, page in enumerate(pdf.pages, start=1):
                    try:
                        tables = []
                        # пытаемся извлечь таблицы со стратегией text
                        try:
                            tables = page.extract_tables(
                                table_settings={
                                    "vertical_strategy": "text",
                                    "horizontal_strategy": "text",
                                }
                            ) or []
                        except TypeError:
                            # старая версия pdfplumber — вызываем без table_settings
                            tables = page.extract_tables() or []
                        except Exception:
                            tables = []
                        # если таблицы не найдены, пробуем extract_table (возвращает одну таблицу)
                        if not tables:
                            try:
                                t = page.extract_table(
                                    table_settings={
                                        "vertical_strategy": "text",
                                        "horizontal_strategy": "text",
                                    }
                                )
                                if t:
                                    tables = [t]
                            except Exception:
                                pass
                        if tables:
                            for table in tables:
                                if not table:
                                    continue
                                header, *rows = table
                                try:
                                    df = pd.DataFrame(rows, columns=header)
                                except Exception:
                                    df = pd.DataFrame(table)
                                norm = _normalize_dataframe(df)
                                aggregated.extend(norm)
                        else:
                            text = page.extract_text() or ""
                            if text:
                                aggregated.append(pd.DataFrame({"text": [text]}))
                    except Exception as ex:
                        logger.error(
                            "Ошибка обработки страницы %s: %s", page_index, ex, exc_info=True
                        )
                        text = page.extract_text() or ""
                        if text:
                            aggregated.append(pd.DataFrame({"text": [text]}))
        except Exception as ex:
            logger.error("Ошибка при чтении PDF: %s", ex, exc_info=True)
            raise RuntimeError(f"Ошибка конвертации: {ex}")

    # 2.5 Запись всех DataFrame в один лист Excel
    try:
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
            sheet_name = "Sheet1"
            start_row = 0
            for df in aggregated:
                try:
                    df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                except Exception:
                    # в случае ошибки преобразуем все значения в строки
                    df.astype(str).to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                start_row += len(df) + 1
    except Exception as ex:
        logger.error("Ошибка записи в Excel: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации: {ex}")


# 3. Построение вкладки «Конвертация»
def build_convert_tab(page: Any, tab: QtWidgets.QWidget) -> None:
    """
    Создаёт интерфейс вкладки «Конвертация».

    :param page: экземпляр ProjectPage, в котором размещаются виджеты
    :param tab: виджет вкладки, на котором разместится интерфейс
    """
    # Контейнер с вертикальной компоновкой
    v = QtWidgets.QVBoxLayout(tab)

    # 3.1 Заголовок/описание
    lbl_desc = QtWidgets.QLabel(
        "Выберите режим конвертации и перетащите PDF‑файл в поле ниже, чтобы конвертировать его в Excel.\n"
        "После перетаскивания выберите место для сохранения результирующего файла .xlsx."
    )
    lbl_desc.setWordWrap(True)
    v.addWidget(lbl_desc)

    # 3.1.1 Выбор режима конвертации
    # Добавляем выпадающий список с режимами: Rentman VSG и Rentman Jamteck.
    mode_layout = QtWidgets.QHBoxLayout()
    mode_label = QtWidgets.QLabel("Режим:")
    mode_combo = QtWidgets.QComboBox()
    # Используем данные (данные элементов) для хранения значения режима
    mode_combo.addItem("Rentman VSG", userData="vsg")
    mode_combo.addItem("Rentman Jamteck", userData="jamteck")
    mode_combo.setToolTip("Выберите алгоритм конвертации PDF")
    mode_layout.addWidget(mode_label)
    mode_layout.addWidget(mode_combo)
    # Помещаем компоновку в общий вертикальный контейнер
    v.addLayout(mode_layout)

    # 3.2 Класс виджета для D&D области
    class PdfDropFrame(QtWidgets.QFrame):
        """
        QFrame с поддержкой drag-and-drop для PDF. Приём нескольких файлов
        возможен, но диалог сохранения будет показан для каждого файла
        отдельно. Добавленные имена файлов отображаются в QLabel.
        """

        def __init__(self, title: str, get_mode: Any, parent: Optional[QtWidgets.QWidget] = None) -> None:
            """
            :param title: текст, отображаемый в области D&D
            :param get_mode: функция без параметров, возвращающая выбранный режим
            :param parent: родительский виджет
            """
            super().__init__(parent)
            self.setFrameShape(QtWidgets.QFrame.StyledPanel)
            self.setFrameShadow(QtWidgets.QFrame.Sunken)
            self.setAcceptDrops(True)
            self.setMinimumHeight(120)
            self.label = QtWidgets.QLabel(title, self)
            self.label.setAlignment(QtCore.Qt.AlignCenter)
            layout = QtWidgets.QVBoxLayout(self)
            layout.addWidget(self.label)
            self._get_mode = get_mode

        def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:  # type: ignore[override]
            # Принимаем только файлы
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
            else:
                event.ignore()

        def dropEvent(self, event: QtGui.QDropEvent) -> None:  # type: ignore[override]
            """
            Обрабатывает перетаскивание PDF‑файлов.

            3.3.1 Получаем список URL и фильтруем только PDF.
            3.3.2 Для каждого найденного файла определяем, куда сохранять
                    сконвертированный XLSX: если задан идентификатор проекта,
                    файл автоматически помещается в папку материалов проекта
                    (подкаталог ``Excel``). В противном случае выводится
                    диалог сохранения. Файлы с одинаковыми именами получают
                    числовой суффикс.
            3.3.3 Вызываем convert_pdf_to_excel и выводим сообщение об
                    успешной конвертации или ошибке.
            """
            try:
                urls = event.mimeData().urls()
                if not urls:
                    return
                for url in urls:
                    src_path = Path(url.toLocalFile())
                    # Пропускаем несуществующие файлы и не-PDF
                    if not src_path.exists() or src_path.suffix.lower() != ".pdf":
                        continue
                    try:
                        # 3.3.2.a Запрашиваем путь сохранения у пользователя всегда.
                        # Предлагаем имя исходного PDF с расширением .xlsx и
                        # автоматически подставляем папку загрузок пользователя.
                        suggested_name = src_path.with_suffix(".xlsx").name
                        # Определяем путь каталога загрузок по стандартным путям системы
                        download_dir = QtCore.QStandardPaths.writableLocation(
                            QtCore.QStandardPaths.DownloadLocation
                        )
                        initial_path = str(Path(download_dir) / suggested_name)
                        dest_name, _ = QtWidgets.QFileDialog.getSaveFileName(
                            page,
                            "Сохранить как...",
                            initial_path,
                            "Excel (*.xlsx)",
                        )
                        if not dest_name:
                            continue
                        dest_path = Path(dest_name)
                        # 3.3.2.b Выполняем конвертацию согласно выбранному режиму
                        mode = self._get_mode()
                        try:
                            if mode == "jamteck":
                                convert_pdf_to_excel_jamteck(src_path, dest_path)
                            else:
                                convert_pdf_to_excel(src_path, dest_path)
                            msg = f"Файл '{src_path.name}' конвертирован в '{dest_path.name}' (режим {mode})."
                        except Exception as ex:
                            # перехватываем ошибки, чтобы отобразить их в метке
                            err_msg = f"Ошибка конвертации {src_path.name}: {ex}"
                            self.label.setText(err_msg)
                            if hasattr(page, "_log") and callable(page._log):
                                page._log(err_msg, "error")
                            logger.error(err_msg, exc_info=True)
                            continue
                        # конвертация успешна
                        # Обновляем текст метки сообщением о конвертации
                        self.label.setText(msg)
                        # Выводим сообщение в пользовательский лог
                        if hasattr(page, "_log") and callable(page._log):
                            page._log(msg)
                        logger.info(msg)
                        # Запоминаем путь к последнему созданному Excel в родителе вкладки
                        try:
                            parent_tab = self.parent()
                            if parent_tab is not None:
                                # type: ignore[attr-defined]
                                setattr(parent_tab, "last_excel_path", dest_path)
                        except Exception:
                            # игнорируем возможные ошибки при присвоении
                            logger.debug("Не удалось сохранить путь последнего файла")
                    except Exception as ex:
                        # 3.3.3 Логируем ошибку и отображаем её
                        err_msg = f"Ошибка конвертации {src_path.name}: {ex}"
                        self.label.setText(err_msg)
                        if hasattr(page, "_log") and callable(page._log):
                            page._log(err_msg, "error")
                        logger.error(err_msg, exc_info=True)
            except Exception:
                logger.error("Ошибка обработки события drop", exc_info=True)

    # Создаём и добавляем виджет
    # Передаем функцию, возвращающую выбранный режим из выпадающего списка.
    drop_frame = PdfDropFrame(
        "Перетащите сюда PDF",
        get_mode=lambda: mode_combo.currentData(),
        parent=tab,
    )
    v.addWidget(drop_frame)

    # Храним путь к последнему успешно сконвертированному файлу
    # Он используется кнопкой «Отправить в импорт смет» для передачи файла в другой модуль.
    tab.last_excel_path: Optional[Path] = None  # type: ignore[assignment]

    # Функция отправки файла в импорт смет. Обёртка, чтобы иметь доступ к drop_frame и tab.
    def _send_to_import() -> None:
        """Отправляет последний конвертированный файл в модуль импорта смет."""
        last_path: Optional[Path] = getattr(tab, "last_excel_path", None)
        # Если файл ещё не был создан
        if not last_path:
            # Отображаем сообщение пользователю
            drop_frame.label.setText("Нет конвертированного файла для импорта.")
            logger.info("Попытка отправить в импорт смет без конвертированного файла")
            return
        # Если страница имеет специальный метод для отправки – вызываем его
        # Иначе просто логируем действие
        if hasattr(page, "send_to_import_smet") and callable(page.send_to_import_smet):
            try:
                page.send_to_import_smet(last_path)
                msg = f"Файл '{last_path.name}' отправлен в импорт смет."
                drop_frame.label.setText(msg)
                if hasattr(page, "_log") and callable(page._log):
                    page._log(msg)
                logger.info(msg)
            except Exception as ex:
                err_msg = f"Ошибка при отправке файла '{last_path.name}' в импорт смет: {ex}"
                drop_frame.label.setText(err_msg)
                if hasattr(page, "_log") and callable(page._log):
                    page._log(err_msg, "error")
                logger.error(err_msg, exc_info=True)
        else:
            # Метод отправки не реализован – выводим лог
            msg = f"Файл '{last_path.name}' готов для импорта смет (метод отправки не найден)."
            drop_frame.label.setText(msg)
            if hasattr(page, "_log") and callable(page._log):
                page._log(msg)
            logger.info(msg)

    # Кнопка для отправки последнего файла в импорт смет
    import_button = QtWidgets.QPushButton("Отправить в импорт смет")
    import_button.setToolTip("Перебросить последний сконвертированный Excel в модуль импортирования смет")
    import_button.clicked.connect(_send_to_import)
    v.addWidget(import_button)

    # Добавляем растяжку, чтобы поле располагалось сверху
    v.addStretch(1)
