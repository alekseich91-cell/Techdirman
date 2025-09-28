"""
Назначение
===========

Этот модуль реализует вкладку «Конвертация» для приложения TechDirRentMan.
Она предоставляет простую drag‑and‑drop область для загрузки PDF‑файлов и
перевода их в формат Excel. Начиная с этой версии поддерживается только
один вариант конвертации — **режим Rentman VSG**. Он предназначен для
коммерческих предложений, сформированных системой Rentman/VSG, где таблицы
не имеют видимых линий, но следуют строгому порядку: номер позиции,
название, цена за единицу, количество, коэффициент, сумма. Прочие строки
(заголовки, сводные итоги) игнорируются. Строки, не содержащие всех
обязательных числовых полей (количество, коэффициент, сумма), исключаются.
Для раздела «Расходная часть», который обозначается в документе словом
«расход», коэффициент не указывается — он фиксируется равным 1, а строки
собираются из цены за единицу, количества и суммы.

При перетаскивании файла пользователю предлагается выбрать путь сохранения
результирующего XLSX‑файла. Файл сохраняется на указанный путь; все операции
и возможные ошибки отображаются в пользовательском логе и фиксируются в
лог‑файле ``convert_tab.log``.

Принцип работы
--------------

* ``build_convert_tab(page, tab)`` – создаёт интерфейс вкладки: текстовое
  описание и область для перетаскивания файла.
* ``PdfDropFrame`` – наследник ``QFrame`` с поддержкой drag‑and‑drop. При
  отпускании файла автоматически запрашивает путь сохранения и вызывает
  ``convert_pdf_to_excel``.
* ``convert_pdf_to_excel`` – извлекает данные из PDF, используя
  библиотеку **PyMuPDF**. Каждая строка страницы собирается из слов, после
  чего анализируется: если она содержит две суммы (значения с символом
  «₽»), алгоритм считает, что это строка таблицы и распределяет элементы
  между колонками: «Наименование», «Кол‑во», «Цена за ед.», «Коэфф.»,
  «Сумма». Строки без двух сумм трактуются как текст и сохраняются в
  отдельной колонке. Конечный результат записывается в один лист Excel.

Стиль кода
-----------

* Файл разделён на пронумерованные секции с краткими заголовками.
* Каждая функция снабжена комментарием, описывающим её назначение и
  логику работы.
* Все исключительные ситуации протоколируются в логах для упрощения
  отладки.
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


# 2. Вспомогательная функция: конвертация PDF → Excel
def convert_pdf_to_excel(pdf_path: Path, dest_path: Path, engine: str = "pdfplumber", *, manual_bounds: Optional[list[float]] = None) -> None:
    """
    Конвертирует PDF‑файл в XLSX, собирая все строки таблиц на одном листе.

    Несмотря на наличие параметров ``engine`` и ``manual_bounds`` (оставленных для
    обратной совместимости), функция всегда использует встроенный режим Rentman VSG.
    Этот режим анализирует каждую строку, извлечённую из PDF с помощью PyMuPDF,
    и ищет две суммы (значения с символом «₽»). Если строка относится к
    разделу аренды, она должна содержать как минимум количество и коэффициент;
    строки, в которых числовые поля отсутствуют, не попадают в таблицу.
    В разделе расходной части коэффициент отсутствует и принимается равным 1.
    Заголовочные и сводные строки пропускаются.

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
        cleaned = token.replace(" ", "").replace("\u00A0", "").replace("\u202F", "").replace(",", "")
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
                # Сбор суммы (последняя). Собираем все подряд идущие числовые токены
                # непосредственно перед знаком валюты. В некоторых PDF суммы могут
                # состоять из трёх частей (например, «16 000 00»), поэтому
                # ограничение на два токена приводит к потере значащих цифр.
                sum_tokens_list: list[str] = []
                idx = last_idx - 1
                while idx >= 0 and _is_numeric_token(tokens[idx]):
                    sum_tokens_list.insert(0, tokens[idx])
                    idx -= 1
                # Сбор цены (предпоследняя). Аналогично собираем все числовые токены
                # перед предыдущим знаком валюты.
                price_tokens_list: list[str] = []
                idx_p = second_last_idx - 1
                while idx_p >= 0 and _is_numeric_token(tokens[idx_p]):
                    price_tokens_list.insert(0, tokens[idx_p])
                    idx_p -= 1
                if not price_tokens_list or not sum_tokens_list:
                    continue
                # На основе найденных токенов суммы и цены заранее вычисляем их
                # числовые значения. Если указанная в документе сумма равна нулю,
                # то позиция считается скрытой и пропускается без дальнейшего
                # вычисления. Это позволяет не учитывать строки с «0 ₽» еще до
                # определения количества и коэффициента.
                # Помощник для объединения числовых токенов в строку
                def _join_number_tokens(num_tokens: list[str]) -> str:
                    if not num_tokens:
                        return ""
                    # Если последний токен из двух цифр и токенов >=3,
                    # интерпретируем как дробную часть
                    if len(num_tokens) >= 3 and len(num_tokens[-1]) == 2:
                        ints = num_tokens[:-1]
                        decimals = num_tokens[-1]
                        return " ".join(ints) + "," + decimals
                    return " ".join(num_tokens)

                try:
                    price_val_tmp = _to_number(_join_number_tokens(price_tokens_list))
                except Exception:
                    price_val_tmp = None
                try:
                    sum_val_extracted = _to_number(_join_number_tokens(sum_tokens_list))
                    if not isinstance(sum_val_extracted, (int, float)):
                        sum_val_extracted = None
                except Exception:
                    sum_val_extracted = None

                # Если сумма распознана и равна нулю — пропускаем строку, не
                # выполняя дальнейших вычислений
                if sum_val_extracted == 0:
                    continue
                # Чтение количества и коэффициента между ценой и суммой.
                qty_coeff_tokens: list[str] = [
                    t for t in tokens[second_last_idx + 1 : last_idx] if _is_numeric_token(t)
                ]
                qty_str: Optional[str] = None
                coeff_str: Optional[str] = None
                if not expenses_section:
                    # Для раздела аренды желательно иметь хотя бы один числовой токен
                    if not qty_coeff_tokens:
                        continue
                    # Формируем список кандидатов: пара (qty, coeff) и вариант с coeff=1.
                    candidates: list[tuple[str, str]] = []
                    if len(qty_coeff_tokens) >= 2:
                        candidates.append((qty_coeff_tokens[0], qty_coeff_tokens[1]))
                    # Вариант с коэффициентом 1 всегда добавляем.
                    candidates.append((qty_coeff_tokens[0], "1"))
                    # Перебираем кандидаты и выбираем тот, у которого произведение
                    # совпадает с указанной в документе суммой (если она распознана).
                    chosen: Optional[tuple[str, str]] = None
                    for cand_qty, cand_coeff in candidates:
                        try:
                            cand_qty_num = _to_number(cand_qty)
                            cand_coeff_num = _to_number(cand_coeff)
                            # убеждаемся, что цена тоже числовая
                            if not (
                                isinstance(cand_qty_num, (int, float))
                                and isinstance(cand_coeff_num, (int, float))
                                and isinstance(price_val_tmp, (int, float))
                            ):
                                continue
                            product = price_val_tmp * cand_qty_num * cand_coeff_num
                            if sum_val_extracted is not None:
                                if abs(product - sum_val_extracted) < 1:
                                    chosen = (cand_qty, cand_coeff)
                                    break
                            else:
                                # если сумма не распознана, берём первый валидный вариант
                                chosen = (cand_qty, cand_coeff)
                                break
                        except Exception:
                            continue
                    if chosen is None:
                        chosen = (qty_coeff_tokens[0], "1")
                    qty_str, coeff_str = chosen
                else:
                    # В расходной части коэффициент всегда 1, количество — первый токен
                    if not qty_coeff_tokens:
                        continue
                    qty_str = qty_coeff_tokens[0]
                    coeff_str = "1"
                # Имя – всё до начала цены
                price_start_index = second_last_idx - len(price_tokens_list)
                name_tokens = tokens[:price_start_index]
                name = " ".join(name_tokens).strip()
                if not name:
                    continue
                # Собираем строки суммы и цены. Если сумма или цена состоят из трёх
                # частей (например, «16 000 00»), преобразуем их в формат с запятой,
                # чтобы корректно распознать десятичные дроби.
                def _build_number_str(num_tokens: list[str]) -> str:
                    """Объединяет список числовых токенов в одну строку. Если
                    последняя часть имеет длину два символа и токенов три и более,
                    считаем, что это дробная часть, и вставляем запятую перед
                    ней (пример: ['16','000','00'] → '16 000,00'). В остальных
                    случаях просто объединяем через пробел.
                    """
                    if not num_tokens:
                        return ""
                    if len(num_tokens) >= 3 and len(num_tokens[-1]) == 2:
                        ints = num_tokens[:-1]
                        decimals = num_tokens[-1]
                        return " ".join(ints) + "," + decimals
                    return " ".join(num_tokens)

                unit_price_str = _build_number_str(price_tokens_list)
                total_str = _build_number_str(sum_tokens_list)

                try:
                    price_val = _to_number(unit_price_str)
                    qty_val = _to_number(qty_str)
                    coeff_val = _to_number(coeff_str)
                except Exception:
                    # В случае ошибки парсинга пропускаем строку, чтобы не
                    # придумывать значения для некорректных позиций.
                    continue

                # Пропускаем строки, где цена, количество или коэффициент не
                # распознались как числа. Это позволяет исключить позиции
                # без указания цены или количества.
                if not (
                    isinstance(price_val, (int, float))
                    and isinstance(qty_val, (int, float))
                    and isinstance(coeff_val, (int, float))
                ):
                    continue

                # Вычисляем итоговую сумму как произведение цены, количества и коэффициента.
                total_val = price_val * qty_val * coeff_val
                # Если итоговая сумма равна нулю, считаем, что цена скрыта или позиция не указана –
                # такую строку в таблицу не добавляем.
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
        df_result = pd.DataFrame(records, columns=["Наименование", "Кол-во", "Цена за ед.", "Коэфф.", "Сумма"])
        with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="Sheet1", index=False)
    except Exception as ex:
        logger.error("Ошибка записи в Excel: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации: {ex}")
    return

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
            logger.error("Некорректный список границ столбцов для режима manual: %s", bounds)
            raise RuntimeError("Некорректный список границ столбцов")
        try:
            doc = fitz.open(str(pdf_path))
            for page_index, page in enumerate(doc, start=1):
                try:
                    words = page.get_text("words")  # type: ignore[attr-defined]
                except Exception as ex:
                    logger.error("Ошибка извлечения слов на странице %s: %s", page_index, ex, exc_info=True)
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
                            frame = _pd.DataFrame([
                                {
                                    "Оборудование": name_val,
                                    "Кол-во": _to_number(qty_val),
                                    "Цена за ед.": _to_number(price_val),
                                    "Коэфф.": _to_number(coeff_val),
                                    "Сумма": _to_number(total_val),
                                }
                            ])
                        except Exception:
                            # fallback: сохраняем без преобразования
                            frame = _pd.DataFrame([
                                {
                                    "Оборудование": name_val,
                                    "Кол-во": qty_val,
                                    "Цена за ед.": price_val,
                                    "Коэфф.": coeff_val,
                                    "Сумма": total_val,
                                }
                            ])
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
                        tables = page.find_tables(horizontal_strategy="text", vertical_strategy="text")
                    except Exception:
                        tables = page.find_tables()  # type: ignore[attr-defined]
                    if tables:
                        for tbl in tables:
                            try:
                                df = tbl.to_pandas()
                                norm = _normalize_dataframe(df)
                                aggregated.extend(norm)
                            except Exception as ex:
                                logger.error("Ошибка обработки таблицы на странице %s: %s", page_index, ex, exc_info=True)
                    else:
                        try:
                            text = page.get_text("text")  # type: ignore[attr-defined]
                        except Exception:
                            text = ""
                        if text:
                            aggregated.append(pd.DataFrame({"text": [text]}))
                except Exception as ex:
                    logger.error("Ошибка обработки страницы %s: %s", page_index, ex, exc_info=True)
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
                        cleaned = txt.replace(" ", "").replace("\u00A0", "").replace("\u202F", "")
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
                        price_val, qty_val, coeff_val, total_val = unit_price_str, qty_str, coeff_str, total_str
                    aggregated.append(
                        pd.DataFrame([
                            {
                                "Оборудование": name,
                                "Кол-во": qty_val,
                                "Цена за ед.": price_val,
                                "Коэфф.": coeff_val,
                                "Сумма": total_val,
                            }
                        ])
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
                                table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}
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
                                    table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}
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
        "Перетащите PDF‑файл в поле ниже, чтобы конвертировать его в Excel.\n"
        "После перетаскивания выберите место для сохранения результирующего файла .xlsx."
    )
    lbl_desc.setWordWrap(True)
    v.addWidget(lbl_desc)

    # В текущей версии движок конвертации фиксирован (Rentman VSG),
    # поэтому выбор механизма и поле ручных границ удалены. Оставляем
    # только описание и область перетаскивания.

    # 3.2 Класс виджета для D&D области
    class PdfDropFrame(QtWidgets.QFrame):
        """
        QFrame с поддержкой drag-and-drop для PDF. Приём нескольких файлов
        возможен, но диалог сохранения будет показан для каждого файла
        отдельно. Добавленные имена файлов отображаются в QLabel.
        """
        def __init__(self, title: str, parent: Optional[QtWidgets.QWidget] = None) -> None:
            super().__init__(parent)
            self.setFrameShape(QtWidgets.QFrame.StyledPanel)
            self.setFrameShadow(QtWidgets.QFrame.Sunken)
            self.setAcceptDrops(True)
            self.setMinimumHeight(120)
            self.label = QtWidgets.QLabel(title, self)
            self.label.setAlignment(QtCore.Qt.AlignCenter)
            layout = QtWidgets.QVBoxLayout(self)
            layout.addWidget(self.label)

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
                        # Даже если проект выбран, пользователь самостоятельно
                        # выбирает место и имя итогового файла. Предлагаем
                        # исходное имя PDF с заменой расширения на .xlsx.
                        suggested = src_path.with_suffix(".xlsx").name
                        dest_name, ok = QtWidgets.QFileDialog.getSaveFileName(
                            page,
                            "Сохранить как...",
                            suggested,
                            "Excel (*.xlsx)"
                        )
                        if not ok or not dest_name:
                            continue
                        dest_path = Path(dest_name)
                        # 3.3.2.b Выполняем конвертацию. Используется режим Rentman VSG.
                        convert_pdf_to_excel(src_path, dest_path)
                        msg = f"Файл '{src_path.name}' конвертирован в '{dest_path.name}'."
                        # Обновляем текст метки сообщением о конвертации
                        self.label.setText(msg)
                        # Выводим сообщение в пользовательский лог
                        if hasattr(page, "_log") and callable(page._log):
                            page._log(msg)
                        logger.info(msg)
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
    drop_frame = PdfDropFrame("Перетащите сюда PDF", tab)
    v.addWidget(drop_frame)
    # Добавляем растяжку, чтобы поле располагалось сверху
    v.addStretch(1)
