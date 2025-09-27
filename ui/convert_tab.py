"""
Назначение:
    Этот модуль реализует вкладку «Конвертация» для приложения TechDirRentMan.
    Вкладка предоставляет простое drag‑and‑drop поле для загрузки PDF‑файлов
    и их конвертации в формат Excel. При перетаскивании файла пользователю
    предлагается выбрать путь сохранения полученного XLSX‑файла. Конвертация
    выполняется средствами Python: данные извлекаются из PDF с помощью
    библиотеки ``pdfplumber``, затем сохраняются в Excel через ``pandas``.
    В текущей версии все найденные таблицы и текстовые блоки собираются
    последовательно на один лист Excel — это избавляет от множества
    отдельных страниц в итоговом файле. После конвертации файл автоматически
    помещается в папку материалов выбранного проекта (подкаталог ``Excel``);
    если проект не выбран, пользователь может выбрать путь вручную. В процессе
    выгрузки числовые строки (в том числе с пробелами и запятыми в качестве
    разделителя) преобразуются в числа. Все события (успешные операции
    и ошибки) фиксируются в логах.

Принцип работы:
    • build_convert_tab(page, tab) создаёт интерфейс вкладки, включая
      область для перетаскивания и инструкции.
    • Класс PdfDropFrame обрабатывает события dragEnter/drop: по
      получению PDF‑файлов запрашивает у пользователя путь сохранения
      и инициирует конвертацию через ``pdfplumber`` и ``pandas``.
    • Конвертированные файлы сохраняются в указанный путь; сообщения
      об успешной конвертации выводятся в пользовательский лог.

Стиль:
    • Код разделён на пронумерованные секции с краткими заголовками.
    • Для ключевых действий даны поясняющие комментарии.
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
def convert_pdf_to_excel(pdf_path: Path, dest_path: Path) -> None:
    """
    Выполняет конвертацию PDF‑файла в XLSX, собирая все таблицы в один лист.

    Функция использует библиотеку `pdfplumber` для извлечения таблиц и текста
    из страниц PDF. Вместо создания отдельного листа для каждой страницы и
    таблицы, как было реализовано ранее, все найденные таблицы и текстовые
    блоки последовательно записываются на один лист Excel. Для удобства
    чтения между таблицами оставляется пустая строка. Числовые строки (с
    пробелами и запятыми) конвертируются в числа там, где это возможно.

    Требуемые зависимости: `pdfplumber`, `pandas`, `openpyxl` (см. requirements.txt).

    :param pdf_path: путь к исходному PDF
    :param dest_path: путь к итоговому XLSX
    :raises RuntimeError: при ошибке чтения или записи
    """
    # 2.1 Импорт внешних зависимостей
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
        import pandas as pd  # type: ignore
    except ImportError as ex:
        msg = (
            "Библиотека pandas не установлена. Добавьте её в requirements.txt или "
            "установите вручную."
        )
        logger.error(msg)
        raise RuntimeError(msg) from ex

    # 2.2 Вспомогательная функция для преобразования строк в числа
    def _to_number(val: Any) -> Any:
        """Пытается преобразовать строку в число, убирая пробелы и заменяя запятые на точки."""
        if isinstance(val, str):
            s = val.strip().replace(" ", "").replace(",", ".")
            try:
                num = float(s)
                return int(num) if num.is_integer() else num
            except Exception:
                return val
        return val

    # 2.3 Открываем PDF и собираем все таблицы/текст в список DataFrame
    aggregated: list = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_index, page in enumerate(pdf.pages, start=1):
                try:
                    tables = page.extract_tables() or []
                    if tables:
                        # На странице могут быть несколько таблиц; добавляем каждую отдельно
                        for table in tables:
                            header, *rows = table
                            df = pd.DataFrame(rows, columns=header)
                            # 2.2.a Применяем преобразование к каждой ячейке
                            df = df.applymap(_to_number)
                            # 2.2.b После ячейковой обработки пробуем привести столбцы к числовому типу
                            df = df.apply(lambda col: pd.to_numeric(col, errors='ignore'))
                            aggregated.append(df)
                    else:
                        # Если таблиц нет, сохраняем текст в одно поле
                        text = page.extract_text() or ""
                        aggregated.append(pd.DataFrame({"text": [text]}))
                except Exception as ex:
                    # Логируем ошибку и сохраняем текст страницы
                    logger.error(
                        "Ошибка обработки страницы %s: %s",
                        page_index,
                        ex,
                        exc_info=True,
                    )
                    text = page.extract_text() or ""
                    aggregated.append(pd.DataFrame({"text": [text]}))
    except Exception as ex:
        logger.error("Ошибка при чтении PDF: %s", ex, exc_info=True)
        raise RuntimeError(f"Ошибка конвертации: {ex}")

    # 2.4 Запись собранных данных в один лист Excel
    try:
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
            sheet_name = "Sheet1"
            start_row = 0
            for df in aggregated:
                # Записываем DataFrame начиная с текущей строки; индекс не сохраняем
                df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                # Обновляем строку: длина DataFrame + пустая строка между таблицами
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
                        # 3.3.2.b Выполняем конвертацию
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
