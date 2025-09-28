"""
Модуль ``export_tab`` реализует вкладку «Экспорт в PDF».

Назначение:
    Предоставляет пользователю графический интерфейс для выбора типа
    экспортируемого документа (смета, погрузочная ведомость, финансовый
    отчёт, тайминг) и настройки параметров экспорта. По выбранным
    параметрам формирует PDF‑файл с использованием русских шрифтов.

    Вкладка синхронизируется с остальными частями проекта: данные
    «Сводной сметы» и «Бюджета» получаются через виджет
    ``page.tab_finance_widget`` или провайдер ``DBDataProvider``, а
    сведения для тайминга — из атрибутов, установленных модулем
    ``timing_tab`` (``timing_blocks``, ``timing_column_names``,
    ``timing_start_date`` и т.д.). Таким образом, отчёт всегда
    формируется на основе актуальных данных выбранного проекта.

Принцип работы:
    • ``build_export_tab(page, tab)`` — создаёт интерфейс вкладки. Вкладка
      состоит из двух страниц: «Параметры» и «Предпросмотр». В первой
      странице доступны настройки экспорта (выбор типа отчёта, опции
      сметы, погрузочной, финансового отчёта, тайминга, общие флаги).
      Кнопка «Сгенерировать PDF» формирует отчёт во временный файл и
      отображает его в «Предпросмотре», не запрашивая путь. Вторая
      страница показывает предпросмотр PDF и содержит кнопки
      «Обновить предпросмотр» и «Сохранить PDF». Кнопка сохранения
      запрашивает у пользователя имя файла и сохраняет PDF.
    • ``_DragDropFrame`` — вспомогательный класс, который реализует
      приём изображений путём перетаскивания (drag‑and‑drop). Сохраняет
      путь к выбранному изображению для последующей вставки в PDF.
    • ``generate_pdf(page, opts)`` — собирает данные из «Сводной сметы»,
      «Бюджета» и «Тайминга» в соответствии с выбранным типом и опциями,
      запрашивает у пользователя папку и имя файла, создаёт PDF‑файл
      посредством библиотеки ``reportlab``. Для шрифтов используются
      файлы TTF, расположенные в подкаталоге ``fonts``.

Формат PDF:
    Создаётся документ формата A4 в портретной ориентации с полями
    15 мм. Текст и таблицы выводятся на русском языке, заголовки
    выделяются жирным начертанием. При необходимости длинные строки
    автоматически переносятся. Ширина столбцов подстраивается в
    зависимости от содержимого.

Стиль:
    • Код разделён на пронумерованные секции с краткими заголовками.
    • Для каждой секции приведены короткие комментарии, поясняющие
      назначение и логику работы.
    • Все потенциально ошибочные участки оборачиваются в try/except с
      записью ошибки в лог.
"""

# 1. Импорт стандартных и внешних библиотек
from __future__ import annotations
import os
import logging
import json  # для чтения файлов снимков
from pathlib import Path  # для работы с путями
import shutil  # копирование файлов изображений
from typing import Dict, Any, Optional, Tuple, List, Set

from PySide6 import QtWidgets, QtCore, QtGui

# Дополнительные утилиты из проекта
from .common import CLASS_EN2RU, DATA_DIR, normalize_case, fmt_num, fmt_sign  # для перевода классов, путей, нормализации и форматирования

import textwrap  # для переноса длинных строк

# Импортируем необходимые компоненты из reportlab для генерации PDF.
# Пакет reportlab может отсутствовать в среде. В этом случае мы
# устанавливаем флаг REPORTLAB_AVAILABLE=False и не будем генерировать PDF.
try:
    from reportlab.lib.pagesizes import A4  # type: ignore
    from reportlab.lib import colors  # type: ignore
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle  # type: ignore
    from reportlab.lib.units import mm  # type: ignore
    from reportlab.pdfbase import pdfmetrics  # type: ignore
    from reportlab.pdfbase.ttfonts import TTFont  # type: ignore
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer  # type: ignore
    REPORTLAB_AVAILABLE = True
except Exception:
    # Если reportlab не установлен, устанавливаем флаг и логируем предупреждение
    REPORTLAB_AVAILABLE = False
    # Создаём заглушки для используемых символов, чтобы избежать ошибок
    A4 = None
    colors = None  # type: ignore
    getSampleStyleSheet = None  # type: ignore
    ParagraphStyle = object  # type: ignore
    mm = 1  # type: ignore
    pdfmetrics = None  # type: ignore
    TTFont = None  # type: ignore
    SimpleDocTemplate = None  # type: ignore
    Table = None  # type: ignore
    TableStyle = None  # type: ignore
    Paragraph = None  # type: ignore
    Spacer = None  # type: ignore
    logging.warning("Библиотека reportlab не установлена. Экспорт в PDF недоступен.")

# 2. Настройка логирования
LOG_DIR = os.path.join(os.path.dirname(__file__), "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "export_tab.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("export_tab")

# 3. Константы и пути
FONTS_DIR = os.path.join(os.path.dirname(__file__), "..", "fonts")
DEFAULT_LOGO_PATH = os.path.join(FONTS_DIR, "default_logo.png")  # если нужен логотип по умолчанию

# Ширина таблиц финансового отчёта (в миллиметрах). При альбомной ориентации
# доступная ширина листа A4 составляет 297 мм минус поля (по 15 мм с каждой
# стороны). Для таблиц финансового отчёта выбираем ширину 240 мм, чтобы
# максимально растянуть таблицы, но сохранить отступы.
FIN_TABLE_WIDTH_MM = 240.0

# Палитра цветов для групп (ReportLab). Используется для отображения групп
# в сметном отчёте при активации опции «отображать группировки». Цвета
# подобраны в соответствии с палитрой из summary_tab, но конвертированы
# в формат reportlab (RGB от 0 до 1).
# Расширенная палитра групповых цветов (ReportLab). Цвета соответствуют
# палитре в summary_tab и приведены к диапазону [0,1] для reportlab.
GROUP_COLORS_RL = [
    colors.Color(255/255.0, 102/255.0, 102/255.0),   # светло‑красный
    colors.Color(255/255.0, 153/255.0, 102/255.0),   # персиковый
    colors.Color(255/255.0, 204/255.0, 102/255.0),   # нежно‑оранжевый
    colors.Color(255/255.0, 255/255.0, 102/255.0),   # светло‑жёлтый
    colors.Color(204/255.0, 255/255.0, 102/255.0),   # салатовый
    colors.Color(153/255.0, 255/255.0, 102/255.0),   # лаймовый
    colors.Color(102/255.0, 255/255.0, 178/255.0),   # бирюзовый
    colors.Color(102/255.0, 255/255.0, 204/255.0),   # бирюзовый светлый
    colors.Color(102/255.0, 204/255.0, 255/255.0),   # небесно‑голубой
    colors.Color(102/255.0, 178/255.0, 255/255.0),   # голубой
    colors.Color(178/255.0, 102/255.0, 255/255.0),   # сиреневый
    colors.Color(204/255.0, 102/255.0, 255/255.0),   # фиолетовый
    colors.Color(255/255.0, 102/255.0, 178/255.0),   # розовый
    colors.Color(255/255.0, 102/255.0, 204/255.0),   # розово‑фиолетовый
]


def _register_fonts() -> None:
    """Регистрирует русские шрифты из папки ``fonts`` для использования в PDF.

    Если шрифты отсутствуют, запись в лог с предупреждением.
    """
    # Если reportlab недоступен, не регистрируем шрифты
    if not REPORTLAB_AVAILABLE:
        logger.warning("Невозможно зарегистрировать шрифты: библиотека reportlab отсутствует")
        return
    try:
        fonts = [
            ("DejaVuSans", "DejaVuSans.ttf"),
            ("DejaVuSans-Bold", "DejaVuSans-Bold.ttf"),
            ("DejaVuSans-Oblique", "DejaVuSans-Oblique.ttf"),
            ("DejaVuSans-BoldOblique", "DejaVuSans-BoldOblique.ttf"),
        ]
        for name, fname in fonts:
            path = os.path.join(FONTS_DIR, fname)
            if os.path.exists(path) and pdfmetrics:
                pdfmetrics.registerFont(TTFont(name, path))
            else:
                logger.warning("Шрифт %s не найден: %s", name, path)
        # Добавляем сопоставление стилей (bold/italic) для базового шрифта
        if pdfmetrics:
            pdfmetrics.registerFontFamily(
                "DejaVuSans",
                normal="DejaVuSans",
                bold="DejaVuSans-Bold",
                italic="DejaVuSans-Oblique",
                boldItalic="DejaVuSans-BoldOblique",
            )
    except Exception:
        logger.error("Ошибка регистрации шрифтов", exc_info=True)


class _DragDropFrame(QtWidgets.QFrame):
    """
    Виджет, принимающий изображение перетаскиванием (drag‑and‑drop).

    Пользователь может перетащить файл изображения в этот фрейм. После
    успешного приёма путь к файлу сохраняется в атрибуте ``image_path``.
    """
    def __init__(self, title: str = "Перетащите изображение сюда", parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.setAcceptDrops(True)
        self.setMinimumHeight(80)
        self._title = title
        self.image_path: Optional[str] = None
        self._label = QtWidgets.QLabel(title, self)
        self._label.setAlignment(QtCore.Qt.AlignCenter)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self._label)

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if os.path.isfile(path) and any(path.lower().endswith(ext) for ext in [".png", ".jpg", ".jpeg"]):
                self.image_path = path
                self._label.setText(os.path.basename(path))
                logger.info("Выбрано изображение: %s", path)
            else:
                QtWidgets.QMessageBox.warning(self, "Неверный файл", "Пожалуйста, перетащите изображение PNG или JPEG.")


def build_export_tab(page: Any, tab: QtWidgets.QWidget) -> None:
    """
    Строит графический интерфейс вкладки «Экспорт в PDF» с двумя
    подпунктами: «Параметры» и «Предпросмотр».

    Вкладка «Параметры» содержит все элементы для настройки экспорта:
        • выбор типа отчёта;
        • параметры для сметы, погрузочной, финансового отчёта и тайминга;
        • общие настройки (логотип, сравнение со снимком);
        • кнопку генерации PDF.

    Вкладка «Предпросмотр» пытается отобразить созданный PDF с помощью
    ``QtPdfWidgets.QPdfView``. Если компонент недоступен в используемой
    версии Qt, отображается сообщение о недоступности предпросмотра.

    :param page: объект ProjectPage для доступа к данным проекта
    :param tab: QWidget, контейнер для вкладки экспорта
    """
    # 4.1 Регистрация шрифтов один раз
    _register_fonts()

    # 4.2 Создаём виджет с вкладками: «Параметры» и «Предпросмотр»
    tab_widget = QtWidgets.QTabWidget()
    tab_layout = QtWidgets.QVBoxLayout(tab)
    tab_layout.addWidget(tab_widget)

    # 4.3 Первая вкладка: настройка параметров
    settings_widget = QtWidgets.QWidget()
    vbox = QtWidgets.QVBoxLayout(settings_widget)

    # 4.3 Выбор типа экспорта
    type_layout = QtWidgets.QHBoxLayout()
    type_layout.addWidget(QtWidgets.QLabel("Тип отчёта:"))
    cmb_type = QtWidgets.QComboBox()
    cmb_type.addItems(["Смета", "Погрузочная ведомость", "Финансовый отчёт", "Тайминг"])
    type_layout.addWidget(cmb_type)
    type_layout.addStretch(1)
    vbox.addLayout(type_layout)

    # 4.4 Стек страниц с параметрами (различные группы опций для каждого отчёта)
    stack = QtWidgets.QStackedWidget()
    vbox.addWidget(stack, 1)

    # 4.5 Опции для «Смета»
    smeta_widget = QtWidgets.QWidget()
    smeta_layout = QtWidgets.QVBoxLayout(smeta_widget)
    smeta_layout.setSpacing(4)
    # Флажки для сметы
    # Шапка всегда должна присутствовать для сметы, поэтому флажок скрыт и всегда активен
    chk_header = QtWidgets.QCheckBox("Добавить шапку с информацией о проекте")
    # По умолчанию устанавливаем флажок, но скрываем его из интерфейса
    chk_header.setChecked(True)
    chk_header.setVisible(False)
    chk_taxes = QtWidgets.QCheckBox("Отобразить с налогами")
    chk_discounts = QtWidgets.QCheckBox("Отобразить со скидками")
    # Новый флажок: отображать все цены с учётом налога
    chk_smeta_with_tax = QtWidgets.QCheckBox("Показать все цены с учётом налога")
    chk_sort_dept = QtWidgets.QCheckBox("Сортировать по отделам и выделять логистику/персонал/расходники")
    chk_sort_zones = QtWidgets.QCheckBox("Сортировать по зонам (с выделением логистики/персонала/расходников)")
    chk_header_img = QtWidgets.QCheckBox("Прикрепить картинку к шапке")
    frame_header_img = _DragDropFrame("Перетащите изображение для шапки")
    chk_zones_img = QtWidgets.QCheckBox("Прикрепить картинки к зонам")
    frame_zones_img = _DragDropFrame("Перетащите изображения для зон")

    # Новый флажок: отображать группировки (цветные группы как в сводной смете)
    chk_smeta_show_groups = QtWidgets.QCheckBox("Отображать группировки")
    # Опция: выводить смету только для выбранного подрядчика
    chk_smeta_vendor_only = QtWidgets.QCheckBox("Вывод по конкретному подрядчику")
    # Список подрядчиков для фильтрации сметы. Заполняется при создании вкладки.
    cmb_smeta_vendor = QtWidgets.QComboBox()
    cmb_smeta_vendor.setEnabled(False)
    # Функция обновления списка подрядчиков из данных проекта
    def refresh_smeta_vendor_list() -> None:
        """
        Заполняет выпадающий список подрядчиков, присутствующих в текущем проекте.

        Использует данные из вкладки «Бухгалтерия» или при отсутствии — из базы данных.
        Доступные подрядчики сортируются по имени.
        """
        try:
            cmb_smeta_vendor.clear()
            vendors: set[str] = set()
            # Получаем список подрядчиков из FinanceTab
            try:
                ft = getattr(page, "tab_finance_widget", None)
                if ft and hasattr(ft, "items"):
                    for it in ft.items:
                        v = getattr(it, "vendor", None)
                        if v:
                            vendors.add(str(v))
                else:
                    # Загружаем через провайдера, если доступно
                    if getattr(page, "db", None) and getattr(page, "project_id", None):
                        from .finance_tab import DBDataProvider  # type: ignore
                        prov = DBDataProvider(page)
                        items = prov.load_items() or []
                        for it in items:
                            v = getattr(it, "vendor", None)
                            if v:
                                vendors.add(str(v))
            except Exception:
                vendors = set()
            # Добавляем элементы в выпадающий список
            for v in sorted(vendors):
                cmb_smeta_vendor.addItem(v)
        except Exception:
            # Логируем и игнорируем ошибки при обновлении списка подрядчиков
            logger.error("Ошибка обновления списка подрядчиков для фильтра сметы", exc_info=True)
    # Первичное заполнение списка подрядчиков
    refresh_smeta_vendor_list()
    # При изменении флажка включаем/выключаем список
    def on_vendor_only_toggled(checked: bool) -> None:
        # При переключении режима фильтра включаем/выключаем список и обновляем его
        try:
            cmb_smeta_vendor.setEnabled(bool(checked))
            # Обновляем список подрядчиков каждый раз при включении фильтра
            if checked:
                if callable(refresh_smeta_vendor_list):
                    refresh_smeta_vendor_list()
        except Exception:
            logger.error("Ошибка при переключении фильтра подрядчика", exc_info=True)
    chk_smeta_vendor_only.toggled.connect(on_vendor_only_toggled)
    # Сохраняем виджеты для доступа в collect_options
    try:
        setattr(page, "chk_smeta_vendor_only", chk_smeta_vendor_only)
        setattr(page, "cmb_smeta_vendor", cmb_smeta_vendor)
    except Exception:
        logger.error("Не удалось сохранить элементы фильтра подрядчика в объекте страницы", exc_info=True)
    # Для переименования зоны «Без зоны» в экспорте добавим поле ввода
    # Пользователь может задать альтернативное отображаемое имя, которое будет
    # использоваться исключительно в PDF‑отчётах. По умолчанию — «Без зоны».
    lbl_no_zone_name = QtWidgets.QLabel("Название зоны без имени:")
    edt_no_zone_name = QtWidgets.QLineEdit()
    edt_no_zone_name.setPlaceholderText("Без зоны")
    edt_no_zone_name.setText("Без зоны")
    # Сохраняем виджет в объекте страницы для использования в других функциях
    try:
        setattr(page, "edt_no_zone_name", edt_no_zone_name)
    except Exception:
        logger.error("Не удалось сохранить поле переименования зоны в объекте страницы", exc_info=True)
    # Добавляем виджеты для сметы в интерфейс. Порядок важен: сначала
    # основные флаги, затем параметры изображений, далее настройка имени зоны,
    # после чего — фильтр по подрядчику и список подрядчиков.
    for w in (
        chk_header,
        chk_taxes,
        chk_discounts,
        chk_smeta_with_tax,
        chk_sort_dept,
        chk_sort_zones,
        chk_smeta_show_groups,
        chk_header_img,
        frame_header_img,
        chk_zones_img,
        frame_zones_img,
        lbl_no_zone_name,
        edt_no_zone_name,
        # -- 4.4.2 Фильтр по подрядчику (чекбокс и выпадающий список) --
        chk_smeta_vendor_only,
        cmb_smeta_vendor,
    ):
        smeta_layout.addWidget(w)
    smeta_layout.addStretch(1)
    stack.addWidget(smeta_widget)

    # 4.5.1 Загрузка сохранённых картинок экспорта (шапка, логотип, зоны)
    def _load_export_images() -> None:
        """
        Загружает ранее сохранённые изображения экспорта для текущего проекта
        (шапка, логотип, зоны) и отмечает соответствующие флажки.

        Изображения сохраняются в JSON-файле export_settings.json в
        папке assets/project_<id>. Каждая запись хранит путь к файлу.
        Если файл найден, устанавливается image_path у фрейма и
        активируется связанный чекбокс.
        """
        try:
            proj_id = getattr(page, "project_id", None)
            if not proj_id:
                return
            from .common import ASSETS_DIR  # type: ignore
            dest_dir = ASSETS_DIR / f"project_{proj_id}"
            settings_path = dest_dir / "export_settings.json"
            if not settings_path.exists():
                return
            with open(settings_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            # Загружаем шапку
            hdr = data.get("header_image")
            if hdr and os.path.exists(hdr):
                frame_header_img.image_path = hdr
                chk_header_img.setChecked(True)
                # Обновляем текст
                frame_header_img._label.setText(os.path.basename(hdr))
            # Загружаем логотип
            logo = data.get("logo_image")
            if logo and os.path.exists(logo):
                frame_logo_img.image_path = logo
                chk_logo.setChecked(True)
                frame_logo_img._label.setText(os.path.basename(logo))
            # Загружаем картинки зон
            zones = data.get("zones_images")
            if zones and os.path.exists(zones):
                frame_zones_img.image_path = zones
                chk_zones_img.setChecked(True)
                frame_zones_img._label.setText(os.path.basename(zones))
        except Exception:
            logger.error("Не удалось загрузить сохранённые изображения экспорта", exc_info=True)

    # Сразу пытаемся загрузить сохранённые изображения
    _load_export_images()

    def _save_export_images() -> None:
        """
        Сохраняет выбранные изображения экспорта (шапка, логотип, зоны) в
        папку проекта и записывает пути в JSON. Если изображение
        расположено вне каталога проекта, копируется в каталог
        assets/project_<id> и путь обновляется. Если изображение не
        выбрано (флажок снят), соответствующая запись становится None.
        """
        try:
            proj_id = getattr(page, "project_id", None)
            if not proj_id:
                return
            from .common import ASSETS_DIR  # type: ignore
            dest_dir = ASSETS_DIR / f"project_{proj_id}"
            dest_dir.mkdir(parents=True, exist_ok=True)
            settings: Dict[str, Any] = {}
            # Обработчик для копирования файла
            def handle_image(checked: bool, frame: _DragDropFrame, key: str, prefix: str) -> None:
                path = None
                if checked and frame.image_path:
                    src = Path(frame.image_path)
                    # Определяем расширение и цель
                    ext = src.suffix.lower()
                    dest_name = f"{prefix}{ext}"
                    dest_path = dest_dir / dest_name
                    try:
                        # Копируем, только если файл не совпадает
                        if not dest_path.exists() or src.resolve() != dest_path.resolve():
                            shutil.copy2(src, dest_path)
                        frame.image_path = str(dest_path)
                        # Обновляем текст
                        frame._label.setText(os.path.basename(dest_path))
                        path = str(dest_path)
                    except Exception:
                        logger.error("Не удалось скопировать изображение %s", src, exc_info=True)
                        path = str(src)
                settings[key] = path
            # Копируем шапку
            handle_image(chk_header_img.isChecked(), frame_header_img, "header_image", "export_header")
            # Копируем логотип
            handle_image(chk_logo.isChecked(), frame_logo_img, "logo_image", "export_logo")
            # Копируем картинки зон
            handle_image(chk_zones_img.isChecked(), frame_zones_img, "zones_images", "export_zones")
            # Записываем JSON
            settings_path = dest_dir / "export_settings.json"
            with open(settings_path, "w", encoding="utf-8") as fh:
                json.dump(settings, fh, ensure_ascii=False, indent=2)
        except Exception:
            logger.error("Не удалось сохранить изображения экспорта", exc_info=True)

    # 4.6 Опции для «Погрузочная ведомость»
    load_widget = QtWidgets.QWidget()
    load_layout = QtWidgets.QVBoxLayout(load_widget)
    load_layout.setSpacing(4)
    # Для простоты, только флажок шапки
    chk_load_header = QtWidgets.QCheckBox("Добавить шапку с информацией о проекте")
    load_layout.addWidget(chk_load_header)
    load_layout.addStretch(1)
    stack.addWidget(load_widget)

    # 4.7 Опции для «Финансовый отчёт»
    fin_widget = QtWidgets.QWidget()
    fin_layout = QtWidgets.QVBoxLayout(fin_widget)
    fin_layout.setSpacing(4)
    chk_fin_agents = QtWidgets.QCheckBox("Показать агентские комиссии")
    chk_fin_internal = QtWidgets.QCheckBox("Показать внутренние расчёты, доходы/расходы и прибыльность")
    # Новый флажок: только внутренние расчёты (показывает наши скидки, доходы и расходы)
    chk_fin_internal_only = QtWidgets.QCheckBox("Только внутренние расчёты")
    # Новый флажок: показывать все цены с учётом налога
    # Флажок «С налогами» — включает отображение налогов и сумм с налогом в отчёте
    chk_fin_with_tax = QtWidgets.QCheckBox("С налогами")
    # Новый флажок: показывать только итоговую таблицу по зонам
    chk_fin_zones_only = QtWidgets.QCheckBox("Только зоны")
    # Новый флажок: выводить только зоны и распределять суммы по подрядчикам
    # Этот режим показывает список зон и под каждой зоной выводит
    # распределение сумм между подрядчиками. Флажок активируется только
    # когда выбран режим «Только зоны».
    chk_fin_zones_by_vendor = QtWidgets.QCheckBox("Только зоны + распределить по подрядчикам")
    chk_fin_zones_by_vendor.setEnabled(False)
    fin_layout.addWidget(chk_fin_agents)
    fin_layout.addWidget(chk_fin_internal)
    fin_layout.addWidget(chk_fin_internal_only)
    fin_layout.addWidget(chk_fin_with_tax)
    fin_layout.addWidget(chk_fin_zones_only)
    fin_layout.addWidget(chk_fin_zones_by_vendor)
    # Новый флажок: специальный формат отчёта для Ксюши
    chk_fin_ksyusha = QtWidgets.QCheckBox("Отчёт для Ксюши")
    fin_layout.addWidget(chk_fin_ksyusha)
    # При выборе отчёта для Ксюши отключаем остальные параметры финансового отчёта,
    # поскольку они не применяются к данному формату. Если флажок снят, возвращаем
    # доступность остальных настроек.
    def _toggle_fin_ksyusha(state: bool) -> None:
        try:
            for w in (chk_fin_agents, chk_fin_internal, chk_fin_internal_only, chk_fin_with_tax, chk_fin_zones_only, chk_fin_zones_by_vendor):
                w.setEnabled(not state)
                if state:
                    w.setChecked(False)
        except Exception:
            pass
    chk_fin_ksyusha.toggled.connect(_toggle_fin_ksyusha)
    fin_layout.addStretch(1)
    stack.addWidget(fin_widget)

    # Связываем переключение режима «Только зоны» с доступностью флажка
    # «Распределить по подрядчикам»: при отключении «Только зоны»
    # флажок распределения отключается и его состояние сбрасывается.
    def _toggle_fin_zones_by_vendor(state: bool) -> None:
        try:
            chk_fin_zones_by_vendor.setEnabled(state)
            if not state:
                chk_fin_zones_by_vendor.setChecked(False)
        except Exception:
            pass
    chk_fin_zones_only.toggled.connect(_toggle_fin_zones_by_vendor)

    # 4.8 Опции для «Тайминг»
    timing_widget = QtWidgets.QWidget()
    timing_layout = QtWidgets.QVBoxLayout(timing_widget)
    timing_layout.setSpacing(4)
    # 4.8.1 Флажок: показывать таблицу тайминга (графический вариант 1)
    # По умолчанию выбран первый вариант графического тайминга
    chk_timing_show_table = QtWidgets.QCheckBox("Отобразить тайминг в виде таблицы")
    chk_timing_show_table.setChecked(True)
    # 4.8.2 Флажок: группировать по столбцам (каждый столбец отдельным блоком)
    chk_timing_per_column = QtWidgets.QCheckBox("Создавать тайминги по каждому столбцу отдельно")
    # 4.8.3 Флажок: второй вариант графического тайминга
    chk_timing_graphic2 = QtWidgets.QCheckBox("Второй вариант графического тайминга")
    # 4.8.4 Флажок: третий вариант графического тайминга (дерево времени)
    chk_timing_graphic3 = QtWidgets.QCheckBox("Графический тайминг 3 (дерево времени)")
    # Добавляем флажки в макет тайминга
    timing_layout.addWidget(chk_timing_show_table)
    timing_layout.addWidget(chk_timing_per_column)
    timing_layout.addWidget(chk_timing_graphic2)
    timing_layout.addWidget(chk_timing_graphic3)
    timing_layout.addStretch(1)
    stack.addWidget(timing_widget)

    # 4.9 Общие дополнительные опции
    common_frame = QtWidgets.QGroupBox("Дополнительные настройки")
    common_layout = QtWidgets.QVBoxLayout(common_frame)
    chk_logo = QtWidgets.QCheckBox("Добавить логотип на страницы")
    frame_logo_img = _DragDropFrame("Перетащите изображение логотипа")
    chk_compare = QtWidgets.QCheckBox("Использовать режим сравнения (снимок)")
    # Выпадающий список для выбора снимка
    cmb_snap = QtWidgets.QComboBox()
    cmb_snap.addItem("<Без сравнения>", None)

    def refresh_snapshots() -> None:
        """
        Обновляет список доступных снимков для текущего проекта.

        Снимки хранятся в каталоге ``assets/project_<id>/snapshots`` (см. summary_tab).
        Для каждого файла считывается поле ``name`` и добавляется в combobox
        вместе с путём к файлу. При отсутствии снимков список будет содержать
        только пункт «<Без сравнения>».
        """
        try:
            # Запоминаем текущий выбранный путь, чтобы восстановить выбор после обновления
            current_data = None
            try:
                idx = cmb_snap.currentIndex()
                current_data = cmb_snap.itemData(idx) if idx >= 0 else None
            except Exception:
                current_data = None
            cmb_snap.blockSignals(True)
            cmb_snap.clear()
            cmb_snap.addItem("<Без сравнения>", None)
            proj_id = getattr(page, "project_id", None)
            if proj_id is not None:
                from .common import ASSETS_DIR  # импорт здесь, чтобы избежать циклов
                snap_dir = Path(ASSETS_DIR) / f"project_{proj_id}" / "snapshots"
                if snap_dir.exists():
                    for f in sorted(snap_dir.glob(f"project_{proj_id}_*.json")):
                        try:
                            with open(f, "r", encoding="utf-8") as fh:
                                snap_data = json.load(fh)
                            name = snap_data.get("name") or f.stem
                            cmb_snap.addItem(name, str(f))
                        except Exception:
                            continue
            # Восстанавливаем выбор, если ранее выбранный путь присутствует в новом списке
            if current_data:
                for i in range(cmb_snap.count()):
                    try:
                        if str(cmb_snap.itemData(i)) == str(current_data):
                            cmb_snap.setCurrentIndex(i)
                            break
                    except Exception:
                        continue
        except Exception:
            logger.error("Ошибка обновления списка снимков", exc_info=True)
        finally:
            cmb_snap.blockSignals(False)

    # Немедленно заполняем список снимков (возможно пустой, если проект не выбран)
    refresh_snapshots()
    # Предоставляем метод страницы для обновления снимков при смене проекта
    try:
        setattr(page, "refresh_export_tab_snapshots", refresh_snapshots)
    except Exception:
        logger.error("Не удалось добавить метод refresh_export_tab_snapshots в объект страницы", exc_info=True)

    # Создаём элементы для пользовательской строки в оглавлении. Пользователь
    # может ввести любую подпись, которая будет отображаться в заголовке PDF.
    lbl_custom_title = QtWidgets.QLabel("Дополнительная строка в оглавление:")
    edt_custom_title = QtWidgets.QLineEdit()
    edt_custom_title.setPlaceholderText("Введите подпись или заголовок")
    # Добавляем виджеты в общую раскладку. Важен порядок: сначала логотип и
    # сравнение, затем снимок, а затем поле пользовательского заголовка.
    for w in (chk_logo, frame_logo_img, chk_compare, cmb_snap, lbl_custom_title, edt_custom_title):
        common_layout.addWidget(w)
    vbox.addWidget(common_frame)

    # 4.10 Кнопка генерации PDF
    btn_export = QtWidgets.QPushButton("Сгенерировать PDF")
    vbox.addWidget(btn_export, 0, QtCore.Qt.AlignRight)

    # 4.11 Статус
    lbl_status = QtWidgets.QLabel("")
    vbox.addWidget(lbl_status)

    # 4.12 Переключение страниц в зависимости от выбранного типа
    def on_type_changed(index: int) -> None:
        """
        Переключает видимую страницу в зависимости от выбранного типа отчёта.

        Дополнительно при выборе отчёта «Смета» обновляет список
        подрядчиков для фильтра, чтобы он отражал актуальные данные
        проекта. Это позволяет пользователю увидеть доступные
        подрядчики, даже если данные загружаются после инициализации
        вкладки.
        """
        try:
            stack.setCurrentIndex(index)
            # Обновляем список подрядчиков, когда выбрана смета
            try:
                # Используем название типа по индексу: элемент ComboBox
                current_type = cmb_type.itemText(index)
                if normalize_case(current_type) == normalize_case("Смета"):
                    # Обновляем список подрядчиков из актуальных данных
                    if callable(refresh_smeta_vendor_list):
                        refresh_smeta_vendor_list()
            except Exception:
                # Игнорируем ошибки обновления списка
                logger.error("Не удалось обновить список подрядчиков при переключении типа", exc_info=True)
        except Exception:
            logger.error("Ошибка переключения вкладок экспортного отчёта", exc_info=True)
    cmb_type.currentIndexChanged.connect(on_type_changed)

    # 4.13 Функция, собирающая текущие настройки экспорта
    def collect_options() -> Dict[str, Any]:
        """Читает значения элементов управления и формирует словарь опций."""
        report_type = cmb_type.currentText()
        # Определяем выбранный снимок: сохраняем путь
        snap_path = cmb_snap.currentData() if cmb_snap.currentIndex() > 0 else None
        # Определяем пользовательское имя для зоны без названия (используется в PDF)
        # Получаем текст из поля для зоны без имени. Если поле недоступно,
        # используем пустую строку, чтобы затем подставить значение по умолчанию.
        try:
            no_zone_label_val = edt_no_zone_name.text().strip()
        except Exception:
            no_zone_label_val = ""
        # Словарь опций, включающий переименование зоны без имени для сметы и финансов
        opts: Dict[str, Any] = {
            "report_type": report_type,
            # Опции сметы
            "smeta": {
                "add_header": chk_header.isChecked(),
                "show_taxes": chk_taxes.isChecked(),
                "show_discounts": chk_discounts.isChecked(),
                "with_tax": chk_smeta_with_tax.isChecked(),
                "sort_by_department": chk_sort_dept.isChecked(),
                "sort_by_zone": chk_sort_zones.isChecked(),
                "header_image": frame_header_img.image_path if chk_header_img.isChecked() else None,
                "zones_images": frame_zones_img.image_path if chk_zones_img.isChecked() else None,
                # Переименование зоны без имени
                "no_zone_label": no_zone_label_val,
                # --- 4.13.1 Фильтр по подрядчику для сметного отчёта ---
                # Флаг, активирующий вывод сметы только по одному выбранному подрядчику.
                # Если выставлен, в отчёт попадут только позиции данного подрядчика,
                # сгруппированные по зонам и классам.  При снятом флаге
                # выводятся все подрядчики проекта.  Значение выбранного
                # подрядчика передаётся отдельным полем ``vendor``.
                "vendor_only": chk_smeta_vendor_only.isChecked(),
                # Флаг отображения группировок: если True, позиции в смете окрашиваются
                # в цвета групп, аналогично сводной смете. Значение читается из
                # чекбокса chk_smeta_show_groups.
                "show_groups": chk_smeta_show_groups.isChecked() if 'chk_smeta_show_groups' in locals() else False,
                # Имя выбранного подрядчика (строка). Используется только если
                # ``vendor_only`` == True. Если поле пустое или None,
                # фильтрация по подрядчику не применяется.
                "vendor": (cmb_smeta_vendor.currentText().strip() if hasattr(cmb_smeta_vendor, "currentText") else ""),
            },
            # Опции погрузочной
            "load": {
                "add_header": chk_load_header.isChecked(),
            },
            # Опции финансового отчёта
            "fin": {
                "show_agents": chk_fin_agents.isChecked(),
                "show_internal": chk_fin_internal.isChecked(),
                # Внутренний отчёт: только наши расчёты (доходы/скидки/расходы)
                "internal_only": chk_fin_internal_only.isChecked(),
                "with_tax": chk_fin_with_tax.isChecked(),
                # Переименование зоны без имени
                "no_zone_label": no_zone_label_val,
                # Показывать только итоговую таблицу по зонам
                "zones_only": chk_fin_zones_only.isChecked(),
                # Режим: разбивать суммы по подрядчикам внутри каждой зоны.
                # Работает только при включённой опции "zones_only".
                "zones_by_vendor": chk_fin_zones_by_vendor.isChecked(),
                # Специальный отчёт для Ксюши: простой финансовый отчёт с
                # минимальными колонками. При его выборе другие опции
                # финансового отчёта отключаются.
                "for_ksyusha": chk_fin_ksyusha.isChecked(),
            },
            # Опции тайминга
            "timing": {
                # Первый графический вариант (табличный тайминг / горизонтальная шкала)
                "show_table": chk_timing_show_table.isChecked() if 'chk_timing_show_table' in locals() else True,
                # Создавать тайминги по каждому столбцу отдельно (список)
                "separate_columns": chk_timing_per_column.isChecked() if 'chk_timing_per_column' in locals() else False,
                # Второй графический вариант (горизонтальная шкала с колонками)
                "graphic2": chk_timing_graphic2.isChecked() if 'chk_timing_graphic2' in locals() else False,
                # Третий графический вариант (дерево времени)
                "graphic3": chk_timing_graphic3.isChecked() if 'chk_timing_graphic3' in locals() else False,
            },
            # Опции общего
            "common": {
                "logo": frame_logo_img.image_path if chk_logo.isChecked() else None,
                "compare_mode": chk_compare.isChecked(),
                # Передаём путь к файлу снимка, если выбран
                "snapshot": snap_path,
                # Пользовательская строка для оглавления. Если поле пустое, будет
                # проигнорировано при генерации PDF.
                # Читаем значение пользовательской строки. Если объект недоступен
                # (например, поле не создано), используем пустую строку, чтобы
                # при генерации PDF дополнительный заголовок не выводился.
                # Пользовательская строка для оглавления. Если чтение поля
                # вызывает исключение (например, поле ещё не создано),
                # возвращаем пустую строку.
                "custom_title": (  # type: ignore
                    (edt_custom_title.text().strip() if hasattr(edt_custom_title, "text") else "")
                ),
            },
        }
        return opts

    # 4.14 Обработчик кнопки экспорта
    def on_export_clicked() -> None:
        """
        Обработчик кнопки экспорта. Формирует словарь настроек и вызывает
        ``generate_pdf``. Предварительно обновляет список снимков и
        запрашивает у пользователя путь для сохранения файла. При успехе
        отображает путь в статусной строке.
        """
        try:
            # Обновляем список снимков перед чтением
            try:
                if callable(refresh_snapshots):
                    refresh_snapshots()
            except Exception:
                pass
            opts = collect_options()
            # Генерируем PDF во временный файл без запроса пути для предпросмотра
            out_path = generate_pdf(page, opts, target_path=None, ask_save=False)
            if out_path:
                lbl_status.setText("Предпросмотр создан")
                logger.info("Предпросмотр PDF успешно создан: %s", out_path)
                # Загружаем файл в предпросмотр, если доступен виджет
                try:
                    pdf_doc = getattr(page, "export_pdf_doc", None)
                    pdf_view = getattr(page, "export_pdf_view", None)
                    if pdf_doc and pdf_view:
                        pdf_doc.load(out_path)
                        # переключаем на вкладку предпросмотра
                        tab_widget.setCurrentIndex(1)
                except Exception:
                    logger.error("Ошибка отображения предпросмотра", exc_info=True)
            else:
                lbl_status.setText("Ошибка создания предпросмотра")
        except Exception:
            logger.error("Ошибка генерации PDF", exc_info=True)
            QtWidgets.QMessageBox.critical(tab, "Ошибка", "Не удалось создать PDF. Подробности в логах.")

    btn_export.clicked.connect(on_export_clicked)

    # 4.15 Обработчик кнопки обновления предпросмотра
    def on_refresh_preview() -> None:
        """
        Обновляет предпросмотр PDF без запроса имени файла. Создаёт
        временный файл и формирует отчёт на основе текущих настроек.
        """
        try:
            # Обновляем список снимков перед чтением
            try:
                if callable(refresh_snapshots):
                    refresh_snapshots()
            except Exception:
                pass
            opts = collect_options()
            # Генерируем PDF во временный файл без запроса пути
            out_path = generate_pdf(page, opts, target_path=None, ask_save=False)
            if out_path:
                lbl_status.setText("Предпросмотр обновлён")
                # Загружаем файл в предпросмотр
                try:
                    pdf_doc = getattr(page, "export_pdf_doc", None)
                    pdf_view = getattr(page, "export_pdf_view", None)
                    if pdf_doc and pdf_view:
                        pdf_doc.load(out_path)
                        tab_widget.setCurrentIndex(1)
                except Exception:
                    logger.error("Ошибка обновления предпросмотра", exc_info=True)
            else:
                lbl_status.setText("Ошибка создания предпросмотра")
        except Exception:
            logger.error("Ошибка при обновлении предпросмотра", exc_info=True)
    # Подключаем кнопку, если она есть
    try:
        if getattr(page, "btn_refresh_preview", None):
            page.btn_refresh_preview.clicked.connect(on_refresh_preview)  # type: ignore
    except Exception:
        logger.error("Не удалось подключить кнопку обновления предпросмотра", exc_info=True)

    # 4.16 Обработчик кнопки сохранения PDF в предпросмотре
    def on_save_pdf() -> None:
        """
        Открывает диалог сохранения и формирует PDF в выбранный файл на
        основе текущих настроек. После сохранения при желании можно
        обновить предпросмотр.
        """
        try:
            # Обновляем список снимков
            try:
                if callable(refresh_snapshots):
                    refresh_snapshots()
            except Exception:
                pass
            opts = collect_options()
            # Перед сохранением копируем выбранные изображения в папку проекта и сохраняем настройки
            _save_export_images()
            # Генерируем PDF с запросом пути сохранения
            out_path = generate_pdf(page, opts, target_path=None, ask_save=True)
            if out_path:
                lbl_status.setText(f"Файл сохранён: {out_path}")
                logger.info("PDF сохранён по запросу: %s", out_path)
                # Загружаем сохранённый файл в предпросмотр
                try:
                    pdf_doc = getattr(page, "export_pdf_doc", None)
                    pdf_view = getattr(page, "export_pdf_view", None)
                    if pdf_doc and pdf_view:
                        pdf_doc.load(out_path)
                        tab_widget.setCurrentIndex(1)
                except Exception:
                    logger.error("Ошибка обновления предпросмотра", exc_info=True)
            else:
                lbl_status.setText("Файл не был сохранён")
        except Exception:
            logger.error("Ошибка сохранения PDF", exc_info=True)
            QtWidgets.QMessageBox.critical(tab, "Ошибка", "Не удалось сохранить PDF. Подробности в логах.")
    # Подключаем кнопку сохранения, если она есть
    try:
        if getattr(page, "btn_save_pdf", None):
            page.btn_save_pdf.clicked.connect(on_save_pdf)  # type: ignore
    except Exception:
        logger.error("Не удалось подключить кнопку сохранения PDF", exc_info=True)

    # 4.14 Если reportlab недоступен, отключаем функциональность экспорта
    if not REPORTLAB_AVAILABLE:
        lbl_status.setText("Для экспорта необходимо установить библиотеку reportlab")
        btn_export.setEnabled(False)
        # Так как библиотека недоступна, завершаем настройку вкладки
        # Добавляем вкладки перед возвратом
        # Создаём предпросмотр (ничего не будет отображаться)
        preview_widget = QtWidgets.QWidget()
        prev_layout = QtWidgets.QVBoxLayout(preview_widget)
        prev_layout.addWidget(QtWidgets.QLabel("Предпросмотр недоступен"))
        # Добавляем вкладки в виджет
        tab_widget.addTab(settings_widget, "Параметры")
        tab_widget.addTab(preview_widget, "Предпросмотр")
        return

    # 4.14 Первоначальное отображение первой страницы
    stack.setCurrentIndex(0)

    # 4.15 Создаём вкладку предпросмотра PDF
    preview_widget = QtWidgets.QWidget()
    prev_layout = QtWidgets.QVBoxLayout(preview_widget)

    # Пытаемся создать виджет просмотра PDF
    try:
        from PySide6.QtPdf import QPdfDocument  # type: ignore
        from PySide6.QtPdfWidgets import QPdfView  # type: ignore
        pdf_doc = QPdfDocument(preview_widget)
        pdf_view = QPdfView(preview_widget)
        pdf_view.setDocument(pdf_doc)
        # Кнопка обновления предпросмотра (сгенерировать для просмотра)
        btn_refresh_preview = QtWidgets.QPushButton("Обновить предпросмотр")
        # Кнопка сохранения PDF
        # Сохраняет текущий отчёт в файл, предлагая выбрать путь сохранения.
        # Надпись явно указывает, что происходит сохранение PDF.
        btn_save_pdf = QtWidgets.QPushButton("Сохранить PDF")
        # Горизонтальный контейнер для кнопок предпросмотра
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.addWidget(btn_refresh_preview)
        btn_layout.addWidget(btn_save_pdf)
        btn_layout.addStretch(1)
        prev_layout.addLayout(btn_layout)
        prev_layout.addWidget(pdf_view)
        # Сохраняем ссылки в объекте страницы
        page.export_pdf_doc = pdf_doc  # type: ignore
        page.export_pdf_view = pdf_view  # type: ignore
        page.btn_refresh_preview = btn_refresh_preview  # type: ignore
        page.btn_save_pdf = btn_save_pdf  # type: ignore
        # Подключаем кнопку сохранения к обработчику здесь. Ранее подключение
        # выполнялось до создания новой кнопки, из-за чего новое подключение не
        # происходило, и кнопка оставалась без действия. Теперь мы явно
        # подключаем её к обработчику on_save_pdf.
        try:
            btn_save_pdf.clicked.connect(on_save_pdf)  # type: ignore
        except Exception:
            logger.error("Не удалось подключить кнопку сохранения PDF", exc_info=True)
        # Настраиваем режим отображения PDF: если поддерживается, показываем все
        # страницы подряд, чтобы пользователь мог пролистывать документ в
        # предпросмотре. PySide6 может не поддерживать эту функцию в старых
        # версиях, поэтому оборачиваем вызов в try.
        try:
            # В новых версиях PySide6 перечисление PageMode определено на классе
            # QPdfView. Задаём режим MultiPage, если он доступен.
            pdf_view.setPageMode(QPdfView.PageMode.MultiPage)  # type: ignore[attr-defined]
        except Exception:
            # На случай, если PageMode не найден, пробуем доступ через
            # экземпляр. Если и это не получится, просто игнорируем.
            try:
                pdf_view.setPageMode(pdf_view.PageMode.MultiPage)  # type: ignore[attr-defined]
            except Exception:
                pass
    except Exception:
        prev_layout.addWidget(QtWidgets.QLabel("Предпросмотр PDF недоступен в этой версии Qt"))
        page.export_pdf_doc = None  # type: ignore
        page.export_pdf_view = None  # type: ignore
        page.btn_refresh_preview = None  # type: ignore
        page.btn_save_pdf = None  # type: ignore

    # Добавляем вкладки в общий TabWidget
    tab_widget.addTab(settings_widget, "Параметры")
    tab_widget.addTab(preview_widget, "Предпросмотр")

    logger.info("Вкладка 'Экспорт в PDF' успешно создана")


def generate_pdf(
    page: Any,
    options: Dict[str, Any],
    *,
    target_path: Optional[str] = None,
    ask_save: bool = True,
) -> Optional[str]:
    """
    Собирает данные проекта и формирует PDF‑отчёт в зависимости от выбранного типа и опций.

    Поведение контролируется параметрами ``target_path`` и ``ask_save``:
        • Если ``ask_save`` равно ``True``, у пользователя запрашивается путь
          для сохранения PDF через диалог ``QFileDialog.getSaveFileName``.
          Параметр ``target_path`` при этом игнорируется. Если пользователь
          отменяет выбор, генерация не выполняется и возвращается ``None``.
        • Если ``ask_save`` равно ``False``, отчёт создаётся во временном
          файле (или по пути ``target_path``, если он указан), без
          отображения диалога. Это используется для предпросмотра.

    Данные сметы берутся из ``page.tab_finance_widget`` либо провайдера
    ``DBDataProvider``; данные тайминга — из атрибутов, установленных
    ``timing_tab``. При включённом режиме сравнения использует JSON‑снимок
    из ``options['common']['snapshot']`` для подсветки изменений.

    :param page: объект ProjectPage, содержащий данные проекта
    :param options: словарь настроек, собранный из интерфейса
    :param target_path: путь для сохранения файла; если не задан и
      ``ask_save`` равно ``True``, путь спрашивается у пользователя,
      иначе создаётся временный файл
    :param ask_save: если ``True``, вызывает диалог сохранения; иначе
      генерирует отчёт в ``target_path`` или временный файл
    :return: путь к созданному PDF‑файлу или None при ошибке/отказе
    """
    # Проверяем доступность reportlab
    if not REPORTLAB_AVAILABLE:
        logger.error("Библиотека reportlab отсутствует. PDF не будет создан.")
        return None
    try:
        report_type = options.get("report_type", "").strip() or "Отчёт"

        # Определяем путь сохранения. Если ask_save=True и путь не задан, спрашиваем пользователя.
        if not target_path:
            # Формируем имя файла по умолчанию
            default_name = f"{report_type.lower().replace(' ', '_')}.pdf"
            if ask_save:
                try:
                    file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                        page if isinstance(page, QtWidgets.QWidget) else None,
                        "Сохранить отчёт как...",
                        default_name,
                        "PDF Files (*.pdf)"
                    )
                    if not file_path:
                        # Пользователь отменил выбор
                        return None
                    if not file_path.lower().endswith(".pdf"):
                        file_path += ".pdf"
                    target_path = file_path
                except Exception:
                    logger.error("Не удалось запросить путь сохранения файла", exc_info=True)
                    return None
            else:
                # Без запроса и без указанного пути использовать временное имя
                import tempfile
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                target_path = temp_file.name
                temp_file.close()

        file_path = target_path

        # Подготовка документа. По умолчанию используем формат A4 (портрет).
        # В зависимости от типа отчёта и выбранных опций меняем ориентацию.
        from reportlab.lib.pagesizes import landscape
        page_size = A4  # базовый размер
        try:
            # 1. Для отчёта «Смета» всегда используем альбомную ориентацию, чтобы
            #    таблицы можно было растянуть по ширине.
            if report_type == "Смета":
                page_size = landscape(A4)
            # 2. Для тайминга выбираем ориентацию по вариантам:
            elif report_type == "Тайминг":
                timing_opts = options.get("timing", {}) if isinstance(options, dict) else {}
                # Отдельные таблицы по колонкам (separate_columns) выводятся
                # в портретной ориентации, так как вертикальная таблица.
                if timing_opts.get("separate_columns", False):
                    page_size = A4
                # Дерево тайминга (graphic3) — всегда портрет.
                elif timing_opts.get("graphic3", False):
                    page_size = A4
                else:
                    # Остальные варианты тайминга (графические 1 и 2) используют
                    # альбомный формат для большей ширины.
                    page_size = landscape(A4)
            # 3. Для финансового отчёта и погрузочной ведомости также используем
            #    альбомную ориентацию, чтобы вместить широкие таблицы.
            elif report_type in ("Финансовый отчёт", "Погрузочная ведомость"):
                page_size = landscape(A4)
        except Exception:
            pass
        doc = SimpleDocTemplate(
            file_path,
            pagesize=page_size,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=20 * mm,
            bottomMargin=20 * mm,
        )
        elements: List[Any] = []
        styles = getSampleStyleSheet()
        # Стили с русскими шрифтами
        normal_style = styles["Normal"].clone("ruNormal")
        normal_style.fontName = "DejaVuSans"
        header_style = styles["Heading1"].clone("ruHeader")
        header_style.fontName = "DejaVuSans-Bold"
        # Устанавливаем крупный размер шрифта и межстрочный интервал для заголовков.
        # Это обеспечивает, что пользовательская строка и основной заголовок будут одного размера.
        header_style.fontSize = 18
        header_style.leading = 22

        # Добавляем логотип, если указан
        common_opts = options.get("common", {})
        logo_path = common_opts.get("logo")
        if logo_path and os.path.exists(logo_path):
            try:
                from reportlab.platypus import Image
                img = Image(logo_path, width=40 * mm, height=15 * mm)
                img.hAlign = "RIGHT"
                elements.append(img)
                elements.append(Spacer(1, 5 * mm))
            except Exception:
                logger.error("Ошибка вставки логотипа", exc_info=True)

        # Заголовок
        elements.append(Paragraph(report_type, header_style))
        # Вставляем пользовательскую строку в оглавление (если указана)
        try:
            custom_title: Optional[str] = None
            if isinstance(common_opts, dict):
                custom_title = common_opts.get("custom_title") or None
            if custom_title:
                # Добавляем пользовательский заголовок отдельным параграфом.
                # Используем тот же стиль, что и для основного заголовка, чтобы размеры совпадали.
                elements.append(Paragraph(str(custom_title), header_style))
        except Exception:
            # Не прерываем создание PDF, просто логируем ошибку
            logger.error("Ошибка добавления пользовательского заголовка", exc_info=True)
        elements.append(Spacer(1, 4 * mm))

        # Подготовка данных для отчёта.
        # Формируем карту различий между текущими элементами и выбранным снимком (diff_map).
        # Эта карта будет использоваться для расчёта дельт и подсветки строк в отчёте.
        diff_map: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
        compare_enabled = bool(common_opts.get("compare_mode") and common_opts.get("snapshot"))
        # Данные снимка для финансового отчёта (суммы по подрядчикам/зонам/отделам/классам)
        fin_snapshot_data: Optional[Dict[str, Any]] = None
        snap_map: Dict[Tuple[str, str, str, str], Tuple[float, float, float, float, str]] = {}
        if compare_enabled:
            snap_path: Optional[str] = common_opts.get("snapshot")
            try:
                if snap_path and os.path.exists(snap_path):
                    with open(snap_path, "r", encoding="utf-8") as f:
                        snap_data = json.load(f)
                    # Извлекаем снимок финансового отчёта, если он есть
                    fin_snapshot_data = snap_data.get("fin_snapshot") if isinstance(snap_data, dict) else None
                    # Считываем элементы снимка. Ключи могут быть строковыми, приводим к int
                    items_raw: Dict[Any, Any] = snap_data.get("items", {}) or {}
                    snap_items: Dict[int, Any] = {}
                    for k, v in items_raw.items():
                        try:
                            snap_items[int(k)] = v
                        except Exception:
                            continue
                    # Строим карту snap_map по ключу (подрядчик, наименование, отдел, зона)
                    for itm in snap_items.values():
                        try:
                            v_key = normalize_case(itm.get("vendor", ""))
                            n_key = normalize_case(itm.get("name", ""))
                            d_key = normalize_case(itm.get("department", ""))
                            z_key = normalize_case(itm.get("zone", ""))
                            qty_s = float(itm.get("qty", 0.0))
                            coeff_s = float(itm.get("coeff", 0.0))
                            price_s = float(itm.get("unit_price", 0.0))
                            amount_s = qty_s * coeff_s * price_s
                            cls_s = itm.get("class", "equipment")
                            snap_map[(v_key, n_key, d_key, z_key)] = (
                                qty_s, coeff_s, price_s, amount_s, cls_s
                            )
                        except Exception:
                            continue
                else:
                    logger.warning("Файл снимка не найден: %s", snap_path)
                    compare_enabled = False
            except Exception:
                logger.error("Ошибка обработки режима сравнения", exc_info=True)
                compare_enabled = False
        # Получаем текущие позиции
        current_items: List[Any] = []
        try:
            ft = getattr(page, "tab_finance_widget", None)
            if ft and hasattr(ft, "items"):
                current_items = ft.items
            else:
                if getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    current_items = prov.load_items() or []
        except Exception:
            current_items = []
        # Собираем diff_map для всех текущих элементов
        current_keys: Set[Tuple[str, str, str, str]] = set()
        for it in current_items:
            try:
                v_key = normalize_case(getattr(it, "vendor", ""))
                n_key = normalize_case(getattr(it, "name", ""))
                d_key = normalize_case(getattr(it, "department", ""))
                z_key = normalize_case(getattr(it, "zone", ""))
                key = (v_key, n_key, d_key, z_key)
                current_keys.add(key)
                qty_c = float(getattr(it, "qty", 0.0))
                coeff_c = float(getattr(it, "coeff", 0.0))
                price_c = float(getattr(it, "price", 0.0))
                amount_c = qty_c * coeff_c * price_c
                cls_en = getattr(it, "cls", "equipment")
                if compare_enabled and key in snap_map:
                    qty_s, coeff_s, price_s, amount_s, cls_s = snap_map[key]
                    diff_qty = qty_c - qty_s
                    diff_coeff = coeff_c - coeff_s
                    diff_price = price_c - price_s
                    diff_amount = amount_c - amount_s
                    # Состояние изменения
                    if abs(diff_qty) < 1e-6 and abs(diff_price) < 1e-6 and abs(diff_coeff) < 1e-6:
                        state = "не изменилось"
                    else:
                        state = "изменилось"
                    diff_map[key] = {
                        "state": state,
                        "diff_qty": diff_qty,
                        "diff_coeff": diff_coeff,
                        "diff_price": diff_price,
                        "diff_amount": diff_amount,
                        "snap_qty": qty_s,
                        "snap_coeff": coeff_s,
                        "snap_price": price_s,
                        "snap_amount": amount_s,
                        "class": cls_s,
                    }
                elif compare_enabled:
                    # новая строка
                    diff_map[key] = {
                        "state": "добавлено",
                        "diff_qty": qty_c,
                        "diff_coeff": coeff_c,
                        "diff_price": price_c,
                        "diff_amount": amount_c,
                        "snap_qty": 0.0,
                        "snap_coeff": 0.0,
                        "snap_price": 0.0,
                        "snap_amount": 0.0,
                        "class": cls_en,
                    }
                else:
                    # если сравнение отключено, просто отмечаем отсутствие изменений
                    diff_map[key] = {
                        "state": "",
                        "diff_qty": 0.0,
                        "diff_coeff": 0.0,
                        "diff_price": 0.0,
                        "diff_amount": 0.0,
                        "snap_qty": 0.0,
                        "snap_coeff": 0.0,
                        "snap_price": 0.0,
                        "snap_amount": 0.0,
                        "class": cls_en,
                    }
            except Exception:
                continue
        # Добавляем удалённые элементы в diff_map, если сравнение включено
        if compare_enabled:
            for key, vals in snap_map.items():
                if key not in current_keys:
                    qty_s, coeff_s, price_s, amount_s, cls_s = vals
                    diff_map[key] = {
                        "state": "удалено",
                        "diff_qty": -qty_s,
                        "diff_coeff": -coeff_s,
                        "diff_price": -price_s,
                        "diff_amount": -amount_s,
                        "snap_qty": qty_s,
                        "snap_coeff": coeff_s,
                        "snap_price": price_s,
                        "snap_amount": amount_s,
                        "class": cls_s,
                    }

        # В зависимости от выбранного отчёта формируем содержимое
        if report_type == "Смета":
            # Обновляем опции сметы локально, добавляя карту налогов, если требуется
            smeta_opts = dict(options.get("smeta", {}))
            if smeta_opts.get("with_tax", False):
                # Формируем карту налоговых коэффициентов по подрядчикам
                vendor_tax: Dict[str, float] = {}
                try:
                    ft = getattr(page, "tab_finance_widget", None)
                    if ft and hasattr(ft, "preview_tax_pct"):
                        for v in ft.preview_tax_pct:
                            try:
                                vendor_tax[normalize_case(v)] = float(ft.preview_tax_pct.get(v, 0.0)) / 100.0
                            except Exception:
                                continue
                    else:
                        # Попытка загрузить из провайдера
                        if getattr(page, "db", None) and getattr(page, "project_id", None):
                            from .finance_tab import DBDataProvider  # type: ignore
                            prov = DBDataProvider(page)
                            _, _, preview_tax = prov.load_finance()
                            # preview_tax может быть словарём vendor->pct
                            try:
                                for v, pct in preview_tax.items():
                                    vendor_tax[normalize_case(v)] = float(pct) / 100.0
                            except Exception:
                                pass
                except Exception:
                    vendor_tax = {}
                smeta_opts["vendor_tax"] = vendor_tax
            # Используем расширенный генератор сметы
            elements += _build_smeta_report_diff(
                page,
                smeta_opts,
                header_style,
                normal_style,
                diff_map,
                current_items,
                compare_enabled,
            )
        elif report_type == "Погрузочная ведомость":
            elements += _build_load_report(page, options.get("load", {}), header_style, normal_style)
        elif report_type == "Финансовый отчёт":
            # Обновляем опции финансового отчёта для режима отображения сумм с учётом налога
            fin_opts = dict(options.get("fin", {}))
            if fin_opts.get("with_tax", False):
                vendor_tax: Dict[str, float] = {}
                try:
                    ft = getattr(page, "tab_finance_widget", None)
                    if ft and hasattr(ft, "preview_tax_pct"):
                        for v in ft.preview_tax_pct:
                            try:
                                vendor_tax[normalize_case(v)] = float(ft.preview_tax_pct.get(v, 0.0)) / 100.0
                            except Exception:
                                continue
                    else:
                        # Попытка загрузить данные из провайдера
                        if getattr(page, "db", None) and getattr(page, "project_id", None):
                            from .finance_tab import DBDataProvider  # type: ignore
                            prov = DBDataProvider(page)
                            _, _, preview_tax = prov.load_finance()
                            try:
                                for v, pct in preview_tax.items():
                                    vendor_tax[normalize_case(v)] = float(pct) / 100.0
                            except Exception:
                                pass
                except Exception:
                    vendor_tax = {}
                fin_opts["vendor_tax"] = vendor_tax
            # Если отчёт не является внутренним (internal_only=False) и не формат «Ксюша», выводим шапку с итогами
            if not fin_opts.get("internal_only", False) and not fin_opts.get("for_ksyusha", False):
                try:
                    subtotal_sum, tax_sum, total_sum = _compute_fin_report_totals(page, fin_opts)
                    # Формируем строку шапки: Итого, Налог, Итого с налогом
                    header_line = (
                        f"Итого: {subtotal_sum:,.2f} ₽; "
                        f"Налог: {tax_sum:,.2f} ₽; "
                        f"Итого с налогом: {total_sum:,.2f} ₽"
                    ).replace(",", " ")
                    elements.append(Paragraph(header_line, header_style))
                    elements.append(Spacer(1, 4 * mm))
                except Exception:
                    logger.error("Ошибка формирования шапки финансового отчёта", exc_info=True)
            # Затем строим сам финансовый отчёт
            # Передаем снимок финансовых данных, если режим сравнения включен
            elements += _build_fin_report(
                page,
                fin_opts,
                header_style,
                normal_style,
                fin_snapshot_data if compare_enabled else None
            )
        elif report_type == "Тайминг":
            timing_opts = options.get("timing", {})
            elements += _build_timing_report(page, header_style, normal_style, timing_opts)
        else:
            elements.append(Paragraph("Неизвестный тип отчёта", normal_style))

        # Генерируем PDF
        doc.build(elements)
        return file_path
    except Exception:
        logger.error("Ошибка формирования PDF", exc_info=True)
        return None


def _build_smeta_report(page: Any, smeta_opts: Dict[str, Any], header_style: ParagraphStyle, normal_style: ParagraphStyle, changed_map: Optional[Dict[Tuple[str, str, str, str], bool]] = None) -> List[Any]:
    """
    Формирует элементы отчёта для сметы.

    Принимает набор опций, управляющих добавлением шапки, отображением налогов
    и скидок, сортировкой и прикреплением изображений. В режиме сравнения
    выделяет изменённые строки, передавая карту ``changed_map``, где ключ —
    кортеж (подрядчик, наименование), а значение — флаг изменения.

    :param page: текущая страница проекта
    :param smeta_opts: словарь опций сметы
    :param header_style: стиль заголовков
    :param normal_style: стиль обычного текста
    :param changed_map: карта изменённых элементов для подсветки
    :return: список элементов для вставки в PDF
    """
    elements: List[Any] = []
    try:
        # Подготовка для отображения групп в смете. Если включён флаг
        # ``show_groups`` в опциях, строки отчёта будут окрашены в
        # соответствующий цвет группы. Для сопоставления названий групп
        # с цветами используется палитра GROUP_COLORS_RL. Словарь
        # ``_group_color_map`` хранит уже назначенные цвета, чтобы
        # одинаковые группы получали одинаковый цвет. Переменная
        # ``_group_color_index`` используется для перебора цветов
        # палитры.  Функция ``_get_rl_group_color`` возвращает цвет
        # ReportLab для заданной группы или None, если группу не
        # следует окрашивать (например, пустое название или
        # специальное значение 'аренда оборудования').
        show_groups: bool = bool(smeta_opts.get("show_groups"))
        _group_color_map: Dict[str, Any] = {}
        _group_color_index = [0]  # счётчик индекса цвета

        def _get_rl_group_color(gname: str) -> Optional[Any]:
            """Возвращает объект цвета reportlab для указанной группы.

            Цвет выбирается циклически из палитры GROUP_COLORS_RL и
            сохраняется в локальном словаре. Пустые или служебные
            группы (аренда оборудования) возвращают None, чтобы не
            окрашивать строки. Используется normalize_case для
            регистронезависимого сравнения.

            :param gname: исходное имя группы
            :return: объект цвета или None
            """
            try:
                gkey = normalize_case(gname or "")
            except Exception:
                gkey = gname or ""
            try:
                gkey_lower = str(gkey).strip().lower()
            except Exception:
                gkey_lower = str(gkey).lower() if gkey else ""
            # Не окрашиваем пустые и предопределённые группы
            if not gkey_lower or gkey_lower == normalize_case("аренда оборудования").strip().lower():
                return None
            # Если цвет уже назначен, возвращаем его
            if gkey in _group_color_map:
                return _group_color_map[gkey]
            # Назначаем новый цвет из палитры
            try:
                idx = _group_color_index[0] % len(GROUP_COLORS_RL)
                col = GROUP_COLORS_RL[idx]
                _group_color_map[gkey] = col
                _group_color_index[0] += 1
                return col
            except Exception:
                return None
        # Имя для зоны без названия (по умолчанию "Без зоны"). Используется при выводе группировки.
        no_zone_label = smeta_opts.get("no_zone_label", "").strip() or "Без зоны"
        # 1. Заголовочная информация (шапка проекта)
        if smeta_opts.get("add_header", False):
            # Список строк шапки
            info_lines: List[str] = []
            try:
                # Собираем поля из вкладки «Информация», если они существуют
                def add_field(label: str, attr_name: str) -> None:
                    try:
                        w = getattr(page, attr_name, None)
                        if w is not None:
                            # Для QTextEdit используем toPlainText, для QLineEdit — text
                            text = ""
                            if hasattr(w, "toPlainText"):
                                text = w.toPlainText().strip()
                            elif hasattr(w, "text"):
                                text = w.text().strip()
                            if text:
                                info_lines.append(f"{label} {text}")
                    except Exception:
                        pass
                # Название проекта (если есть отдельный атрибут project_name)
                if hasattr(page, "project_name") and page.project_name:
                    info_lines.append(f"Проект: {page.project_name}")
                # Поля из формы вкладки «Информация»
                add_field("Дата:", "ed_date")
                add_field("Заказчик:", "ed_customer")
                add_field("Дата и время заезда на монтаж:", "ed_mount_datetime")
                add_field("Готовность площадки:", "ed_site_ready")
                add_field("Адрес:", "ed_address")
                add_field("Готовность площадки (повтор):", "ed_site_ready_dup")
                add_field("Время демонтажа:", "ed_dismount_time")
                add_field("Этаж и наличие лифта:", "ed_floor_elevator")
                add_field("Количество электричества на площадке:", "ed_power_capacity")
                add_field("Возможность складирования кофров:", "ed_storage_possible")
                # Комментарии могут быть многострочными
                try:
                    comments_widget = getattr(page, "ed_comments", None)
                    if comments_widget is not None:
                        if hasattr(comments_widget, "toPlainText"):
                            comm = comments_widget.toPlainText().strip()
                            if comm:
                                # Делим комментарии на строки и добавляем каждую строку отдельно
                                for line in comm.splitlines():
                                    info_lines.append(f"Комментарий: {line}")
                except Exception:
                    pass
                # Теперь вычисляем суммы проекта и отображаем скидку/налог, если нужно
                show_tax = smeta_opts.get("show_taxes", False)
                show_disc = smeta_opts.get("show_discounts", False)
                if show_tax or show_disc:
                    # Готовим данные из вкладки Бухгалтерия или провайдера
                    total_discount = 0.0
                    total_tax_amt = 0.0
                    total_without_tax = 0.0
                    total_with_tax = 0.0
                    try:
                        # Загружаем элементы
                        items: List[Any] = []
                        vendor_settings: Dict[str, Any] = {}
                        ft = getattr(page, "tab_finance_widget", None)
                        if ft and hasattr(ft, "items"):
                            items = ft.items
                            # Используем настройки подрядчиков из вкладки
                            try:
                                vendor_settings = {v: s for v, s in ft.vendors_settings.items()}  # type: ignore
                            except Exception:
                                vendor_settings = {}
                        else:
                            # Пытаемся загрузить через провайдер
                            if getattr(page, "db", None) and getattr(page, "project_id", None):
                                from .finance_tab import DBDataProvider  # type: ignore
                                prov = DBDataProvider(page)
                                items = prov.load_items() or []
                                vendors, _, _ = prov.load_finance()
                                vendor_settings = vendors
                        # Собираем уникальные подрядчики
                        unique_vendors = sorted({it.vendor or "(без подрядчика)" for it in items})
                        # Заменяем отсутствующие настройки дефолтными
                        # импорт VendorSettings из finance_tab
                        try:
                            from .finance_tab import VendorSettings, aggregate_by_vendor, compute_client_flow  # type: ignore
                        except Exception:
                            VendorSettings = None  # type: ignore
                            aggregate_by_vendor = None  # type: ignore
                            compute_client_flow = None  # type: ignore
                        if VendorSettings is not None and aggregate_by_vendor is not None and compute_client_flow is not None:
                            # Заполняем vendor_settings для каждого подрядчика, если отсутствует
                            for v in unique_vendors:
                                if v not in vendor_settings:
                                    vendor_settings[v] = VendorSettings()
                            # Подготавливаем словарь coeffs, где ключи только для включённых коэффициентов
                            preview_coeffs: Dict[str, float] = {}
                            for v in unique_vendors:
                                s = vendor_settings.get(v)
                                try:
                                    if s and getattr(s, "coeff_enabled", False):
                                        preview_coeffs[v] = float(getattr(s, "coeff", 1.0))
                                except Exception:
                                    continue
                            # Агрегируем суммы (equip и other)
                            agg = aggregate_by_vendor(items, preview_coeffs)
                            # Подсчёт итогов
                            for v in unique_vendors:
                                data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
                                s = vendor_settings.get(v, VendorSettings())
                                equip_sum = data.get("equip_sum", 0.0)
                                other_sum = data.get("other_sum", 0.0)
                                disc_amt, comm_amt, tax_amt, subtotal, total_taxed = compute_client_flow(
                                    equip_sum, other_sum,
                                    getattr(s, "discount_pct", 0.0),
                                    getattr(s, "commission_pct", 0.0),
                                    getattr(s, "tax_pct", 0.0),
                                )
                                total_discount += disc_amt
                                total_tax_amt += tax_amt
                                total_without_tax += subtotal
                                total_with_tax += total_taxed
                    except Exception:
                        logger.error("Не удалось вычислить сумму/налог/скидку для шапки сметы", exc_info=True)
                    # Сумма до налога (и без налога) отражается как "Сумма проекта"
                    info_lines.append(f"Сумма проекта: {total_without_tax:,.2f} ₽".replace(",", " "))
                    # Если отображаем скидки
                    if show_disc:
                        info_lines.append(f"Скидка по проекту: {total_discount:,.2f} ₽".replace(",", " "))
                    # Если отображаем налоги
                    if show_tax:
                        # Избегаем деления на ноль
                        tax_rate = 0.0
                        try:
                            if total_without_tax > 0.0:
                                tax_rate = total_tax_amt / total_without_tax * 100.0
                        except Exception:
                            tax_rate = 0.0
                        info_lines.append(
                            f"Налог {round(tax_rate):.0f}%: {total_tax_amt:,.2f} ₽".replace(",", " ")
                        )
                        info_lines.append(
                            f"Сумма с налогом: {total_with_tax:,.2f} ₽".replace(",", " ")
                        )
                else:
                    # Если ни скидки, ни налоги не нужны, считаем сумму по позиции без учёта скидок/налога
                    total_sum = 0.0
                    try:
                        items_list: List[Any] = []
                        ft = getattr(page, "tab_finance_widget", None)
                        if ft and hasattr(ft, "items"):
                            items_list = ft.items
                        else:
                            # Загрузка через провайдер
                            if getattr(page, "db", None) and getattr(page, "project_id", None):
                                from .finance_tab import DBDataProvider  # type: ignore
                                prov = DBDataProvider(page)
                                items_list = prov.load_items() or []
                        for it in items_list:
                            total_sum += it.amount()
                        # 1.4.2.1 Подстановка суммы проекта при отсутствии данных
                        # Если список элементов пуст или сумма получилась нулевой, считаем
                        # сумму проекта через базу данных. Это предотвращает вывод нуля,
                        # когда вкладка Бухгалтерия ещё не открыта.
                        try:
                            if (not items_list) or (abs(total_sum) < 1e-6):
                                if getattr(page, "db", None) and getattr(page, "project_id", None):
                                    total_sum = page.db.project_total(page.project_id)
                        except Exception:
                            pass
                        info_lines.append(f"Сумма проекта: {total_sum:,.2f} ₽".replace(",", " "))
                    except Exception:
                        logger.error("Ошибка при подсчёте суммы проекта", exc_info=True)
            except Exception:
                logger.error("Ошибка при формировании шапки сметы", exc_info=True)
            # Добавляем строки шапки в документ. Если указана картинка для шапки,
            # располагаем её справа от текста в таблице из двух колонок.
            header_img_path = smeta_opts.get("header_image")
            # Импортируем классы под локальными именами, чтобы не затенять глобальные Table
            RLImage = None  # type: ignore
            RLTable = None  # type: ignore
            RLTableStyle = None  # type: ignore
            try:
                from reportlab.platypus import Image as RLImageCls, Table as RLTableCls, TableStyle as RLTableStyleCls
                RLImage, RLTable, RLTableStyle = RLImageCls, RLTableCls, RLTableStyleCls
            except Exception:
                RLImage = None
            # Если картинка существует и классы доступны
            if header_img_path and RLImage and os.path.exists(header_img_path):
                try:
                    # Формируем единый текст с <br/> для разделения строк, чтобы позволить перенос на новую страницу
                    text = "<br/>".join(info_lines)
                    text_para = Paragraph(text, normal_style)
                    # Создаём изображение
                    img = RLImage(header_img_path, width=60 * mm, height=30 * mm)
                    img.hAlign = "RIGHT"
                    # Таблица: левая колонка — текст, правая — картинка
                    table = RLTable(
                        [[text_para, img]],
                        colWidths=[110 * mm, 60 * mm],
                    )
                    table.setStyle(RLTableStyle([
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ]))
                    elements.append(table)
                    elements.append(Spacer(1, 4 * mm))
                except Exception:
                    # В случае ошибки возвращаемся к простому выводу строк и картинки
                    try:
                        for line in info_lines:
                            elements.append(Paragraph(line, normal_style))
                        elements.append(Spacer(1, 4 * mm))
                        img = RLImage(header_img_path, width=60 * mm, height=30 * mm)
                        img.hAlign = "RIGHT"
                        elements.append(img)
                        elements.append(Spacer(1, 3 * mm))
                    except Exception:
                        logger.error("Ошибка вставки изображения шапки", exc_info=True)
            else:
                # Картинка не задана или отсутствует: выводим строки по одной
                for line in info_lines:
                    elements.append(Paragraph(line, normal_style))
                elements.append(Spacer(1, 4 * mm))
        # 2. Картинка шапки: изображение уже вставляется вместе с текстовой шапкой (в таблице),
        # поэтому здесь ничего не делаем. Ранее использовался отдельный блок для картинки.
        # 3. Получаем список позиций
        items: List[Any] = []
        try:
            ft = getattr(page, "tab_finance_widget", None)
            if ft and hasattr(ft, "items"):
                items = ft.items
            else:
                if getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items = prov.load_items() or []
        except Exception:
            logger.error("Не удалось получить список позиций для сметы", exc_info=True)
            return elements
        # 3.1. Фильтрация по выбранному подрядчику (если активирована опция vendor_only)
        try:
            vendor_only_flag = bool(smeta_opts.get("vendor_only")) if smeta_opts else False
            selected_vendor = smeta_opts.get("vendor", "") if smeta_opts else ""
            # Если фильтр включён и задано имя подрядчика, оставляем только его позиции
            if vendor_only_flag and selected_vendor:
                # Нормализуем для точного сравнения (без учёта регистра)
                items = [it for it in items if normalize_case(getattr(it, "vendor", "")) == normalize_case(selected_vendor)]
        except Exception:
            # В случае ошибки фильтрации выводим предупреждение и продолжаем без фильтра
            logger.error("Ошибка фильтрации по подрядчику в сметном отчёте", exc_info=True)
        # Определяем базовые заголовки таблицы
        headers = ["Подрядчик", "Наименование", "Кол-во", "Коэф.", "Цена", "Сумма"]
        # Функция для формирования строки таблицы из элемента
        def make_row(it: Any) -> List[str]:
            qty_str = f"{getattr(it, 'qty', 0.0):.3f}".rstrip("0").rstrip(".")
            coeff_str = f"{getattr(it, 'coeff', 0.0):.3f}".rstrip("0").rstrip(".")
            price_str = f"{getattr(it, 'price', 0.0):.2f}".rstrip("0").rstrip(".")
            sum_str = f"{it.amount():.2f}".replace(",", ",").rstrip("0").rstrip(",")
            vendor = textwrap.fill(str(getattr(it, 'vendor', '')), width=50)
            name = textwrap.fill(str(getattr(it, 'name', '')), width=70)
            return [vendor, name, qty_str, coeff_str, price_str, sum_str]
        # Группировка и сортировка
        sort_by_zone = smeta_opts.get("sort_by_zone", False)
        sort_by_dept = smeta_opts.get("sort_by_department", False)
        # Стиль заголовков для зон/отделов — используем уменьшенную копию header_style
        zone_style = header_style.clone("ZoneHeader") if hasattr(header_style, "clone") else header_style
        dept_style = header_style.clone("DeptHeader") if hasattr(header_style, "clone") else header_style
        try:
            if hasattr(zone_style, "fontSize"):
                zone_style.fontSize = max(zone_style.fontSize - 2, 10)
                zone_style.leading = max(zone_style.leading - 2, zone_style.fontSize + 2)
            if hasattr(dept_style, "fontSize"):
                dept_style.fontSize = max(dept_style.fontSize - 3, 9)
                dept_style.leading = max(dept_style.leading - 3, dept_style.fontSize + 2)
        except Exception:
            pass
        # Сгруппированные вывод
        if sort_by_zone:
            # Группируем по зоне, затем по отделу
            # Сортировка зон
            zones = sorted(set(normalize_case(getattr(it, "zone", "") or "") for it in items))
            for zone in zones:
                # Фильтр по зоне
                zone_items = [it for it in items if normalize_case(getattr(it, "zone", "") or "") == zone]
                if not zone_items:
                    continue
                # Заголовок зоны
                # Подставляем альтернативное имя для пустой зоны
                zone_disp = zone or no_zone_label
                elements.append(Paragraph(f"Зона: {zone_disp}", zone_style))
                elements.append(Spacer(1, 2 * mm))
                # Группируем по отделу внутри зоны
                departments = sorted(set(normalize_case(getattr(it, "department", "") or "") for it in zone_items))
                for dept in departments:
                    dept_items = [it for it in zone_items if normalize_case(getattr(it, "department", "") or "") == dept]
                    if not dept_items:
                        continue
                    # Заголовок отдела (показываем даже если пустая строка для наглядности)
                    elements.append(Paragraph(f"Отдел: {dept or 'Без отдела'}", dept_style))
                    # Формируем таблицу
                    data: List[List[str]] = [headers]
                    for it in dept_items:
                        data.append(make_row(it))
                    col_widths_mm = [30, 80, 10, 10, 15, 15]
                    table = Table(
                        data,
                        repeatRows=1,
                        colWidths=[w * mm for w in col_widths_mm],
                    )
                    # Стилизация таблицы
                    table_style = TableStyle([
                        ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 10),
                        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                        ("FONTSIZE", (0, 1), (-1, -1), 9),
                    ])
                    # Если включено отображение групп, окрашиваем строки
                    if show_groups:
                        for row_idx, item in enumerate(dept_items, start=1):
                            try:
                                # Определяем имя группы для позиции
                                gname = getattr(item, "group_name", "") or ""
                                if not gname or not str(gname).strip():
                                    # Если group_name пустой, пытаемся извлечь префикс до двоеточия
                                    raw_name = getattr(item, "name", "") or ""
                                    # Приводим неразрывные и узкие пробелы к обычному для корректного деления
                                    try:
                                        raw = str(raw_name).replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
                                    except Exception:
                                        raw = str(raw_name)
                                    if ":" in raw:
                                        gname = raw.split(":", 1)[0].strip()
                                colr = _get_rl_group_color(str(gname))
                                if colr:
                                    table_style.add("BACKGROUND", (0, row_idx), (-1, row_idx), colr)
                            except Exception:
                                continue
                    # Подсветка изменённых строк
                    if changed_map:
                        for idx, it in enumerate(dept_items, start=1):
                            try:
                                v_key = normalize_case(getattr(it, "vendor", ""))
                                n_key = normalize_case(getattr(it, "name", ""))
                                d_key = normalize_case(getattr(it, "department", ""))
                                z_key = normalize_case(getattr(it, "zone", ""))
                                if (v_key, n_key, d_key, z_key) in changed_map:
                                    table_style.add("BACKGROUND", (0, idx), (-1, idx), colors.lavender)
                            except Exception:
                                continue
                    table.setStyle(table_style)
                    elements.append(table)
                    elements.append(Spacer(1, 3 * mm))
        elif sort_by_dept:
            # Группируем по отделу, затем по зоне
            departments = sorted(set(normalize_case(getattr(it, "department", "") or "") for it in items))
            for dept in departments:
                dept_items = [it for it in items if normalize_case(getattr(it, "department", "") or "") == dept]
                if not dept_items:
                    continue
                # Заголовок отдела
                elements.append(Paragraph(f"Отдел: {dept or 'Без отдела'}", zone_style))
                elements.append(Spacer(1, 2 * mm))
                # Далее разбиваем на зоны
                zones = sorted(set(normalize_case(getattr(it, "zone", "") or "") for it in dept_items))
                for zone in zones:
                    z_items = [it for it in dept_items if normalize_case(getattr(it, "zone", "") or "") == zone]
                    if not z_items:
                        continue
                    zone_disp = zone or no_zone_label
                    elements.append(Paragraph(f"Зона: {zone_disp}", dept_style))
                    data: List[List[str]] = [headers]
                    for it in z_items:
                        data.append(make_row(it))
                    col_widths_mm = [30, 80, 10, 10, 15, 15]
                    table = Table(
                        data,
                        repeatRows=1,
                        colWidths=[w * mm for w in col_widths_mm],
                    )
                    table_style = TableStyle([
                        ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 10),
                        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                        ("FONTSIZE", (0, 1), (-1, -1), 9),
                    ])
                    # Окраска строк по группам, если включена опция
                    if show_groups:
                        for row_idx, item in enumerate(z_items, start=1):
                            try:
                                gname = getattr(item, "group_name", "") or ""
                                if not gname or not str(gname).strip():
                                    raw_name = getattr(item, "name", "") or ""
                                    try:
                                        raw = str(raw_name).replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
                                    except Exception:
                                        raw = str(raw_name)
                                    if ":" in raw:
                                        gname = raw.split(":", 1)[0].strip()
                                colr = _get_rl_group_color(str(gname))
                                if colr:
                                    table_style.add("BACKGROUND", (0, row_idx), (-1, row_idx), colr)
                            except Exception:
                                continue
                    if changed_map:
                        for idx, it in enumerate(z_items, start=1):
                            try:
                                v_key = normalize_case(getattr(it, "vendor", ""))
                                n_key = normalize_case(getattr(it, "name", ""))
                                d_key = normalize_case(getattr(it, "department", ""))
                                z_key = normalize_case(getattr(it, "zone", ""))
                                if (v_key, n_key, d_key, z_key) in changed_map:
                                    table_style.add("BACKGROUND", (0, idx), (-1, idx), colors.lavender)
                            except Exception:
                                continue
                    table.setStyle(table_style)
                    elements.append(table)
                    elements.append(Spacer(1, 3 * mm))
            # Дополнительно выводим элементы классов Персонал/Логистика/Расходник как отдельные блоки
            special_classes = ["personnel", "logistic", "consumable"]
            for cls_en in special_classes:
                cls_items = [it for it in items if getattr(it, "cls", "") == cls_en]
                if not cls_items:
                    continue
                # Заголовок класса
                cls_ru = CLASS_EN2RU.get(cls_en, cls_en.capitalize())
                elements.append(Paragraph(f"{cls_ru}", zone_style))
                elements.append(Spacer(1, 2 * mm))
                # Группируем по зонам внутри класса
                class_zones = sorted(set(normalize_case(getattr(it, "zone", "") or "") for it in cls_items))
                for zone in class_zones:
                    cz_items = [it for it in cls_items if normalize_case(getattr(it, "zone", "") or "") == zone]
                    if not cz_items:
                        continue
                    zone_disp = zone or no_zone_label
                    elements.append(Paragraph(f"Зона: {zone_disp}", dept_style))
                    data_cls: List[List[str]] = [headers]
                    for it in cz_items:
                        data_cls.append(make_row(it))
                    col_widths_mm = [30, 80, 10, 10, 15, 15]
                    table_cls = Table(
                        data_cls,
                        repeatRows=1,
                        colWidths=[w * mm for w in col_widths_mm],
                    )
                    table_style_cls = TableStyle([
                        ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                        ("FONTSIZE", (0, 0), (-1, 0), 10),
                        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                        ("FONTSIZE", (0, 1), (-1, -1), 9),
                    ])
                    if changed_map:
                        for idx, it in enumerate(cz_items, start=1):
                            try:
                                v_key = normalize_case(getattr(it, "vendor", ""))
                                n_key = normalize_case(getattr(it, "name", ""))
                                d_key = normalize_case(getattr(it, "department", ""))
                                z_key = normalize_case(getattr(it, "zone", ""))
                                if (v_key, n_key, d_key, z_key) in changed_map:
                                    table_style_cls.add("BACKGROUND", (0, idx), (-1, idx), colors.lavender)
                            except Exception:
                                continue
                    table_cls.setStyle(table_style_cls)
                    elements.append(table_cls)
                    elements.append(Spacer(1, 3 * mm))
        else:
            # Без сортировки: одна таблица
            data: List[List[str]] = [headers]
            for it in items:
                data.append(make_row(it))
            col_widths_mm = [30, 80, 10, 10, 15, 15]
            table = Table(
                data,
                repeatRows=1,
                colWidths=[w * mm for w in col_widths_mm],
            )
            table_style = TableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                ("FONTSIZE", (0, 1), (-1, -1), 9),
            ])
            # Окраска строк по группам, если активирована опция. Работает только при
            # отсутствии сортировки, когда выводится одна общая таблица.
            if show_groups:
                for row_idx, item in enumerate(items, start=1):
                    try:
                        gname = getattr(item, "group_name", "") or ""
                        if not gname or not str(gname).strip():
                            raw_name = getattr(item, "name", "") or ""
                            try:
                                raw = str(raw_name).replace("\u00A0", " ").replace("\u202F", " ").replace("\u2007", " ")
                            except Exception:
                                raw = str(raw_name)
                            if ":" in raw:
                                gname = raw.split(":", 1)[0].strip()
                        colr = _get_rl_group_color(str(gname))
                        if colr:
                            table_style.add("BACKGROUND", (0, row_idx), (-1, row_idx), colr)
                    except Exception:
                        continue
            if changed_map:
                for idx, it in enumerate(items, start=1):
                    try:
                        v_key = normalize_case(getattr(it, "vendor", ""))
                        n_key = normalize_case(getattr(it, "name", ""))
                        d_key = normalize_case(getattr(it, "department", ""))
                        z_key = normalize_case(getattr(it, "zone", ""))
                        if (v_key, n_key, d_key, z_key) in changed_map:
                            table_style.add("BACKGROUND", (0, idx), (-1, idx), colors.lavender)
                    except Exception:
                        continue
            table.setStyle(table_style)
            elements.append(table)
        return elements
    except Exception:
        logger.error("Ошибка создания сметного отчёта", exc_info=True)
        return elements


def _build_load_report(page: Any, load_opts: Dict[str, Any], header_style: ParagraphStyle, normal_style: ParagraphStyle) -> List[Any]:
    """Формирует отчёт для погрузочной ведомости."""
    elements: List[Any] = []
    try:
        # Шапка при необходимости
        if load_opts.get("add_header", False):
            elements.append(Paragraph("Погрузочная ведомость", header_style))
            elements.append(Spacer(1, 4 * mm))
        # Получаем список позиций (equipment и consumable)
        items: List[Any] = []
        try:
            ft = getattr(page, "tab_finance_widget", None)
            if ft and hasattr(ft, "items"):
                items = ft.items
            else:
                if getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items = prov.load_items() or []
        except Exception:
            logger.error("Не удалось получить список позиций для погрузочной ведомости", exc_info=True)
        # Суммируем qty по имени и классу
        key_map: Dict[Tuple[str, str], float] = {}
        for it in items:
            try:
                if getattr(it, "cls", None) not in ("equipment", "consumable"):
                    continue
                cls_ru = CLASS_EN2RU.get(getattr(it, "cls", ""), getattr(it, "cls", ""))
                name = getattr(it, "name", "") or ""
                key = (name, cls_ru)
                key_map[key] = key_map.get(key, 0.0) + float(getattr(it, "qty", 0.0))
            except Exception:
                continue
        # Формируем данные для таблицы
        data: List[List[Any]] = []
        data.append(["Наименование", "Класс", "Кол-во"])
        for (name, cls_ru), qty_total in key_map.items():
            qty_str = f"{qty_total:.2f}".rstrip("0").rstrip(".")
            data.append([name, cls_ru, qty_str])
        # Создаём таблицу и применяем стили
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ]))
        elements.append(table)
    except Exception:
        logger.error("Ошибка создания погрузочного отчёта", exc_info=True)
    return elements


# ---------------------------------------------------------
# 5. Расширенный отчёт по смете с поддержкой сравнения
#
# Функция ``_build_smeta_report_diff`` формирует отчёт по смете,
# поддерживающий подсветку изменений, вывод блоков по зонам и
# отделам, отдельные блоки для классов «Персонал», «Логистика» и
# «Расходники», а также итоговые суммы по каждому блоку. Отчёт
# использует карту ``diff_map``, построенную в ``generate_pdf``, для
# вычисления дельт по количеству, цене и сумме, а также для
# определения состояния каждой позиции (добавлена, изменена, удалена).

def _build_smeta_report_diff(
    page: Any,
    smeta_opts: Dict[str, Any],
    header_style: ParagraphStyle,
    normal_style: ParagraphStyle,
    diff_map: Optional[Dict[Tuple[str, str, str, str], Dict[str, Any]]],
    current_items: Optional[List[Any]],
    compare_mode: bool,
) -> List[Any]:
    """
    Формирует отчёт для сметы с поддержкой сравнения.

    При наличии ``diff_map`` каждая позиция включает информацию о
    количественных и ценовых изменениях. Строки подсвечиваются в
    соответствии с логикой вкладки «Сводная смета». В отчёт входят
    блоки по зонам или отделам (в зависимости от опций), специальные
    классы выводятся отдельно. В конце каждого блока выводится его
    суммарная стоимость.

    :param page: объект ProjectPage
    :param smeta_opts: опции формирования сметы
    :param header_style: стиль заголовков
    :param normal_style: стиль обычного текста
    :param diff_map: карта различий, построенная в ``generate_pdf``
    :param current_items: список текущих позиций (если None, загружается
                          самостоятельно)
    :param compare_mode: признак, что режим сравнения включен. Если False,
                         таблицы строятся без колонок дельт.
    :return: список элементов для PDF
    """
    elements: List[Any] = []
    try:
        # 1. Шапка (всегда отображается)
        info_lines: List[str] = []
        # Имя для зоны без названия (по умолчанию "Без зоны"). Пользователь может задать своё
        # отображаемое название через параметры экспорта. Это имя используется только
        # для вывода заголовков «Зона: ...» и не влияет на внутренние ключи группировки
        # или сравнения. Если пользователь не ввёл собственного названия, используется
        # значение «Без зоны».
        no_zone_label = smeta_opts.get("no_zone_label", "").strip() or "Без зоны"
        try:
            # Вспомогательная функция для добавления полей проекта в шапку
            def add_field(label: str, attr_name: str) -> None:
                try:
                    w = getattr(page, attr_name, None)
                    if w is not None:
                        text = ""
                        if hasattr(w, "toPlainText"):
                            text = w.toPlainText().strip()
                        elif hasattr(w, "text"):
                            text = w.text().strip()
                        if text:
                            info_lines.append(f"{label} {text}")
                except Exception:
                    pass
            # Основные сведения о проекте
            if hasattr(page, "project_name") and page.project_name:
                info_lines.append(f"Проект: {page.project_name}")
            add_field("Дата:", "ed_date")
            add_field("Заказчик:", "ed_customer")
            add_field("Дата и время заезда на монтаж:", "ed_mount_datetime")
            add_field("Готовность площадки:", "ed_site_ready")
            add_field("Адрес:", "ed_address")
            add_field("Готовность площадки (повтор):", "ed_site_ready_dup")
            add_field("Время демонтажа:", "ed_dismount_time")
            add_field("Этаж и наличие лифта:", "ed_floor_elevator")
            add_field("Количество электричества на площадке:", "ed_power_capacity")
            add_field("Возможность складирования кофров:", "ed_storage_possible")
            # Комментарии
            try:
                comments_widget = getattr(page, "ed_comments", None)
                if comments_widget is not None and hasattr(comments_widget, "toPlainText"):
                    comm = comments_widget.toPlainText().strip()
                    if comm:
                        for line in comm.splitlines():
                            info_lines.append(f"Комментарий: {line}")
            except Exception:
                pass
            # Подготовка сумм для отображения скидок, налогов и итогов
            show_tax = bool(smeta_opts.get("show_taxes", False))
            show_disc = bool(smeta_opts.get("show_discounts", False))
            total_discount = 0.0
            total_tax_amt = 0.0
            total_without_tax = 0.0
            total_with_tax = 0.0
            if show_tax or show_disc:
                try:
                    items_for_sum: List[Any] = []
                    vendor_settings: Dict[str, Any] = {}
                    ft = getattr(page, "tab_finance_widget", None)
                    if ft and hasattr(ft, "items"):
                        items_for_sum = ft.items
                        try:
                            vendor_settings = {v: s for v, s in ft.vendors_settings.items()}  # type: ignore
                        except Exception:
                            vendor_settings = {}
                    else:
                        # Получаем данные из провайдера, если ``tab_finance_widget`` недоступен
                        if getattr(page, "db", None) and getattr(page, "project_id", None):
                            from .finance_tab import DBDataProvider  # type: ignore
                            prov = DBDataProvider(page)
                            items_for_sum = prov.load_items() or []
                            vendors, _, _ = prov.load_finance()
                            vendor_settings = vendors
                    # Собираем уникальные подрядчики
                    unique_vendors = sorted({it.vendor or "(без подрядчика)" for it in items_for_sum})
                    try:
                        from .finance_tab import VendorSettings, aggregate_by_vendor, compute_client_flow  # type: ignore
                    except Exception:
                        VendorSettings = None  # type: ignore
                        aggregate_by_vendor = None  # type: ignore
                        compute_client_flow = None  # type: ignore
                    if VendorSettings and aggregate_by_vendor and compute_client_flow:
                        # Убедимся, что у каждого подрядчика есть настройки
                        for v in unique_vendors:
                            if v not in vendor_settings:
                                vendor_settings[v] = VendorSettings()
                        # Коэффициенты для просмотра (агрегирования) с учётом активности коэффициента
                        preview_coeffs: Dict[str, float] = {}
                        for v in unique_vendors:
                            s = vendor_settings.get(v)
                            try:
                                if s and getattr(s, "coeff_enabled", False):
                                    preview_coeffs[v] = float(getattr(s, "coeff", 1.0))
                            except Exception:
                                continue
                        agg = aggregate_by_vendor(items_for_sum, preview_coeffs)
                        for v in unique_vendors:
                            data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
                            s = vendor_settings.get(v, VendorSettings())
                            equip_sum = data.get("equip_sum", 0.0)
                            other_sum = data.get("other_sum", 0.0)
                            disc_amt, comm_amt, tax_amt, subtotal, total_taxed = compute_client_flow(
                                equip_sum,
                                other_sum,
                                getattr(s, "discount_pct", 0.0),
                                getattr(s, "commission_pct", 0.0),
                                getattr(s, "tax_pct", 0.0),
                            )
                            total_discount += disc_amt
                            total_tax_amt += tax_amt
                            total_without_tax += subtotal
                            total_with_tax += total_taxed
                except Exception:
                    logger.error("Не удалось вычислить сумму/налог/скидку для шапки сметы", exc_info=True)
            # Формируем строки для скидок и налогов (если выбраны опции)
            if show_disc:
                try:
                    disc_str = fmt_num(total_discount, 2)
                    info_lines.append(f"Скидка: {disc_str}")
                except Exception:
                    pass
            if show_tax:
                try:
                    tax_str = fmt_num(total_tax_amt, 2)
                    info_lines.append(f"Налог: {tax_str}")
                except Exception:
                    pass
            # Пометка, если все цены пересчитываются с учётом налога
            try:
                if smeta_opts.get("with_tax"):
                    info_lines.append("ВСЕ ЦЕНЫ С НАЛОГОМ")
            except Exception:
                pass
            # Вывод итоговой суммы по проекту с учётом скидок/налогов
            try:
                if show_disc and show_tax:
                    # Сумма со скидкой (без налога) и сумма со скидкой и налогом
                    info_lines.append(f"Сумма проекта со скидкой: {fmt_num(total_without_tax, 2)}")
                    info_lines.append(f"Сумма проекта со скидкой и налогами: {fmt_num(total_with_tax, 2)}")
                elif show_disc:
                    # Только скидка
                    info_lines.append(f"Сумма проекта со скидкой: {fmt_num(total_without_tax, 2)}")
                elif show_tax:
                    # Только налог
                    info_lines.append(f"Сумма проекта: {fmt_num(total_without_tax, 2)}")
                    info_lines.append(f"Сумма проекта с налогом: {fmt_num(total_with_tax, 2)}")
                else:
                    # Без скидок/налогов
                    info_lines.append(f"Сумма проекта: {fmt_num(total_without_tax, 2)}")
            except Exception:
                logger.error("Не удалось сформировать строки сумм для шапки сметы", exc_info=True)
        except Exception:
            logger.error("Не удалось подготовить шапку сметы", exc_info=True)
        # Добавляем строки шапки в документ
        if info_lines:
            for line in info_lines:
                elements.append(Paragraph(line, normal_style))
            elements.append(Spacer(1, 4 * mm))
        # 2. Получаем список позиций
        try:
            # Загружаем текущие позиции из переданного списка или из виджета/базы
            if current_items is not None:
                items = list(current_items)
            else:
                ft = getattr(page, "tab_finance_widget", None)
                if ft and hasattr(ft, "items"):
                    items = ft.items
                elif getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items = prov.load_items() or []
                else:
                    items = []
            # 2.1 Фильтр по выбранному подрядчику (опция vendor_only)
            vendor_only_flag = bool(smeta_opts.get("vendor_only")) if smeta_opts else False
            selected_vendor = smeta_opts.get("vendor", "") if smeta_opts else ""
            if vendor_only_flag and selected_vendor:
                # Фильтруем текущие позиции по подрядчику
                items = [it for it in items if normalize_case(getattr(it, "vendor", "")) == normalize_case(selected_vendor)]
                # Фильтруем diff_map для сравнения: оставляем только записи выбранного подрядчика
                try:
                    if diff_map:
                        # Создаём новый словарь с ключами, где первый элемент (vendor) совпадает с выбранным подрядчиком
                        filtered_diff = {k: d for k, d in diff_map.items() if (k[0] == normalize_case(selected_vendor))}
                        diff_map = filtered_diff  # type: ignore
                except Exception:
                    logger.error("Ошибка фильтрации diff_map по подрядчику", exc_info=True)
        except Exception:
            logger.error("Не удалось получить список позиций для сметы", exc_info=True)
            items = []
        # 3. Определяем режим сортировки
        sort_by_zone = smeta_opts.get("sort_by_zone", False)
        sort_by_dept = smeta_opts.get("sort_by_department", False)
        # 4. Подготовка заголовков для зон/отделов
        zone_style = header_style.clone("ZoneHeader") if hasattr(header_style, "clone") else header_style
        dept_style = header_style.clone("DeptHeader") if hasattr(header_style, "clone") else header_style
        try:
            if hasattr(zone_style, "fontSize"):
                zone_style.fontSize = max(zone_style.fontSize - 2, 10)
                zone_style.leading = max(zone_style.leading - 2, zone_style.fontSize + 2)
            if hasattr(dept_style, "fontSize"):
                dept_style.fontSize = max(dept_style.fontSize - 3, 9)
                dept_style.leading = max(dept_style.leading - 3, dept_style.fontSize + 2)
        except Exception:
            pass
        # 5. Вспомогательная функция: создание записей
        def build_records(filter_fn) -> List[Dict[str, Any]]:
            recs: List[Dict[str, Any]] = []
            # Извлекаем карту налогов один раз для функции
            tax_map = {}
            with_tax_flag = False
            try:
                if smeta_opts:
                    tax_map = smeta_opts.get("vendor_tax", {}) or {}
                    with_tax_flag = bool(smeta_opts.get("with_tax"))
            except Exception:
                tax_map = {}
                with_tax_flag = False
            # текущие
            for it in items:
                try:
                    v_key = normalize_case(getattr(it, "vendor", ""))
                    n_key = normalize_case(getattr(it, "name", ""))
                    d_key = normalize_case(getattr(it, "department", ""))
                    z_key = normalize_case(getattr(it, "zone", ""))
                    cls_en = getattr(it, "cls", "equipment")
                    # Фильтрация по критерию
                    if not filter_fn(cls_en, v_key, n_key, d_key, z_key):
                        continue
                    qty = float(getattr(it, "qty", 0.0))
                    coeff = float(getattr(it, "coeff", 0.0))
                    price = float(getattr(it, "price", 0.0))
                    amount = qty * coeff * price
                    # Применяем налог, если включён
                    if with_tax_flag:
                        t = tax_map.get(v_key, 0.0)
                        price = price * (1.0 + t)
                        amount = amount * (1.0 + t)
                    # Сравниваем со снимком
                    key = (v_key, n_key, d_key, z_key)
                    ddata = diff_map.get(key) if diff_map else None
                    if ddata:
                        state = ddata.get("state", "")
                        diff_qty = ddata.get("diff_qty", 0.0)
                        diff_coeff = ddata.get("diff_coeff", 0.0)
                        diff_price = ddata.get("diff_price", 0.0)
                        diff_amount = ddata.get("diff_amount", 0.0)
                        # также масштабируем дельты, если включён налог
                        if with_tax_flag:
                            t = tax_map.get(v_key, 0.0)
                            diff_price = diff_price * (1.0 + t)
                            diff_amount = diff_amount * (1.0 + t)
                    else:
                        state = ""
                        diff_qty = diff_coeff = diff_price = diff_amount = 0.0
                    recs.append({
                        "vendor": v_key,
                        "name": n_key,
                        "department": d_key,
                        "zone": z_key,
                        "cls": cls_en,
                        "qty": qty,
                        "coeff": coeff,
                        "price": price,
                        "amount": amount,
                        "diff_qty": diff_qty,
                        "diff_coeff": diff_coeff,
                        "diff_price": diff_price,
                        "diff_amount": diff_amount,
                        "state": state,
                    })
                except Exception:
                    continue
            # удалённые
            if diff_map:
                for key, ddata in diff_map.items():
                    v_key, n_key, d_key, z_key = key
                    cls_en = ddata.get("class", "equipment")
                    if ddata.get("state") != "удалено":
                        continue
                    if not filter_fn(cls_en, v_key, n_key, d_key, z_key):
                        continue
                    recs.append({
                        "vendor": v_key,
                        "name": n_key,
                        "department": d_key,
                        "zone": z_key,
                        "cls": cls_en,
                        "qty": 0.0,
                        "coeff": ddata.get("snap_coeff", 0.0),
                        "price": 0.0,
                        "amount": 0.0,
                        "diff_qty": ddata.get("diff_qty", 0.0),
                        "diff_coeff": ddata.get("diff_coeff", 0.0),
                        # Масштабируем дельты для удалённых строк
                        "diff_price": (
                            ddata.get("diff_price", 0.0)
                            * (1.0 + (tax_map.get(v_key, 0.0) if with_tax_flag else 0.0))
                        ),
                        "diff_amount": (
                            ddata.get("diff_amount", 0.0)
                            * (1.0 + (tax_map.get(v_key, 0.0) if with_tax_flag else 0.0))
                        ),
                        "state": "удалено",
                    })
            recs.sort(key=lambda r: (r["vendor"], r["name"]))
            return recs
        # 6. Таблица строителя
        def build_table(records: List[Dict[str, Any]]) -> Table:
            """
            Формирует таблицу для текущего набора записей.

            Если ``compare_mode`` активен, столбцы с дельтами объединяются с
            основными значениями и отображаются в скобках. В противном случае
            выводится упрощённая таблица без столбца состояния и дельт.
            """
            if compare_mode:
                # Заголовки при сравнении
                headers = [
                    "Подрядчик",
                    "Наименование",
                    "Состояние",
                    "Кол-во",
                    "Коэф.",
                    "Цена/шт",
                    "Сумма",
                ]
            else:
                # Заголовки без сравнения
                headers = [
                    "Подрядчик",
                    "Наименование",
                    "Кол-во",
                    "Коэф.",
                    "Цена/шт",
                    "Сумма",
                ]
            data: List[List[str]] = [headers]
            for rec in records:
                # Подготовка отображаемых значений с учётом режима сравнения
                qty_str = fmt_num(rec["qty"], 3)
                coeff_str = fmt_num(rec["coeff"], 3)
                price_str = fmt_num(rec["price"], 2)
                amount_str = fmt_num(rec["amount"], 2)
                if compare_mode:
                    dq = rec.get("diff_qty", 0.0)
                    dcoeff = rec.get("diff_coeff", 0.0)
                    dp = rec.get("diff_price", 0.0)
                    da = rec.get("diff_amount", 0.0)
                    # Формируем текст с дельтой в скобках
                    qty_str = qty_str + (f" ({fmt_sign(dq, 3)})" if abs(dq) >= 1e-6 else "")
                    coeff_str = coeff_str + (f" ({fmt_sign(dcoeff, 3)})" if abs(dcoeff) >= 1e-6 else "")
                    price_str = price_str + (f" ({fmt_sign(dp, 2)})" if abs(dp) >= 1e-6 else "")
                    amount_str = amount_str + (f" ({fmt_sign(da, 2)})" if abs(da) >= 1e-6 else "")
                    # Строка состояния: скрываем "не изменилось"
                    st = rec.get("state", "")
                    state_display = "" if st == "не изменилось" else st
                    row = [
                        textwrap.fill(rec["vendor"], width=25),
                        textwrap.fill(rec["name"], width=25),
                        state_display,
                        qty_str,
                        coeff_str,
                        price_str,
                        amount_str,
                    ]
                else:
                    row = [
                        textwrap.fill(rec["vendor"], width=40),
                        textwrap.fill(rec["name"], width=40),
                        qty_str,
                        coeff_str,
                        price_str,
                        amount_str,
                    ]
                data.append(row)
            # Определяем ширины столбцов для альбомной ориентации (ширина листа 297 мм,
            # поля по 15 мм с каждой стороны, оставшаяся ширина ~267 мм).
            avail_width_mm = 267.0
            if compare_mode:
                # 7 колонок: подрядчик, наименование, состояние, кол-во, коэф., цена, сумма.
                # Задаём фиксированные ширины для первых шести колонок, последнюю
                # определяем как остаток, чтобы заполнить всю доступную ширину.
                fixed = [35.0, 90.0, 30.0, 25.0, 25.0, 25.0]
                last_width = max(20.0, avail_width_mm - sum(fixed))
                col_widths_mm = fixed + [last_width]
            else:
                # 6 колонок: подрядчик, наименование, кол-во, коэф., цена, сумма.
                fixed = [40.0, 110.0, 25.0, 25.0, 25.0]
                last_width = max(20.0, avail_width_mm - sum(fixed))
                col_widths_mm = fixed + [last_width]
            table = Table(
                data,
                repeatRows=1,
                colWidths=[w * mm for w in col_widths_mm],
            )
            # Основные стили таблицы
            table_style = TableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
            ])
            # Подсветка строк и текстов при сравнении
            if compare_mode:
                for idx, rec in enumerate(records, start=1):
                    state = rec.get("state", "")
                    dq = rec.get("diff_qty", 0.0)
                    dcoeff = rec.get("diff_coeff", 0.0)
                    dp = rec.get("diff_price", 0.0)
                    da = rec.get("diff_amount", 0.0)
                    # Любые изменения – жёлтый фон
                    if state and state != "не изменилось":
                        table_style.add("BACKGROUND", (0, idx), (-1, idx), colors.Color(1.0, 0.97, 0.90))
                    # Цвет текста для ячеек с дельтами: красный при увеличении, зелёный при уменьшении
                    # Количество
                    if abs(dq) >= 1e-6:
                        color = colors.Color(0.65, 0.0, 0.0) if dq > 0 else colors.Color(0.0, 0.45, 0.0)
                        # Кол-во колонка index зависит от compare_mode: 3
                        table_style.add("TEXTCOLOR", (3, idx), (3, idx), color)
                    # Коэффициент
                    if abs(dcoeff) >= 1e-6:
                        color = colors.Color(0.65, 0.0, 0.0) if dcoeff > 0 else colors.Color(0.0, 0.45, 0.0)
                        table_style.add("TEXTCOLOR", (4, idx), (4, idx), color)
                    # Цена
                    if abs(dp) >= 1e-6:
                        color = colors.Color(0.65, 0.0, 0.0) if dp > 0 else colors.Color(0.0, 0.45, 0.0)
                        table_style.add("TEXTCOLOR", (5, idx), (5, idx), color)
                    # Сумма
                    if abs(da) >= 1e-6:
                        color = colors.Color(0.65, 0.0, 0.0) if da > 0 else colors.Color(0.0, 0.45, 0.0)
                        table_style.add("TEXTCOLOR", (6, idx), (6, idx), color)
            else:
                # В режиме без сравнения выравниваем первые два столбца по левому краю
                table_style.add("ALIGN", (0, 1), (1, -1), "LEFT")
            table.setStyle(table_style)
            return table
        # 7. Вывод блоков
        if sort_by_zone:
            zones = sorted(set(normalize_case(getattr(it, "zone", "") or "") for it in items))
            for zone in zones:
                # Подставляем пользовательское название для пустой зоны
                zone_label = zone or no_zone_label
                elements.append(Paragraph(f"Зона: {zone_label}", zone_style))
                elements.append(Spacer(1, 2 * mm))
                # Общая сумма по зоне (включает все классы и отделы)
                try:
                    zone_total = sum(r["amount"] for r in build_records(lambda cls_en, v, n, d, z: (z == zone)))
                    elements.append(Paragraph(f"Итого по зоне: {fmt_num(zone_total, 2)}", normal_style))
                    elements.append(Spacer(1, 2 * mm))
                except Exception:
                    pass
                # equipment per departments inside zone
                departments = sorted(set(normalize_case(getattr(it, "department", "") or "") for it in items if getattr(it, "cls", "equipment") == "equipment" and normalize_case(getattr(it, "zone", "") or "") == zone))
                for dept in departments:
                    dept_label = dept or "Без отдела"
                    elements.append(Paragraph(f"Отдел: {dept_label}", dept_style))
                    recs = build_records(lambda cls_en, v, n, d, z: (z == zone and d == dept and cls_en == "equipment"))
                    if recs:
                        block_total = sum(r["amount"] for r in recs)
                        elements.append(Paragraph(f"Итого: {fmt_num(block_total, 2)}", normal_style))
                        # Небольшой отступ перед таблицей, чтобы сумма не "лежала" на таблице
                        elements.append(Spacer(1, 1 * mm))
                        elements.append(build_table(recs))
                        elements.append(Spacer(1, 4 * mm))
                # special classes inside zone
                for cls_en in ["personnel", "logistic", "consumable"]:
                    recs = build_records(lambda cls_e, v, n, d, z: (z == zone and cls_e == cls_en))
                    if recs:
                        cls_ru = CLASS_EN2RU.get(cls_en, cls_en.capitalize())
                        elements.append(Paragraph(f"{cls_ru}", dept_style))
                        block_total = sum(r["amount"] for r in recs)
                        elements.append(Paragraph(f"Итого: {fmt_num(block_total, 2)}", normal_style))
                        elements.append(Spacer(1, 1 * mm))
                        elements.append(build_table(recs))
                        elements.append(Spacer(1, 4 * mm))
        elif sort_by_dept:
            departments = sorted(set(normalize_case(getattr(it, "department", "") or "") for it in items))
            for dept in departments:
                dept_label = dept or "Без отдела"
                elements.append(Paragraph(f"Отдел: {dept_label}", zone_style))
                elements.append(Spacer(1, 2 * mm))
                # Общая сумма по отделу (включает все зоны и классы)
                try:
                    dept_total = sum(r["amount"] for r in build_records(lambda cls_en, v, n, d_, z: (d_ == dept)))
                    elements.append(Paragraph(f"Итого по отделу: {fmt_num(dept_total, 2)}", normal_style))
                    elements.append(Spacer(1, 2 * mm))
                except Exception:
                    pass
                # equipment per zone
                zones = sorted(set(normalize_case(getattr(it, "zone", "") or "") for it in items if normalize_case(getattr(it, "department", "") or "") == dept and getattr(it, "cls", "equipment") == "equipment"))
                for zone in zones:
                    # Подставляем пользовательское название для пустой зоны
                    zone_label = zone or no_zone_label
                    elements.append(Paragraph(f"Зона: {zone_label}", dept_style))
                    recs = build_records(lambda cls_en, v, n, d_, z: (d_ == dept and z == zone and cls_en == "equipment"))
                    if recs:
                        block_total = sum(r["amount"] for r in recs)
                        elements.append(Paragraph(f"Итого: {fmt_num(block_total, 2)}", normal_style))
                        elements.append(Spacer(1, 1 * mm))
                        elements.append(build_table(recs))
                        elements.append(Spacer(1, 4 * mm))
                # special classes per department
                for cls_en in ["personnel", "logistic", "consumable"]:
                    recs = build_records(lambda cls_e, v, n, d_, z: (d_ == dept and cls_e == cls_en))
                    if recs:
                        cls_ru = CLASS_EN2RU.get(cls_en, cls_en.capitalize())
                        elements.append(Paragraph(f"{cls_ru}", dept_style))
                        block_total = sum(r["amount"] for r in recs)
                        elements.append(Paragraph(f"Итого: {fmt_num(block_total, 2)}", normal_style))
                        elements.append(build_table(recs))
                        elements.append(Spacer(1, 4 * mm))
        else:
            # без сортировки
            recs = build_records(lambda cls_en, v, n, d, z: True)
            if recs:
                block_total = sum(r["amount"] for r in recs)
                elements.append(Paragraph(f"Итого: {fmt_num(block_total, 2)}", normal_style))
                elements.append(Spacer(1, 1 * mm))
                elements.append(build_table(recs))
        return elements
    except Exception:
        # Логируем любую ошибку, возникшую при построении отчёта
        logger.error("Ошибка создания расширенного сметного отчёта", exc_info=True)
        return elements


def _build_fin_report(page: Any, fin_opts: Dict[str, Any], header_style: ParagraphStyle, normal_style: ParagraphStyle, fin_snapshot: Optional[Dict[str, Any]] = None) -> List[Any]:
    """
    Формирует расширенный финансовый отчёт с учётом скидок, комиссий и налогов.

    При передаче параметра ``fin_snapshot`` (снимок финансовых данных) строки
    отчёта подсвечиваются в зависимости от изменения сумм: если сумма
    увеличилась по сравнению со снимком — строка окрашивается красным и
    рядом выводится положительная дельта; если уменьшилась — зелёным и
    выводится отрицательная дельта. Снимок должен иметь структуру,
    возвращаемую функцией ``compute_fin_snapshot_data``: словари сумм по
    подрядчикам, зонам, отделам, классам и общая сумма проекта.

    Для каждого подрядчика выводятся суммы оборудования и прочих затрат, размер
    скидки, комиссии, налога, итог до налога и итог с налогом. Если вкладка
    «Бухгалтерия» доступна, используются агрегированные данные и настройки
    конкретного проекта. В противном случае агрегаты рассчитываются на лету.

    :param page: объект страницы проекта
    :param fin_opts: словарь опций (``show_internal``, ``show_agents`` и др.)
    :param header_style: стиль заголовков
    :param normal_style: стиль обычного текста
    :param fin_snapshot: снимок финансового отчёта для подсветки изменений;
        если None, подсветка не применяется
    :return: список элементов отчёта
    """
    # Если выбран специальный формат отчёта для Ксюши, делегируем построение
    # упрощённого финансового отчёта. В этом режиме игнорируются другие
    # параметры (внутренние, зоны и т.п.) и создаётся таблица подрядчик/суммы.
    if fin_opts.get("for_ksyusha", False):
        return _build_fin_report_ksyusha(page, fin_opts, header_style, normal_style)
    elements: List[Any] = []
    try:
        # Получаем виджет бухгалтерии или провайдер для сбора данных
        ft = getattr(page, "tab_finance_widget", None)
        # Имя для зоны без названия (по умолчанию "Без зоны").
        # Пользователь может задать своё отображаемое название для пустой зоны
        # через опции экспорта. Это значение будет использоваться только для
        # отображения в итоговых таблицах и не влияет на внутреннюю логику
        # группирования или суммирования.
        no_zone_label = fin_opts.get("no_zone_label", "").strip() or "Без зоны"

        # Флаг: показывать только итоговую таблицу по зонам. Если True,
        # отчёт будет содержать только блок «Итоги по зонам». В зависимости
        # от флага ``zones_by_vendor`` либо отображается сводная таблица
        # сумм по зонам, либо таблица с распределением по подрядчикам внутри
        # каждой зоны. При False создаются стандартные таблицы по
        # подрядчикам, отделам и классам.
        zones_only = bool(fin_opts.get("zones_only")) if fin_opts else False
        # Новый флаг: распределять суммы по подрядчикам внутри каждой зоны.
        # Используется только в режиме ``zones_only``.
        zones_by_vendor = bool(fin_opts.get("zones_by_vendor")) if fin_opts else False
        # Распаковываем данные снимка, если он передан, чтобы позже сравнивать суммы
        snap_vendors: Dict[str, float] = {}
        snap_zones: Dict[str, float] = {}
        snap_depts: Dict[str, float] = {}
        snap_classes: Dict[str, float] = {}
        snap_project_total: Optional[float] = None
        if fin_snapshot:
            try:
                snap_vendors = dict(fin_snapshot.get("vendors", {}))  # type: ignore
                snap_zones = dict(fin_snapshot.get("zones", {}))  # type: ignore
                snap_depts = dict(fin_snapshot.get("departments", {}))  # type: ignore
                snap_classes = dict(fin_snapshot.get("classes", {}))  # type: ignore
                snap_project_total = fin_snapshot.get("project_total")  # type: ignore
            except Exception:
                snap_vendors = {}
                snap_zones = {}
                snap_depts = {}
                snap_classes = {}
                snap_project_total = None
        # Если выбрана опция «только внутренние расчёты», строим отдельный отчёт
        if fin_opts.get("internal_only", False):
            try:
                # Собираем данные из вкладки «Бухгалтерия» или через провайдер
                items: List[Any] = []
                agg: Dict[str, Dict[str, float]] = {}
                if ft and hasattr(ft, "items"):
                    items = list(ft.items)
                else:
                    # При отсутствии виджета таблицы загружаем позиции через провайдер
                    try:
                        if getattr(page, "db", None) and getattr(page, "project_id", None):
                            from .finance_tab import DBDataProvider  # type: ignore
                            prov = DBDataProvider(page)
                            items = prov.load_items() or []
                    except Exception:
                        items = []
                # Заполняем агрегат по подрядчикам (сумма equipment) для расчёта внутренней скидки
                for it in items:
                    try:
                        v = getattr(it, "vendor", "") or "(без подрядчика)"
                        data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
                        amt = it.amount()
                        if getattr(it, "cls", "") == "equipment":
                            data["equip_sum"] += amt
                        else:
                            data["other_sum"] += amt
                        agg[v] = data
                    except Exception:
                        continue
                # Для расчёта внутренней скидки используем compute_internal_discount
                try:
                    from .finance_tab import compute_internal_discount  # type: ignore
                except Exception:
                    compute_internal_discount = None  # type: ignore
                # Пытаемся импортировать классы и функции из finance_tab (для типов и вычислений)
                try:
                    from .finance_tab import VendorSettings, ProfitItem, ExpenseItem  # type: ignore
                except Exception:
                    VendorSettings = None  # type: ignore
                    ProfitItem = None  # type: ignore
                    ExpenseItem = None  # type: ignore

                # Список доходов (vendor, description, amount)
                income_rows: List[Tuple[str, str, float]] = []
                # 1. Наши внутренние скидки
                for vendor in sorted(agg.keys()):
                    equip_sum = agg[vendor].get("equip_sum", 0.0)
                    # Значения предпросмотра скидки/комиссии из FinanceTab
                    discount_pct = 0.0
                    our_pct: Optional[float] = None
                    our_sum: Optional[float] = None
                    client_disc_amount = 0.0
                    if ft:
                        try:
                            # Скидка клиента
                            discount_pct = ft.preview_discount_pct.get(vendor, ft.vendors_settings.get(vendor, VendorSettings()).discount_pct)  # type: ignore
                            # Наши скидки (% и сумма)
                            our_pct = ft.preview_our_discount_pct.get(vendor, ft.vendors_settings.get(vendor, VendorSettings()).our_discount_pct)  # type: ignore
                            our_sum = ft.preview_our_discount_sum.get(vendor, ft.vendors_settings.get(vendor, VendorSettings()).our_discount_sum)  # type: ignore
                        except Exception:
                            pass
                    try:
                        client_disc_amount = equip_sum * (discount_pct / 100.0)
                    except Exception:
                        client_disc_amount = 0.0
                    if compute_internal_discount:
                        try:
                            # Определяем комиссию для данного подрядчика. Если вкладка FinanceTab доступна,
                            # берём коэффициент комиссии из предпросмотра, иначе используем значение из настроек.
                            commission_pct = 0.0
                            if ft:
                                try:
                                    commission_pct = ft.preview_commission_pct.get(
                                        vendor,
                                        ft.vendors_settings.get(vendor, VendorSettings()).commission_pct,  # type: ignore
                                    )
                                except Exception:
                                    commission_pct = 0.0
                            # Рассчитываем абсолютную комиссию на сумму после клиентской скидки
                            commission_amount = (equip_sum - client_disc_amount) * (commission_pct / 100.0)
                            # Вычисляем внутреннюю скидку с учётом клиентской скидки и комиссии
                            internal_amt = compute_internal_discount(
                                equip_sum,
                                client_disc_amount,
                                commission_amount,
                                our_pct,
                                our_sum,
                            )  # type: ignore
                        except Exception:
                            internal_amt = 0.0
                    else:
                        internal_amt = 0.0
                    if internal_amt > 0.0:
                        income_rows.append((vendor, "Наша скидка", float(internal_amt)))
                # 2. Доходы из сметы (profit_items)
                profits: List[ProfitItem] = []
                if ft and hasattr(ft, "profit_items"):
                    profits = list(ft.profit_items)
                # Если есть profits, добавляем
                for p in profits:
                    vendor = p.vendor or "(без подрядчика)"
                    income_rows.append((vendor, p.description, float(p.amount)))
                # Считаем итог доходов
                total_income = sum(row[2] for row in income_rows)
                # Список расходов (name, qty, price, total)
                expense_rows: List[Tuple[str, float, float, float]] = []
                expenses: List[ExpenseItem] = []
                if ft and hasattr(ft, "expense_items"):
                    expenses = list(ft.expense_items)
                for e in expenses:
                    expense_rows.append((e.name or "", float(e.qty), float(e.price), float(e.total())))
                total_expense = sum(r[3] for r in expense_rows)
                net_profit = total_income - total_expense
                # Формируем элементы отчёта
                elements.append(Paragraph("Внутренние расчёты", header_style))
                elements.append(Spacer(1, 2 * mm))
                # Доходы
                elements.append(Paragraph("Доходы", header_style))
                if income_rows:
                    # Формируем таблицу доходов
                    inc_data: List[List[str]] = [["Подрядчик", "Описание", "Сумма ₽"]]
                    for v, dsc, amt in income_rows:
                        inc_data.append([
                            v,
                            dsc,
                            f"{amt:,.2f}".replace(",", " ")
                        ])
                    # Пытаемся масштабировать ширины столбцов под доступную ширину страницы
                    try:
                        # Вычисляем ширину страницы для ландшафтной ориентации.
                        # По умолчанию A4 представляет портретные размеры (595x842 pt). Используем
                        # большую сторону как ширину в landscape. Затем вычитаем поля (по 15 мм).
                        try:
                            page_width_pts = max(A4[0], A4[1])  # ширина листа в landscape
                        except Exception:
                            page_width_pts = A4[0]  # fallback
                        available_width_mm = (page_width_pts - 2 * 15 * mm) / mm  # type: ignore
                        base_widths = [60.0, 100.0, 40.0]
                        total_base = sum(base_widths) or 1.0
                        scale = available_width_mm / total_base
                        col_widths_mm = [w * scale for w in base_widths]
                        inc_table = Table(inc_data, repeatRows=1, colWidths=[w * mm for w in col_widths_mm])
                    except Exception:
                        inc_table = Table(inc_data, repeatRows=1, colWidths=[60 * mm, 100 * mm, 40 * mm])
                    # Применяем стиль для единообразия шрифта и выравнивания
                    from reportlab.platypus import TableStyle as LocalTableStyle  # type: ignore
                    try:
                        inc_table.setStyle(LocalTableStyle([
                            ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                            # Убираем внутренние отступы таблицы, чтобы она занимала всю доступную ширину
                            ('LEFTPADDING', (0, 0), (-1, -1), 0),
                            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                        ]))
                        elements.append(inc_table)
                    except Exception:
                        elements.append(Paragraph("(Не удалось построить таблицу доходов)", normal_style))
                else:
                    elements.append(Paragraph("Доходы отсутствуют", normal_style))
                elements.append(Spacer(1, 4 * mm))
                # Расходы
                elements.append(Paragraph("Расходы", header_style))
                if expense_rows:
                    exp_data: List[List[str]] = [["Наименование", "Кол-во", "Цена ₽", "Сумма ₽"]]
                    for name, qty, price, tot in expense_rows:
                        exp_data.append([
                            name,
                            f"{qty:,.2f}".replace(",", " "),
                            f"{price:,.2f}".replace(",", " "),
                            f"{tot:,.2f}".replace(",", " ")
                        ])
                    # Масштабируем ширины столбцов для таблицы расходов по аналогии с доходами
                    try:
                        try:
                            page_width_pts = max(A4[0], A4[1])  # ширина листа в landscape
                        except Exception:
                            page_width_pts = A4[0]
                        available_width_mm = (page_width_pts - 2 * 15 * mm) / mm  # type: ignore
                        base_widths_exp = [70.0, 30.0, 30.0, 40.0]
                        total_base_exp = sum(base_widths_exp) or 1.0
                        scale_exp = available_width_mm / total_base_exp
                        col_widths_exp_mm = [w * scale_exp for w in base_widths_exp]
                        exp_table = Table(exp_data, repeatRows=1, colWidths=[w * mm for w in col_widths_exp_mm])
                    except Exception:
                        exp_table = Table(exp_data, repeatRows=1, colWidths=[70 * mm, 30 * mm, 30 * mm, 40 * mm])
                    # Стилизация таблицы
                    from reportlab.platypus import TableStyle as LocalTableStyle  # type: ignore
                    try:
                        exp_table.setStyle(LocalTableStyle([
                            ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('ALIGN', (1, 1), (3, -1), 'RIGHT'),
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                            ('LEFTPADDING', (0, 0), (-1, -1), 0),
                            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                        ]))
                        elements.append(exp_table)
                    except Exception:
                        elements.append(Paragraph("(Не удалось построить таблицу расходов)", normal_style))
                else:
                    elements.append(Paragraph("Расходы отсутствуют", normal_style))
                elements.append(Spacer(1, 4 * mm))
                # Итоги
                elements.append(Paragraph(f"Итого доходы: {total_income:,.2f} ₽".replace(",", " "), header_style))
                elements.append(Paragraph(f"Итого расходы: {total_expense:,.2f} ₽".replace(",", " "), header_style))
                elements.append(Paragraph(f"Чистая прибыль: {net_profit:,.2f} ₽".replace(",", " "), header_style))
                elements.append(Spacer(1, 4 * mm))
                # 3.3.4 Таблица: сколько нужно собрать денег у каждого подрядчика
                # Суммируем все доходы по подрядчикам (из income_rows), получая общую
                # сумму к сбору для каждого подрядчика. Выводим её отдельной таблицей.
                try:
                    vendor_income_sum: Dict[str, float] = {}
                    for _v_name, _desc, amt_val in income_rows:
                        try:
                            key = _v_name or "(без подрядчика)"
                            vendor_income_sum[key] = vendor_income_sum.get(key, 0.0) + float(amt_val)
                        except Exception:
                            continue
                    if vendor_income_sum:
                        elements.append(Paragraph(
                            "Сколько нужно собрать денег у каждого подрядчика",
                            header_style,
                        ))
                        elements.append(Spacer(1, 1 * mm))
                        table_data: List[List[str]] = [["Подрядчик", "Сумма ₽"]]
                        for vn in sorted(vendor_income_sum.keys()):
                            amt_val = vendor_income_sum[vn]
                            table_data.append([
                                vn,
                                f"{amt_val:,.2f}".replace(",", " "),
                            ])
                        try:
                            from reportlab.platypus import Table as LocalTable, TableStyle as LocalTableStyle  # type: ignore
                            # Ширины столбцов: рассчитываем на основе доступной ширины страницы
                            try:
                                page_width_pts = max(A4[0], A4[1])
                            except Exception:
                                page_width_pts = A4[0]
                            available_width_mm = (page_width_pts - 2 * 15 * mm) / mm  # type: ignore
                            base_widths = [100.0, 40.0]
                            total_base = sum(base_widths) or 1.0
                            scale = available_width_mm / total_base
                            col_widths_mm = [w * scale for w in base_widths]
                            tbl_collect = LocalTable(
                                table_data,
                                repeatRows=1,
                                colWidths=[w * mm for w in col_widths_mm],
                            )
                        except Exception:
                            tbl_collect = LocalTable(table_data, repeatRows=1, colWidths=[100 * mm, 40 * mm])  # type: ignore
                        try:
                            tbl_collect.setStyle(LocalTableStyle([
                                ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
                                ('FONTSIZE', (0, 0), (-1, -1), 10),
                                ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                                ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
                            ]))
                            elements.append(tbl_collect)
                        except Exception:
                            elements.append(Paragraph(
                                "(Не удалось построить таблицу сбора средств)",
                                normal_style,
                            ))
                        elements.append(Spacer(1, 4 * mm))
                except Exception:
                    # Логируем ошибку построения таблицы и продолжаем
                    logger.error(
                        "Ошибка построения таблицы итогов по подрядчикам во внутреннем отчёте",
                        exc_info=True,
                    )
                return elements
            except Exception:
                # В случае ошибки при построении внутреннего отчёта логируем и возвращаем сообщение
                logger.error("Ошибка создания внутреннего финансового отчёта", exc_info=True)
                elements.append(Paragraph("Ошибка создания внутреннего финансового отчёта", normal_style))
                return elements

        items: List[Any] = []
        agg: Dict[str, Dict[str, float]] = {}
        try:
            if ft and hasattr(ft, "_agg_latest"):
                agg = ft._agg_latest  # type: ignore
                if hasattr(ft, "items"):
                    items = ft.items
            else:
                if getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items = prov.load_items() or []
        except Exception:
            items = []
        # Если агрегат пуст — заполняем на основе позиций
        if not agg:
            for it in items:
                try:
                    v = getattr(it, "vendor", "") or "(без подрядчика)"
                    data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0, "total_sum": 0.0})
                    amt = it.amount()
                    data["total_sum"] += amt
                    if getattr(it, "cls", "") == "equipment":
                        data["equip_sum"] += amt
                    else:
                        data["other_sum"] += amt
                    agg[v] = data
                except Exception:
                    continue
        if not agg:
            elements.append(Paragraph("Финансовые данные недоступны", normal_style))
            return elements
        # Если выбран режим «цены с налогом», добавляем поясняющую метку.
        # Отдельная строка «Итого» выводится в заголовке отчёта (в generate_pdf),
        # поэтому здесь её опускаем, чтобы избежать дублирования.
        if fin_opts.get("with_tax"):
            elements.append(Paragraph("ВСЕ ЦЕНЫ С НАЛОГОМ", header_style))
            elements.append(Spacer(1, 2 * mm))
        # Импортируем функцию расчёта из модуля finance_tab
        try:
            from .finance_tab import compute_client_flow  # type: ignore
        except Exception:
            compute_client_flow = None  # type: ignore
        # Готовим информацию по каждому подрядчику и определяем, нужно ли
        # выводить колонки скидки и комиссии
        vendors_info: List[Dict[str, Any]] = []
        include_discount = False
        include_commission = False
        # Суммарные итоговые значения
        grand_totals = {
            "equip_sum": 0.0,
            "other_sum": 0.0,
            "discount_sum": 0.0,
            "commission_sum": 0.0,
            "tax_sum": 0.0,
            "subtotal_sum": 0.0,
            "total_sum": 0.0,
        }
        for vendor in sorted(agg.keys()):
            data_vendor = agg[vendor]
            equip_sum = data_vendor.get("equip_sum", 0.0)
            other_sum = data_vendor.get("other_sum", 0.0)
            # Если включён режим цен с налогом, увеличиваем суммы
            with_tax_flag = bool(fin_opts.get("with_tax")) if fin_opts else False
            tax_map = fin_opts.get("vendor_tax", {}) if fin_opts else {}
            if with_tax_flag:
                try:
                    t = tax_map.get(normalize_case(vendor), 0.0)
                    equip_sum = equip_sum * (1.0 + t)
                    other_sum = other_sum * (1.0 + t)
                except Exception:
                    pass
            discount_pct = 0.0
            commission_pct = 0.0
            tax_pct = 0.0
            if ft:
                try:
                    discount_pct = ft.preview_discount_pct.get(vendor, 0.0)  # type: ignore
                    commission_pct = ft.preview_commission_pct.get(vendor, 0.0)  # type: ignore
                    tax_pct = ft.preview_tax_pct.get(vendor, 0.0)  # type: ignore
                except Exception:
                    pass
            if compute_client_flow:
                try:
                    discount_amt, commission_amt, tax_amt, subtotal, total_with_tax = compute_client_flow(
                        equip_sum,
                        other_sum,
                        discount_pct,
                        commission_pct,
                        tax_pct,
                    )
                except Exception:
                    discount_amt = commission_amt = tax_amt = 0.0
                    subtotal = equip_sum + other_sum
                    total_with_tax = subtotal
            else:
                discount_amt = commission_amt = tax_amt = 0.0
                subtotal = equip_sum + other_sum
                total_with_tax = subtotal
            # Флаги отображения
            if abs(discount_amt) >= 1e-6:
                include_discount = True
            if abs(commission_amt) >= 1e-6:
                include_commission = True
            vendors_info.append({
                "vendor": vendor,
                "equip_sum": equip_sum,
                "other_sum": other_sum,
                "discount_amt": discount_amt,
                "commission_amt": commission_amt,
                "tax_amt": tax_amt,
                "subtotal": subtotal,
                "total_with_tax": total_with_tax,
            })
            # Суммы в итоги
            grand_totals["equip_sum"] += equip_sum
            grand_totals["other_sum"] += other_sum
            grand_totals["discount_sum"] += discount_amt
            grand_totals["commission_sum"] += commission_amt
            # Если активен режим цен с налогом, налоги уже включены в equip_sum/other_sum, но для корректности всё равно
            grand_totals["tax_sum"] += tax_amt
            grand_totals["subtotal_sum"] += subtotal
            grand_totals["total_sum"] += total_with_tax

        # Если выбран режим «только зоны», формируем отчёт по зонам. В зависимости
        # от флага ``zones_by_vendor`` либо строится стандартная таблица сумм по
        # зонам, либо таблица с распределением сумм по подрядчикам внутри каждой зоны.
        if zones_only:
            if zones_by_vendor:
                try:
                    # Формируем список позиций для подсчёта сумм по зонам и подрядчикам
                    items_for_sum: List[Any] = []
                    if ft and hasattr(ft, "items"):
                        items_for_sum = ft.items
                    elif getattr(page, "db", None) and getattr(page, "project_id", None):
                        from .finance_tab import DBDataProvider  # type: ignore
                        prov = DBDataProvider(page)
                        items_for_sum = prov.load_items() or []
                    # Накапливаем суммы по классам (equipment/other) для каждой пары зона–подрядчик
                    zone_vendor_data: Dict[Tuple[str, str], Dict[str, float]] = {}
                    for it in items_for_sum:
                        try:
                            eff: Optional[float] = None
                            if ft:
                                v_name = it.vendor or ""
                                if it.cls == "equipment":
                                    if ft.preview_coeff_enabled.get(v_name, True):
                                        eff = float(ft._coeff_user_values.get(v_name, ft.preview_vendor_coeffs.get(v_name, 1.0)))
                                    else:
                                        if getattr(it, "original_coeff", None) is not None:
                                            eff = float(getattr(it, "original_coeff"))
                                        else:
                                            eff = None
                            amt = float(it.amount(effective_coeff=eff))
                            zone_name = (it.zone or "Без зоны").strip() or "Без зоны"
                            vendor_name = (it.vendor or "").strip()
                            key = (zone_name, vendor_name)
                            data = zone_vendor_data.setdefault(key, {"equip_sum": 0.0, "other_sum": 0.0})
                            if getattr(it, "cls", None) == "equipment":
                                data["equip_sum"] += amt
                            else:
                                data["other_sum"] += amt
                        except Exception:
                            continue
                    # Рассчитываем поток клиента для каждой пары зона–подрядчик
                    zone_vendor_flow: Dict[Tuple[str, str], Tuple[float, float, float]] = {}
                    for (zn, vn), sums_dict in zone_vendor_data.items():
                        eq_sum = sums_dict.get("equip_sum", 0.0)
                        oth_sum = sums_dict.get("other_sum", 0.0)
                        # Определяем ставки скидки, комиссии и налога для подрядчика.
                        discount_pct = 0.0
                        commission_pct = 0.0
                        tax_pct = 0.0
                        if ft:
                            try:
                                from .finance_tab import VendorSettings  # type: ignore
                            except Exception:
                                class VendorSettings:  # type: ignore
                                    discount_pct = 0.0
                                    commission_pct = 0.0
                                    tax_pct = 0.0
                            try:
                                discount_pct = ft.preview_discount_pct.get(
                                    vn,
                                    ft.vendors_settings.get(vn, VendorSettings()).discount_pct,  # type: ignore[attr-defined]
                                )
                                commission_pct = ft.preview_commission_pct.get(
                                    vn,
                                    ft.vendors_settings.get(vn, VendorSettings()).commission_pct,  # type: ignore[attr-defined]
                                )
                                tax_pct = ft.preview_tax_pct.get(
                                    vn,
                                    ft.vendors_settings.get(vn, VendorSettings()).tax_pct,  # type: ignore[attr-defined]
                                )
                            except Exception:
                                try:
                                    vs = VendorSettings()  # type: ignore
                                    discount_pct = getattr(vs, 'discount_pct', 0.0)
                                    commission_pct = getattr(vs, 'commission_pct', 0.0)
                                    tax_pct = getattr(vs, 'tax_pct', 0.0)
                                except Exception:
                                    discount_pct = commission_pct = tax_pct = 0.0
                        if compute_client_flow:
                            try:
                                _, _, tax_amt, subtotal, total_with_tax = compute_client_flow(
                                    eq_sum,
                                    oth_sum,
                                    discount_pct,
                                    commission_pct,
                                    tax_pct,
                                )
                            except Exception:
                                tax_amt = 0.0
                                subtotal = eq_sum + oth_sum
                                total_with_tax = subtotal
                        else:
                            tax_amt = 0.0
                            subtotal = eq_sum + oth_sum
                            total_with_tax = subtotal
                        zone_vendor_flow[(zn, vn)] = (subtotal, tax_amt, total_with_tax)
                    if zone_vendor_flow:
                        elements.append(Paragraph("Итоги по зонам", header_style))
                        zone_list = sorted(set([k[0] for k in zone_vendor_flow.keys()]))
                        for zn in zone_list:
                            zone_display = no_zone_label if (not zn or zn.strip() == "Без зоны") else zn
                            elements.append(Paragraph(f"{zone_display}", header_style))
                            elements.append(Spacer(1, 1 * mm))
                            header_row = ["Подрядчик", "Сумма", "Налог", "Сумма с налогом"]
                            zone_data: List[List[str]] = [header_row]
                            zone_row_colors: List[Optional[Any]] = []
                            vendors_for_zone = sorted([k[1] for k in zone_vendor_flow.keys() if k[0] == zn])
                            for vn in vendors_for_zone:
                                subtotal, tax_amt, total_with_tax = zone_vendor_flow.get((zn, vn), (0.0, 0.0, 0.0))
                                diff_str = ""
                                color: Optional[Any] = None
                                row = [
                                    vn or "(без подрядчика)",
                                    f"{subtotal:,.2f}".replace(",", " "),
                                    f"{tax_amt:,.2f}".replace(",", " "),
                                    f"{total_with_tax:,.2f}".replace(",", " ") + diff_str,
                                ]
                                zone_data.append(row)
                                zone_row_colors.append(color)
                            ncols = len(header_row)
                            col_widths = [FIN_TABLE_WIDTH_MM / max(ncols, 1) * mm] * ncols
                            tbl = Table(zone_data, repeatRows=1, colWidths=col_widths)
                            style_cmds: List[Tuple[Any, ...]] = [
                                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                            ]
                            for idx, c in enumerate(zone_row_colors):
                                if c:
                                    style_cmds.append(("TEXTCOLOR", (0, idx + 1), (-1, idx + 1), c))
                            tbl.setStyle(TableStyle(style_cmds))
                            elements.append(tbl)
                            elements.append(Spacer(1, 2 * mm))
                        return elements
                except Exception:
                    logger.error("Ошибка формирования итогов по зонам и подрядчикам", exc_info=True)
                    elements.append(Paragraph("Ошибка формирования итогов по зонам и подрядчикам", normal_style))
                    return elements
            # Стандартный режим суммирования по зонам (без разбивки по подрядчикам)
            try:
                # Формируем список позиций для подсчёта сумм по зонам
                items_for_sum: List[Any] = []
                if ft and hasattr(ft, "items"):
                    items_for_sum = ft.items
                elif getattr(page, "db", None) and getattr(page, "project_id", None):
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items_for_sum = prov.load_items() or []
                # Накапливаем суммы по каждой паре зона–подрядчик
                zone_vendor_data: Dict[Tuple[str, str], Dict[str, float]] = {}
                for it in items_for_sum:
                    try:
                        eff: Optional[float] = None
                        if ft:
                            v_name = it.vendor or ""
                            if it.cls == "equipment":
                                if ft.preview_coeff_enabled.get(v_name, True):
                                    eff = float(ft._coeff_user_values.get(v_name, ft.preview_vendor_coeffs.get(v_name, 1.0)))
                                else:
                                    if getattr(it, "original_coeff", None) is not None:
                                        eff = float(getattr(it, "original_coeff"))
                                    else:
                                        eff = None
                        amt = float(it.amount(effective_coeff=eff))
                        zone_name = (it.zone or "Без зоны").strip() or "Без зоны"
                        vendor_name = (it.vendor or "").strip()
                        key = (zone_name, vendor_name)
                        data = zone_vendor_data.setdefault(key, {"equip_sum": 0.0, "other_sum": 0.0})
                        if getattr(it, "cls", None) == "equipment":
                            data["equip_sum"] += amt
                        else:
                            data["other_sum"] += amt
                    except Exception:
                        continue
                # Рассчитываем поток клиента для каждой пары зона–подрядчик
                zone_vendor_flow: Dict[Tuple[str, str], Tuple[float, float, float]] = {}
                for (zn, vn), sums_dict in zone_vendor_data.items():
                    eq_sum = sums_dict.get("equip_sum", 0.0)
                    oth_sum = sums_dict.get("other_sum", 0.0)
                    # Определяем ставки скидки, комиссии и налога для подрядчика.
                    discount_pct = 0.0
                    commission_pct = 0.0
                    tax_pct = 0.0
                    if ft:
                        try:
                            from .finance_tab import VendorSettings  # type: ignore
                        except Exception:
                            class VendorSettings:  # type: ignore
                                discount_pct = 0.0
                                commission_pct = 0.0
                                tax_pct = 0.0
                        try:
                            discount_pct = ft.preview_discount_pct.get(
                                vn,
                                ft.vendors_settings.get(vn, VendorSettings()).discount_pct,  # type: ignore[attr-defined]
                            )
                            commission_pct = ft.preview_commission_pct.get(
                                vn,
                                ft.vendors_settings.get(vn, VendorSettings()).commission_pct,  # type: ignore[attr-defined]
                            )
                            tax_pct = ft.preview_tax_pct.get(
                                vn,
                                ft.vendors_settings.get(vn, VendorSettings()).tax_pct,  # type: ignore[attr-defined]
                            )
                        except Exception:
                            try:
                                vs = VendorSettings()  # type: ignore
                                discount_pct = getattr(vs, 'discount_pct', 0.0)
                                commission_pct = getattr(vs, 'commission_pct', 0.0)
                                tax_pct = getattr(vs, 'tax_pct', 0.0)
                            except Exception:
                                discount_pct = commission_pct = tax_pct = 0.0
                    if compute_client_flow:
                        try:
                            _, _, tax_amt, subtotal, total_with_tax = compute_client_flow(
                                eq_sum,
                                oth_sum,
                                discount_pct,
                                commission_pct,
                                tax_pct,
                            )
                        except Exception:
                            tax_amt = 0.0
                            subtotal = eq_sum + oth_sum
                            total_with_tax = subtotal
                    else:
                        tax_amt = 0.0
                        subtotal = eq_sum + oth_sum
                        total_with_tax = subtotal
                    zone_vendor_flow[(zn, vn)] = (subtotal, tax_amt, total_with_tax)
                # Агрегируем по зонам без разбивки по подрядчикам
                summary_zone_flow: Dict[str, Dict[str, float]] = {}
                for (zn, vn), (subtotal, tax_amt, total_with_tax) in zone_vendor_flow.items():
                    data = summary_zone_flow.setdefault(zn, {"subtotal": 0.0, "tax": 0.0, "total": 0.0})
                    data["subtotal"] += subtotal
                    data["tax"] += tax_amt
                    data["total"] += total_with_tax
                if summary_zone_flow:
                    elements.append(Paragraph("Итоги по зонам", header_style))
                    header = ["Зона", "Сумма", "Налог", "Сумма с налогом"]
                    zone_data: List[List[str]] = [header]
                    zone_row_colors: List[Optional[Any]] = []
                    for z, flows in sorted(summary_zone_flow.items()):
                        subtotal = flows.get("subtotal", 0.0)
                        tax_amt = flows.get("tax", 0.0)
                        total_with_tax = flows.get("total", 0.0)
                        diff_val: Optional[float] = None
                        color: Optional[Any] = None
                        diff_str = ""
                        if fin_snapshot:
                            try:
                                snap_total = snap_zones.get(z)  # type: ignore[name-defined]
                                curr_total = float(total_with_tax)
                                if snap_total is None:
                                    diff_val = curr_total
                                else:
                                    diff_val = curr_total - float(snap_total)
                                if diff_val is not None and abs(diff_val) > 0.01:
                                    diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                                    sign = "+" if diff_val > 0 else "-"
                                    diff_str = f" (" + sign + diff_fmt + ")"
                                    color = colors.red if diff_val > 0 else colors.green
                            except Exception:
                                diff_str = ""
                                color = None
                        zone_display = no_zone_label if (not z or z.strip() == "Без зоны") else z
                        row = [
                            zone_display,
                            f"{subtotal:,.2f}".replace(",", " "),
                            f"{tax_amt:,.2f}".replace(",", " "),
                            f"{total_with_tax:,.2f}".replace(",", " ") + diff_str,
                        ]
                        zone_data.append(row)
                        zone_row_colors.append(color)
                    ncols = len(header)
                    col_widths = [FIN_TABLE_WIDTH_MM / max(ncols, 1) * mm] * ncols
                    tbl = Table(zone_data, repeatRows=1, colWidths=col_widths)
                    style_cmds: List[Tuple[Any, ...]] = [
                        ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ]
                    for idx, c in enumerate(zone_row_colors):
                        if c:
                            style_cmds.append(("TEXTCOLOR", (0, idx + 1), (-1, idx + 1), c))
                    tbl.setStyle(TableStyle(style_cmds))
                    elements.append(tbl)
                    elements.append(Spacer(1, 2 * mm))
                return elements
            except Exception:
                logger.error("Ошибка формирования итогов по зонам (режим только зон)", exc_info=True)
                elements.append(Paragraph("Ошибка формирования итогов по зонам", normal_style))
                return elements
        # Строим таблицу для каждого подрядчика
        for info in vendors_info:
            vendor = info["vendor"]
            elements.append(Paragraph(textwrap.fill(vendor, width=40), header_style))
            elements.append(Spacer(1, 1.5 * mm))
            header_row = ["Оборудование", "Прочее"]
            value_row = [
                f"{info['equip_sum']:,.2f}".replace(",", " "),
                f"{info['other_sum']:,.2f}".replace(",", " "),
            ]
            if include_discount:
                header_row.append("Скидка")
                value_row.append(f"{info['discount_amt']:,.2f}".replace(",", " "))
            if include_commission:
                header_row.append("Комиссия")
                value_row.append(f"{info['commission_amt']:,.2f}".replace(",", " "))
            # Независимо от режима отображения цен выводим колонки «Налог», «Итого без налога» и
            # «Итого с налогом», чтобы показать структуру стоимости по каждому подрядчику.
            header_row.extend(["Налог", "Итого без налога", "Итого с налогом"])
            value_row.extend([
                f"{info['tax_amt']:,.2f}".replace(",", " "),
                f"{info['subtotal']:,.2f}".replace(",", " "),
                f"{info['total_with_tax']:,.2f}".replace(",", " "),
            ])
            # Если передан снимок, определяем дельту для подрядчика
            row_color: Optional[Any] = None
            if fin_snapshot:
                try:
                    snap_val = snap_vendors.get(vendor)  # type: ignore[name-defined]
                    curr_val = float(info.get("total_with_tax", 0.0))
                    diff_val: Optional[float]
                    if snap_val is None:
                        diff_val = curr_val
                    else:
                        diff_val = curr_val - float(snap_val)
                    if diff_val is not None and abs(diff_val) > 0.01:
                        diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                        sign = "+" if diff_val > 0 else "-"
                        # Добавляем строку дельты в последнюю колонку
                        value_row[-1] = value_row[-1] + f" (" + sign + diff_fmt + ")"
                        row_color = colors.red if diff_val > 0 else colors.green
                except Exception:
                    pass
            table_data = [header_row, value_row]
            # Ширина столбцов подбирается динамически: сумма FIN_TABLE_WIDTH_MM мм.
            try:
                ncols = len(header_row)
                col_w = FIN_TABLE_WIDTH_MM / max(ncols, 1)
                vendor_table = Table(table_data, repeatRows=1, colWidths=[col_w * mm] * ncols)
            except Exception:
                vendor_table = Table(table_data, repeatRows=1)
            # Базовый стиль таблицы подрядчика
            style_cmds: List[Tuple[Any, ...]] = [
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ]
            if row_color:
                # Красим всю строку данных выбранным цветом
                style_cmds.append(("TEXTCOLOR", (0, 1), (-1, 1), row_color))
            vendor_table.setStyle(TableStyle(style_cmds))
            elements.append(vendor_table)
            elements.append(Spacer(1, 3 * mm))
        # Итоговый блок по всем подрядчикам
        elements.append(Paragraph("Итого по всем", header_style))
        elements.append(Spacer(1, 1.5 * mm))
        header_row = ["Оборудование", "Прочее"]
        value_row = [
            f"{grand_totals['equip_sum']:,.2f}".replace(",", " "),
            f"{grand_totals['other_sum']:,.2f}".replace(",", " "),
        ]
        if include_discount:
            header_row.append("Скидка")
            value_row.append(f"{grand_totals['discount_sum']:,.2f}".replace(",", " "))
        if include_commission:
            header_row.append("Комиссия")
            value_row.append(f"{grand_totals['commission_sum']:,.2f}".replace(",", " "))
        # Всегда выводим колонки «Налог», «Итого без налога» и «Итого с налогом»
        header_row.extend(["Налог", "Итого без налога", "Итого с налогом"])
        value_row.extend([
            f"{grand_totals['tax_sum']:,.2f}".replace(",", " "),
            f"{grand_totals['subtotal_sum']:,.2f}".replace(",", " "),
            f"{grand_totals['total_sum']:,.2f}".replace(",", " "),
        ])
        # Подсветка общей строки, если передан снимок
        totals_row_color: Optional[Any] = None
        if fin_snapshot:
            try:
                curr_total = float(grand_totals.get('total_sum', 0.0))
                snap_total = snap_project_total  # type: ignore[name-defined]
                diff_val: Optional[float]
                if snap_total is None:
                    diff_val = curr_total
                else:
                    diff_val = curr_total - float(snap_total)
                if diff_val is not None and abs(diff_val) > 0.01:
                    diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                    sign = "+" if diff_val > 0 else "-"
                    value_row[-1] = value_row[-1] + f" (" + sign + diff_fmt + ")"
                    totals_row_color = colors.red if diff_val > 0 else colors.green
            except Exception:
                pass
        totals_table_data = [header_row, value_row]
        try:
            ncols = len(header_row)
            col_w = FIN_TABLE_WIDTH_MM / max(ncols, 1)
            totals_table = Table(totals_table_data, repeatRows=1, colWidths=[col_w * mm] * ncols)
        except Exception:
            totals_table = Table(totals_table_data, repeatRows=1)
        # Базовый стиль итоговой таблицы
        totals_style_cmds: List[Tuple[Any, ...]] = [
            ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans-Bold"),
            ("BACKGROUND", (0, 1), (-1, 1), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ]
        if totals_row_color:
            totals_style_cmds.append(("TEXTCOLOR", (0, 1), (-1, 1), totals_row_color))
        totals_table.setStyle(TableStyle(totals_style_cmds))
        elements.append(totals_table)
        elements.append(Spacer(1, 4 * mm))

        # Дополнительные итоги: по подрядчикам, зонам, отделам и классам
        try:
            # Загружаем список позиций через виджет бухгалтерии или провайдера
            items_for_totals: List[Any] = []
            ft_items: List[Any] = []
            if ft and hasattr(ft, "items"):
                ft_items = ft.items
                items_for_totals = ft.items
            elif getattr(page, "db", None) and getattr(page, "project_id", None):
                from .finance_tab import DBDataProvider  # type: ignore
                prov = DBDataProvider(page)
                items_for_totals = prov.load_items() or []
            # 2.3 Итоги по подрядчикам, зонам, отделам и классам
            # Собираем суммы на основе агрегированных данных из FinanceTab, если доступен
            summary_vendor: Dict[str, float] = {}
            summary_vendor_tax: Dict[str, Tuple[float, float]] = {}
            summary_zone: Dict[str, float] = {}
            summary_zone_tax: Dict[str, Tuple[float, float]] = {}
            summary_dept: Dict[str, float] = {}
            summary_dept_tax: Dict[str, Tuple[float, float]] = {}
            summary_cls: Dict[str, float] = {}
            summary_cls_tax: Dict[str, Tuple[float, float]] = {}
            total_project = 0.0
            total_project_tax = 0.0
            # Определяем, включено ли отображение налогов
            with_tax_flag = bool(fin_opts.get("with_tax")) if fin_opts else False
            # Словарь налогов по подрядчикам для использования ниже
            vendor_tax_map: Dict[str, float] = {}
            if with_tax_flag:
                # если виджет бухгалтерии доступен, используем его настройки налогов
                try:
                    if ft and hasattr(ft, "preview_tax_pct"):
                        for v in ft.preview_tax_pct:
                            vendor_tax_map[normalize_case(v)] = float(ft.preview_tax_pct.get(v, 0.0)) / 100.0
                    else:
                        vendor_tax_map = fin_opts.get("vendor_tax", {}) or {}
                except Exception:
                    vendor_tax_map = {}
            # Если имеется агрегатор из вкладки "Бухгалтерия", используем его для расчёта сумм по подрядчикам
            if ft and hasattr(ft, "_agg_latest") and ft._agg_latest:
                agg = ft._agg_latest  # type: ignore
                for vend, data in agg.items():
                    try:
                        total_sum = float(data.get("equip_sum", 0.0)) + float(data.get("other_sum", 0.0))
                        summary_vendor[vend] = summary_vendor.get(vend, 0.0) + total_sum
                        if with_tax_flag:
                            t_pct = vendor_tax_map.get(normalize_case(vend), 0.0)
                            tax_amt = total_sum * t_pct
                            summary_vendor_tax[vend] = (summary_vendor_tax.get(vend, (0.0, 0.0))[0] + tax_amt,
                                                        summary_vendor_tax.get(vend, (0.0, 0.0))[1] + total_sum + tax_amt)
                    except Exception:
                        continue
            # Для зон, отделов и классов (и суммы проекта) требуется пройти по каждой позиции с учётом эффективного коэффициента
            # Если доступен виджет ft, повторяем логику recalculate_all для определения коэффициентов
            for it in items_for_totals:
                try:
                    # Определяем эффективный коэффициент только для equipment
                    eff: Optional[float] = None
                    if ft:
                        v = it.vendor or ""
                        if it.cls == "equipment":
                            if ft.preview_coeff_enabled.get(v, True):
                                eff = float(ft._coeff_user_values.get(v, ft.preview_vendor_coeffs.get(v, 1.0)))
                            else:
                                # Используем исходный коэффициент позиции (original_coeff)
                                if getattr(it, "original_coeff", None) is not None:
                                    eff = float(getattr(it, "original_coeff"))
                                else:
                                    eff = None
                    # Сумма позиции (без учёта скидок/комиссий клиента, как на вкладке Бухгалтерия)
                    amt = float(it.amount(effective_coeff=eff))
                    # Обновляем общую сумму проекта
                    total_project += amt
                    # Имена зон/отделов/классов (без нормализации, чтобы сохранить оригинальный регистр)
                    zone_name = (it.zone or "Без зоны").strip() or "Без зоны"
                    dept_name = (it.department or "Без отдела").strip() or "Без отдела"
                    cls_en = getattr(it, "cls", "equipment")
                    # Сохраняем суммы по зонам, отделам и классам
                    summary_zone[zone_name] = summary_zone.get(zone_name, 0.0) + amt
                    summary_dept[dept_name] = summary_dept.get(dept_name, 0.0) + amt
                    # Учитываем сумму по любому классу (включая equipment)
                    summary_cls[cls_en] = summary_cls.get(cls_en, 0.0) + amt
                    # Считаем налог для этой позиции по ставке клиента (preview_tax_pct).
                    # Налог рассчитывается всегда, чтобы сформировать колонки «Налог» и
                    # «Сумма с налогом» для зон, отделов и классов.
                    try:
                        client_tax_pct = 0.0
                        if ft and hasattr(ft, "preview_tax_pct"):
                            try:
                                client_tax_pct = float(ft.preview_tax_pct.get(it.vendor or "", 0.0)) / 100.0  # type: ignore
                            except Exception:
                                client_tax_pct = 0.0
                        tax_amt = amt * client_tax_pct
                        # Обновляем суммарные налоги для зон, отделов и классов
                        prev_tax, prev_sum_tax = summary_zone_tax.get(zone_name, (0.0, 0.0))
                        summary_zone_tax[zone_name] = (prev_tax + tax_amt, prev_sum_tax + amt + tax_amt)
                        prev_tax, prev_sum_tax = summary_dept_tax.get(dept_name, (0.0, 0.0))
                        summary_dept_tax[dept_name] = (prev_tax + tax_amt, prev_sum_tax + amt + tax_amt)
                        # Обновляем налог и сумму с налогом для любого класса (включая equipment)
                        prev_tax, prev_sum_tax = summary_cls_tax.get(cls_en, (0.0, 0.0))
                        summary_cls_tax[cls_en] = (prev_tax + tax_amt, prev_sum_tax + amt + tax_amt)
                        # Общая сумма налога по проекту оставлена для совместимости с другими разделами
                        total_project_tax += tax_amt
                    except Exception:
                        pass
                except Exception:
                    continue
            # Выводим краткие итоги только если не выбраны флаги подробного отчёта
            if (
                not fin_opts.get("show_internal", False)
                and not fin_opts.get("show_agents", False)
                and not fin_opts.get("internal_only", False)
            ):
                # 2.4 Вывод таблицы подрядчиков
                if summary_vendor:
                    try:
                        # Не выводим таблицу «Итоги по подрядчикам» и общую стоимость проекта,
                        # поскольку подробные данные по подрядчикам уже приведены выше, а итог
                        # проекта выводится в шапке отчёта. Этот блок оставлен пустым, чтобы
                        # сохранить вычисление сумм для последующих итогов (зоны, отделы, классы).
                        pass
                    except Exception:
                        logger.error("Ошибка формирования итогов по подрядчикам", exc_info=True)
                # 2.4b Итоги по зонам
                if summary_zone:
                    try:
                        elements.append(Paragraph("Итоги по зонам", header_style))
                        # Всегда выводим колонки «Налог» и «Сумма с налогом» для зон
                        header = ["Зона", "Сумма", "Налог", "Сумма с налогом"]
                        zone_data: List[List[str]] = [header]
                        # Для каждой зоны определяем дельту и цвет строки
                        zone_row_colors: List[Optional[Any]] = []
                        for z, amt in sorted(summary_zone.items()):
                            tax_amt, sum_taxed = summary_zone_tax.get(z, (0.0, amt))
                            diff_val: Optional[float] = None
                            color: Optional[Any] = None
                            diff_str = ""
                            if fin_snapshot:
                                try:
                                    snap_total = snap_zones.get(z)  # type: ignore[name-defined]
                                    curr_total = float(sum_taxed)
                                    if snap_total is None:
                                        diff_val = curr_total
                                    else:
                                        diff_val = curr_total - float(snap_total)
                                    if diff_val is not None and abs(diff_val) > 0.01:
                                        diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                                        sign = "+" if diff_val > 0 else "-"
                                        diff_str = f" (" + sign + diff_fmt + ")"
                                        color = colors.red if diff_val > 0 else colors.green
                                except Exception:
                                    diff_str = ""
                                    color = None
                            # Отображаем имя зоны: если имя пустое или равно «Без зоны»,
                            # подставляем пользовательское название. Иначе используем исходное
                            zone_display = no_zone_label if (not z or z.strip() == "Без зоны") else z
                            row = [
                                zone_display,
                                f"{amt:,.2f}".replace(",", " "),
                                f"{tax_amt:,.2f}".replace(",", " "),
                                f"{sum_taxed:,.2f}".replace(",", " ") + diff_str,
                            ]
                            zone_data.append(row)
                            zone_row_colors.append(color)
                        ncols = len(header)
                        col_widths = [FIN_TABLE_WIDTH_MM / max(ncols, 1) * mm] * ncols
                        tbl = Table(zone_data, repeatRows=1, colWidths=col_widths)
                        # Базовый стиль
                        style_cmds: List[Tuple[Any, ...]] = [
                            ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                            ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ]
                        # Применяем цвета строк при наличии дельт
                        for idx, c in enumerate(zone_row_colors):
                            if c:
                                style_cmds.append(("TEXTCOLOR", (0, idx + 1), (-1, idx + 1), c))
                        tbl.setStyle(TableStyle(style_cmds))
                        elements.append(tbl)
                        elements.append(Spacer(1, 2 * mm))
                    except Exception:
                        logger.error("Ошибка формирования итогов по зонам", exc_info=True)
                # 2.4c Итоги по отделам
                if summary_dept:
                    try:
                        elements.append(Paragraph("Итоги по отделам", header_style))
                        header = ["Отдел", "Сумма", "Налог", "Сумма с налогом"]
                        dept_data: List[List[str]] = [header]
                        dept_row_colors: List[Optional[Any]] = []
                        for d, amt in sorted(summary_dept.items()):
                            tax_amt, sum_taxed = summary_dept_tax.get(d, (0.0, amt))
                            diff_val: Optional[float] = None
                            color: Optional[Any] = None
                            diff_str = ""
                            if fin_snapshot:
                                try:
                                    snap_total = snap_depts.get(d)  # type: ignore[name-defined]
                                    curr_total = float(sum_taxed)
                                    if snap_total is None:
                                        diff_val = curr_total
                                    else:
                                        diff_val = curr_total - float(snap_total)
                                    if diff_val is not None and abs(diff_val) > 0.01:
                                        diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                                        sign = "+" if diff_val > 0 else "-"
                                        diff_str = f" (" + sign + diff_fmt + ")"
                                        color = colors.red if diff_val > 0 else colors.green
                                except Exception:
                                    diff_str = ""
                                    color = None
                            row = [
                                d or "Без отдела",
                                f"{amt:,.2f}".replace(",", " "),
                                f"{tax_amt:,.2f}".replace(",", " "),
                                f"{sum_taxed:,.2f}".replace(",", " ") + diff_str,
                            ]
                            dept_data.append(row)
                            dept_row_colors.append(color)
                        ncols = len(header)
                        col_widths = [FIN_TABLE_WIDTH_MM / max(ncols, 1) * mm] * ncols
                        tbl = Table(dept_data, repeatRows=1, colWidths=col_widths)
                        style_cmds: List[Tuple[Any, ...]] = [
                            ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                            ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ]
                        for idx, c in enumerate(dept_row_colors):
                            if c:
                                style_cmds.append(("TEXTCOLOR", (0, idx + 1), (-1, idx + 1), c))
                        tbl.setStyle(TableStyle(style_cmds))
                        elements.append(tbl)
                        elements.append(Spacer(1, 2 * mm))
                    except Exception:
                        logger.error("Ошибка формирования итогов по отделам", exc_info=True)
                # 2.4d Итоги по классам
                if summary_cls:
                    try:
                        elements.append(Paragraph("Итоги по классам", header_style))
                        header = ["Класс", "Сумма", "Налог", "Сумма с налогом"]
                        cls_data: List[List[str]] = [header]
                        cls_row_colors: List[Optional[Any]] = []
                        for cls_en, amt in summary_cls.items():
                            cls_ru = CLASS_EN2RU.get(cls_en, cls_en.capitalize())
                            tax_amt, sum_taxed = summary_cls_tax.get(cls_en, (0.0, amt))
                            diff_val: Optional[float] = None
                            color: Optional[Any] = None
                            diff_str = ""
                            if fin_snapshot:
                                try:
                                    snap_total = snap_classes.get(cls_en)  # type: ignore[name-defined]
                                    curr_total = float(sum_taxed)
                                    if snap_total is None:
                                        diff_val = curr_total
                                    else:
                                        diff_val = curr_total - float(snap_total)
                                    if diff_val is not None and abs(diff_val) > 0.01:
                                        diff_fmt = f"{abs(diff_val):,.2f}".replace(",", " ")
                                        sign = "+" if diff_val > 0 else "-"
                                        diff_str = f" (" + sign + diff_fmt + ")"
                                        color = colors.red if diff_val > 0 else colors.green
                                except Exception:
                                    diff_str = ""
                                    color = None
                            row = [
                                cls_ru,
                                f"{amt:,.2f}".replace(",", " "),
                                f"{tax_amt:,.2f}".replace(",", " "),
                                f"{sum_taxed:,.2f}".replace(",", " ") + diff_str,
                            ]
                            cls_data.append(row)
                            cls_row_colors.append(color)
                        ncols = len(header)
                        col_widths = [FIN_TABLE_WIDTH_MM / max(ncols, 1) * mm] * ncols
                        tbl = Table(cls_data, repeatRows=1, colWidths=col_widths)
                        style_cmds: List[Tuple[Any, ...]] = [
                            ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                            ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ]
                        for idx, c in enumerate(cls_row_colors):
                            if c:
                                style_cmds.append(("TEXTCOLOR", (0, idx + 1), (-1, idx + 1), c))
                        tbl.setStyle(TableStyle(style_cmds))
                        elements.append(tbl)
                        elements.append(Spacer(1, 4 * mm))
                    except Exception:
                        logger.error("Ошибка формирования итогов по классам", exc_info=True)
        except Exception:
            logger.error("Ошибка формирования итогов по подрядчикам/зонам/отделам/классам", exc_info=True)
        # Внутренние расчёты и комиссия для агентств
        if fin_opts.get("show_internal", False):
            try:
                if ft and hasattr(ft, "_calc_income_total") and hasattr(ft, "_calc_expense_total"):
                    income_total = ft._calc_income_total(agg)  # type: ignore
                    expense_total = ft._calc_expense_total()  # type: ignore
                else:
                    income_total = grand_totals["subtotal_sum"]
                    expense_total = 0.0
                net_total = income_total - expense_total
                elements.append(Paragraph(f"Доходы: {income_total:,.2f} ₽".replace(",", " "), normal_style))
                elements.append(Paragraph(f"Расходы: {expense_total:,.2f} ₽".replace(",", " "), normal_style))
                elements.append(Paragraph(f"Чистая прибыль: {net_total:,.2f} ₽".replace(",", " "), normal_style))
                elements.append(Spacer(1, 4 * mm))
            except Exception:
                logger.error("Ошибка расчёта внутренней информации", exc_info=True)
        # Комиссии
        if fin_opts.get("show_agents", False):
            # Выводим комиссию каждого подрядчика по оборудованию
            data2: List[List[Any]] = []
            data2.append(["Подрядчик", "Комиссия %", "Сумма комиссии"])
            for v, d in agg.items():
                try:
                    pct = 0.0
                    if ft and hasattr(ft, "preview_commission_pct"):
                        pct = ft.preview_commission_pct.get(v, 0.0)  # type: ignore
                    comm = d.get("equip_sum", 0.0) * (pct / 100.0)
                    data2.append([
                        v,
                        f"{pct:.2f}".rstrip("0").rstrip(".") + "%",
                        f"{comm:,.2f}".replace(",", " "),
                    ])
                except Exception:
                    continue
            table2 = Table(data2, repeatRows=1)
            table2.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))
            elements.append(table2)
        return elements
    except Exception:
        logger.error("Ошибка создания финансового отчёта", exc_info=True)
        return elements


def _compute_fin_report_totals(page: Any, fin_opts: Dict[str, Any]) -> Tuple[float, float, float]:
    """
    Вычисляет суммарные значения для финансового отчёта.

    На основе агрегированных данных по подрядчикам подсчитывает:
        • итоговую сумму до налога;
        • сумму налога;
        • итоговую сумму с налогом.

    Эти значения используются для вывода в шапке финансового отчёта, чтобы
    пользователь сразу видел общую стоимость, размер налога и итог с налогом.

    :param page: страница проекта, содержащая виджет «Бухгалтерия» и данные
    :param fin_opts: словарь опций финансового отчёта, содержащий флаг
        ``with_tax`` и карту ``vendor_tax``
    :return: кортеж (subtotal_sum, tax_sum, total_sum)
    """
    try:
        # Получаем виджет бухгалтерии для доступа к агрегированным данным и
        # настройкам скидок/комиссий/налогов
        ft = getattr(page, "tab_finance_widget", None)
        # Собираем агрегированные данные: сумма оборудования и прочего по каждому подрядчику
        agg: Dict[str, Dict[str, float]] = {}
        # 1. Если FinanceTab уже сформировал агрегатор _agg_latest, используем его
        if ft and hasattr(ft, "_agg_latest") and ft._agg_latest:
            agg = ft._agg_latest  # type: ignore
        else:
            # 2. Иначе строим агрегатор самостоятельно на основе списка позиций
            items_for_totals: List[Any] = []
            if ft and hasattr(ft, "items"):
                items_for_totals = list(ft.items)
            elif getattr(page, "db", None) and getattr(page, "project_id", None):
                try:
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items_for_totals = prov.load_items() or []
                except Exception:
                    items_for_totals = []
            for it in items_for_totals:
                try:
                    v = getattr(it, "vendor", "") or "(без подрядчика)"
                    data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
                    amt = 0.0
                    try:
                        # Позиция рассчитывает сумму через метод amount()
                        amt = float(it.amount())
                    except Exception:
                        amt = 0.0
                    try:
                        cls = getattr(it, "cls", "")
                    except Exception:
                        cls = ""
                    if cls == "equipment":
                        data["equip_sum"] += amt
                    else:
                        data["other_sum"] += amt
                    agg[v] = data
                except Exception:
                    continue
        # 3. Определяем, нужно ли включать налог в суммы
        with_tax_flag = bool(fin_opts.get("with_tax")) if fin_opts else False
        # Словарь налогов по подрядчикам: normalize_case(vendor) -> rate
        vendor_tax_map: Dict[str, float] = {}
        if with_tax_flag:
            try:
                vendor_tax_map = fin_opts.get("vendor_tax", {}) or {}
            except Exception:
                vendor_tax_map = {}
        # 4. Инициализируем итоговые суммы
        grand_totals = {
            "equip_sum": 0.0,
            "other_sum": 0.0,
            "discount_sum": 0.0,
            "commission_sum": 0.0,
            "tax_sum": 0.0,
            "subtotal_sum": 0.0,
            "total_sum": 0.0,
        }
        # Импортируем функцию расчёта клиентского потока отдельно, так как она
        # может отсутствовать (например, если модуль finance_tab не загружен)
        try:
            from .finance_tab import compute_client_flow  # type: ignore
        except Exception:
            compute_client_flow = None  # type: ignore
        # Обходим каждого подрядчика и обновляем итоговые суммы
        for vendor in sorted(agg.keys()):
            data_vendor = agg[vendor]
            equip_sum = float(data_vendor.get("equip_sum", 0.0))
            other_sum = float(data_vendor.get("other_sum", 0.0))
            # Корректируем суммы, если выбран режим «с налогом»: для каждого
            # подрядчика добавляем его налог к оборудованию и прочему
            if with_tax_flag:
                try:
                    t_rate = float(vendor_tax_map.get(normalize_case(vendor), 0.0))
                    equip_sum = equip_sum * (1.0 + t_rate)
                    other_sum = other_sum * (1.0 + t_rate)
                except Exception:
                    pass
            # Проценты скидки, комиссии и налога для клиента (из FinanceTab)
            discount_pct = 0.0
            commission_pct = 0.0
            tax_pct = 0.0
            if ft:
                try:
                    discount_pct = float(ft.preview_discount_pct.get(vendor, 0.0))  # type: ignore
                except Exception:
                    discount_pct = 0.0
                try:
                    commission_pct = float(ft.preview_commission_pct.get(vendor, 0.0))  # type: ignore
                except Exception:
                    commission_pct = 0.0
                try:
                    tax_pct = float(ft.preview_tax_pct.get(vendor, 0.0))  # type: ignore
                except Exception:
                    tax_pct = 0.0
            # Выполняем расчёт по схеме «Сумма → минус скидка → минус комиссия → плюс налог»
            if compute_client_flow:
                try:
                    discount_amt, commission_amt, tax_amt, subtotal, total_with_tax = compute_client_flow(
                        equip_sum,
                        other_sum,
                        discount_pct,
                        commission_pct,
                        tax_pct,
                    )
                except Exception:
                    discount_amt = commission_amt = tax_amt = 0.0
                    subtotal = equip_sum + other_sum
                    total_with_tax = subtotal
            else:
                # Если функция отсутствует, считаем без скидок, комиссий и налогов
                discount_amt = commission_amt = tax_amt = 0.0
                subtotal = equip_sum + other_sum
                total_with_tax = subtotal
            # Обновляем итоговые суммы
            grand_totals["equip_sum"] += equip_sum
            grand_totals["other_sum"] += other_sum
            grand_totals["discount_sum"] += discount_amt
            grand_totals["commission_sum"] += commission_amt
            grand_totals["tax_sum"] += tax_amt
            grand_totals["subtotal_sum"] += subtotal
            grand_totals["total_sum"] += total_with_tax
        # Возвращаем кортеж итоговых значений: до налога, налог, с налогом
        return (
            grand_totals.get("subtotal_sum", 0.0),
            grand_totals.get("tax_sum", 0.0),
            grand_totals.get("total_sum", 0.0),
        )
    except Exception:
        # Логируем ошибку и возвращаем нули
        logger.error("Ошибка вычисления итогов для финансового отчёта", exc_info=True)
        return (0.0, 0.0, 0.0)


# ---------------------------------------------------------------------------
# Специальный отчёт «для Ксюши».  Формирует упрощённую таблицу, где для
# каждого подрядчика выводятся две суммы: «Сумма по смете» (с учётом
# скидок и налога, но без учёта комиссии) и «Сумма к оплате» (с учётом
# скидок, комиссии и налога).  В начале отчёта выводятся общие итоги
# проекта для этих сумм.  Остальные параметры финансового отчёта
# (например, сортировка по зонам, отображение внутренней информации) не
# используются.

def _build_fin_report_ksyusha(
    page: Any,
    fin_opts: Dict[str, Any],
    header_style: ParagraphStyle,
    normal_style: ParagraphStyle,
) -> List[Any]:
    """Строит упрощённый финансовый отчёт для Ксюши.

    :param page: объект страницы проекта
    :param fin_opts: словарь опций (используется только флаг with_tax для выбора налогов)
    :param header_style: стиль заголовка
    :param normal_style: стиль обычного текста
    :return: список элементов PDF‑отчёта
    """
    elements: List[Any] = []
    try:
        # Получаем виджет бухгалтерии для доступа к агрегированным данным
        ft = getattr(page, "tab_finance_widget", None)
        # Собираем агрегированные данные по подрядчикам: суммы оборудования и прочего
        agg: Dict[str, Dict[str, float]] = {}
        if ft and hasattr(ft, "_agg_latest") and ft._agg_latest:
            agg = ft._agg_latest  # type: ignore
        else:
            # Создаём агрегатор самостоятельно, проходя по всем позициям проекта
            items: List[Any] = []
            if ft and hasattr(ft, "items"):
                items = list(ft.items)
            elif getattr(page, "db", None) and getattr(page, "project_id", None):
                try:
                    from .finance_tab import DBDataProvider  # type: ignore
                    prov = DBDataProvider(page)
                    items = prov.load_items() or []
                except Exception:
                    items = []
            for it in items:
                try:
                    v = getattr(it, "vendor", "") or "(без подрядчика)"
                    data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
                    amt = 0.0
                    try:
                        amt = float(it.amount())
                    except Exception:
                        amt = 0.0
                    try:
                        cls = getattr(it, "cls", "")
                    except Exception:
                        cls = ""
                    if cls == "equipment":
                        data["equip_sum"] += amt
                    else:
                        data["other_sum"] += amt
                    agg[v] = data
                except Exception:
                    continue
        # Получаем процент скидки, комиссии и налога для каждого подрядчика из FinanceTab
        discount_map: Dict[str, float] = {}
        commission_map: Dict[str, float] = {}
        tax_map: Dict[str, float] = {}
        if ft:
            try:
                discount_map = {normalize_case(v): float(ft.preview_discount_pct.get(v, 0.0)) / 100.0 for v in ft.preview_discount_pct}
            except Exception:
                discount_map = {}
            try:
                commission_map = {normalize_case(v): float(ft.preview_commission_pct.get(v, 0.0)) / 100.0 for v in ft.preview_commission_pct}
            except Exception:
                commission_map = {}
            try:
                tax_map = {normalize_case(v): float(ft.preview_tax_pct.get(v, 0.0)) / 100.0 for v in ft.preview_tax_pct}
            except Exception:
                tax_map = {}
        # Составляем таблицу данных по каждому подрядчику
        table_data: List[List[str]] = [["Подрядчик", "Сумма по смете ₽", "Сумма к оплате ₽"]]
        total_smeta: float = 0.0
        total_pay: float = 0.0
        for vendor in sorted(agg.keys()):
            data = agg[vendor]
            # Базовые суммы: equip_sum для класса «equipment», other_sum — остальные классы.
            equip_sum = float(data.get("equip_sum", 0.0))
            other_sum = float(data.get("other_sum", 0.0))
            # Проценты скидки, комиссии и налога берутся из карт. Значения уже приведены
            # к долям (0.15 соответствует 15 %). Для отсутствующих записей используется 0.
            disc_pct = discount_map.get(normalize_case(vendor), 0.0)
            comm_pct = commission_map.get(normalize_case(vendor), 0.0)
            tax_pct = tax_map.get(normalize_case(vendor), 0.0)

            # --- Расчёт сумм с учётом скидки, комиссии и налога ---
            # 1. Скидка и комиссия применяются только к сумме класса equipment. Сумма other_sum
            #    не изменяется.
            # 2. Сначала вычитаем скидку, затем комиссию, затем добавляем налог.

            # Сумма скидки на оборудование
            discount_amount = equip_sum * disc_pct
            # Стоимость оборудования после скидки
            after_discount_equip = equip_sum - discount_amount
            # Сумма комиссии на оборудование (на сумму после скидки)
            commission_amount = after_discount_equip * comm_pct
            # Стоимость оборудования после вычета комиссии
            after_commission_equip = after_discount_equip - commission_amount

            # Стоимость по смете (без комиссии, но со скидкой) до налога
            smeta_before_tax = after_discount_equip + other_sum
            # Стоимость по смете с учётом налога
            smeta_sum = smeta_before_tax * (1.0 + tax_pct)

            # Итоговая сумма к оплате: после скидки и комиссии, затем налог
            pay_before_tax = after_commission_equip + other_sum
            pay_sum = pay_before_tax * (1.0 + tax_pct)

            # Накопление общих итогов
            total_smeta += smeta_sum
            total_pay += pay_sum
            # Форматируем строки: используем пробелы для отделения тысяч
            try:
                smeta_str = f"{smeta_sum:,.2f}".replace(",", " ")
            except Exception:
                smeta_str = f"{smeta_sum:.2f}"
            try:
                pay_str = f"{pay_sum:,.2f}".replace(",", " ")
            except Exception:
                pay_str = f"{pay_sum:.2f}"
            table_data.append([
                vendor,
                smeta_str,
                pay_str,
            ])
        # Формируем заголовок с итогами
        try:
            smeta_total_str = f"{total_smeta:,.2f}".replace(",", " ")
        except Exception:
            smeta_total_str = f"{total_smeta:.2f}"
        try:
            pay_total_str = f"{total_pay:,.2f}".replace(",", " ")
        except Exception:
            pay_total_str = f"{total_pay:.2f}"
        header_line = f"Итого: {smeta_total_str} ₽; Итого к оплате: {pay_total_str} ₽"
        elements.append(Paragraph(header_line, header_style))
        elements.append(Spacer(1, 4 * mm))
        # Создаём таблицу ReportLab
        col_widths = [80 * mm, 50 * mm, 50 * mm]
        try:
            table = Table(table_data, repeatRows=1, colWidths=col_widths)
            from reportlab.platypus import TableStyle as LocalTableStyle
            table.setStyle(LocalTableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
            ]))
        except Exception:
            table = Table(table_data, repeatRows=1)
        elements.append(table)
    except Exception:
        logger.error("Ошибка создания упрощённого финансового отчёта (Ксюша)", exc_info=True)
    return elements


def _build_timing_report(
    page: Any,
    header_style: ParagraphStyle,
    normal_style: ParagraphStyle,
    opts: Optional[Dict[str, Any]] = None,
) -> List[Any]:
    """
    Формирует отчёт тайминга с учётом выбранных опций.

    Опции ``opts`` могут содержать:
        • ``show_table`` (bool) — использовать первый графический
          вариант (горизонтальная шкала без учёта колонок). Совместимо
          с ``separate_columns``;
        • ``separate_columns`` (bool) — формировать отдельные таблицы по
          каждому столбцу (колонке) тайминга;
        • ``graphic2`` (bool) — второй графический вариант: горизонтальная
          шкала с разделением по колонкам (каждая колонка — отдельная
          строка);
        • ``graphic3`` (bool) — третий графический вариант: «дерево
          времени», где вертикальная временная шкала связана с
          горизонтальными ветвями для каждой колонки.

    :param page: текущая страница проекта
    :param header_style: стиль для заголовков таблицы
    :param normal_style: стиль для обычного текста
    :param opts: словарь опций тайминга
    :return: список элементов для вставки в PDF
    """
    elements: List[Any] = []
    try:
        import math  # для округления количества ячеек при дробной длительности
        # Опции отображения
        show_table = True
        separate_cols = False
        if opts:
            show_table = bool(opts.get("show_table", True))
            separate_cols = bool(opts.get("separate_columns", False))
        # Если не выбраны таблица, ни разделение по колонкам, ни графические варианты,
        # то ничего не выводим. Для графических вариантов (graphic2/graphic3) блокируем
        # преждевременный выход, даже если show_table=False.
        try:
            graphic2_flag = bool(opts.get("graphic2", False)) if opts else False
            graphic3_flag = bool(opts.get("graphic3", False)) if opts else False
        except Exception:
            graphic2_flag = False
            graphic3_flag = False
        if not separate_cols and not show_table and not graphic2_flag and not graphic3_flag:
            return elements
        blocks = getattr(page, "timing_blocks", None)
        if not blocks:
            elements.append(Paragraph("Данные тайминга отсутствуют или не заполнены.", normal_style))
            return elements
        # Параметры сетки
        start_date = getattr(page, "timing_start_date", None)
        step_minutes = getattr(page, "timing_step_minutes", 60)
        column_names = getattr(page, "timing_column_names", [])
        units_per_day = getattr(page, "timing_units_per_day", None)
        from datetime import datetime, timedelta
        def qdate_to_date(qd: Any) -> datetime.date:
            try:
                return datetime(qd.year(), qd.month(), qd.day()).date()
            except Exception:
                return datetime.now().date()
        base_date = qdate_to_date(start_date) if start_date else datetime.now().date()
        # Если units_per_day отсутствует, вычисляем на основе шага
        if units_per_day is None:
            try:
                units_per_day = int((24 * 60) / max(1, int(step_minutes)))
            except Exception:
                units_per_day = 24  # fallback на 24 часа по часу
        # Вариант 1: разбивка по колонкам (отдельные таблицы)
        if separate_cols:
            # Сгруппировать блоки по колонкам
            col_map: Dict[int, List[Any]] = {}
            for blk in blocks:
                try:
                    ci = int(getattr(blk, "col", 0))
                    col_map.setdefault(ci, []).append(blk)
                except Exception:
                    continue
            # Для каждой колонки создаём простую таблицу
            for ci in sorted(col_map.keys()):
                col_blocks = col_map.get(ci, [])
                col_name = (
                    column_names[ci]
                    if isinstance(column_names, list) and 0 <= ci < len(column_names)
                    else f"Колонка {ci + 1}"
                )
                elements.append(Paragraph(col_name, header_style))
                # Подготовка строк: Дата | Начало | Конец | Описание
                data: List[List[Any]] = []
                data.append(["Дата", "Начало", "Конец", "Описание"])
                for blk in col_blocks:
                    try:
                        day_index = int(getattr(blk, "day_index", 0))
                        blk_date = base_date + timedelta(days=day_index)
                        # row_start и длительность могут быть дробными, поэтому
                        # получаем их как float. Индекс строки используем как
                        # целую часть row_start, а время в минутах считаем
                        # по полному значению.
                        row_start_f = float(getattr(blk, "row_start", 0.0))
                        duration_f = float(getattr(blk, "duration_units", 1.0))
                        start_minutes = row_start_f * step_minutes
                        end_minutes = (row_start_f + duration_f) * step_minutes
                        def minutes_to_time_str(total_min: int) -> str:
                            hours = (total_min // 60) % 24
                            minutes = total_min % 60
                            return f"{hours:02d}:{minutes:02d}"
                        start_str = minutes_to_time_str(int(start_minutes))
                        end_str = minutes_to_time_str(int(end_minutes))
                        title = getattr(blk, "title", "")
                        data.append([
                            blk_date.strftime("%d.%m.%Y"),
                            start_str,
                            end_str,
                            textwrap.fill(str(title), width=30),
                        ])
                    except Exception:
                        logger.error("Ошибка обработки блока тайминга", exc_info=True)
                        continue
                # Ширины столбцов для простого списка (4 колонки)
                col_ws = [40, 30, 30, 80]  # 180 мм
                tbl = Table(data, repeatRows=1, colWidths=[w * mm for w in col_ws])
                tbl.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("FONTNAME", (0, 1), (-1, -1), "DejaVuSans"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ]))
                elements.append(tbl)
                elements.append(Spacer(1, 4 * mm))
            return elements
        # Выбор варианта отображения общего тайминга (не разбитого по колонкам).
        # Каждый вариант активируется своей галочкой независимо от других.
        # show_table  – дизайнерский вертикальный тайминг (по умолчанию).
        # graphic2    – горизонтальная шкала с колонками (второй вариант).
        # graphic3    – дерево времени с распределением колонок по обе стороны ствола.
        graphic2 = False
        graphic3 = False
        show_table_flag = show_table
        try:
            if opts:
                graphic2 = bool(opts.get("graphic2", False))
                graphic3 = bool(opts.get("graphic3", False))
                # show_table по умолчанию может быть отключён отдельной галочкой
                show_table_flag = bool(opts.get("show_table", show_table))
        except Exception:
            graphic2 = False
            graphic3 = False
            show_table_flag = show_table
        # Вариант 3: дерево времени — временная шкала в центре и колонки по обе стороны
        if graphic3:
            from reportlab.graphics.shapes import Drawing, Rect, String, Line
            from reportlab.platypus import KeepTogether, PageBreak
            # Для каждого дня рисуем вертикальную шкалу в центре и распределяем колонки по две стороны.
            try:
                day_indices = sorted({int(getattr(b, "day_index", 0)) for b in blocks})
            except Exception:
                day_indices = [0]
            for di, day_idx in enumerate(day_indices):
                # Заголовок с датой
                day_date = base_date + timedelta(days=day_idx)
                elements.append(Paragraph(day_date.strftime("%d.%m.%Y"), header_style))
                # Выбираем блоки текущего дня и сортируем по времени начала
                day_blocks: List[Any] = []
                for blk in blocks:
                    try:
                        if int(getattr(blk, "day_index", 0)) == day_idx:
                            day_blocks.append(blk)
                    except Exception:
                        continue
                # Если нет блоков, пропускаем
                if not day_blocks:
                    continue
                # Рассчитываем старт и длительность в минутах
                for blk in day_blocks:
                    try:
                        blk._start_min = float(getattr(blk, "row_start", 0.0)) * step_minutes
                        blk._end_min = blk._start_min + float(getattr(blk, "duration_units", 1.0)) * step_minutes
                    except Exception:
                        blk._start_min = 0.0
                        blk._end_min = 0.0
                # Сортируем блоки по началу
                day_blocks.sort(key=lambda b: b._start_min)
                # Группируем блоки по времени начала
                groups: Dict[float, List[Any]] = {}
                for blk in day_blocks:
                    groups.setdefault(blk._start_min, []).append(blk)
                # Определяем максимальное число уровней (полумест) для одной стороны
                max_side_levels = 1
                for group_blocks in groups.values():
                    n = len(group_blocks)
                    side_levels = (n + 1) // 2
                    if side_levels > max_side_levels:
                        max_side_levels = side_levels
                # Размеры рисунка: используем ширину 180 мм и высоту 200 мм (портрет),
                # что помещается на страницу с учётом полей (A4)
                timeline_width_mm = 180.0
                timeline_height_mm = 200.0
                trunk_x_mm = timeline_width_mm / 2.0
                branch_len_mm = 4.0  # длина ветви
                half_width_mm = timeline_width_mm / 2.0
                # Вычисляем ширину блока в зависимости от максимального количества уровней
                available_side_mm = max(1e-3, half_width_mm - branch_len_mm)
                width_box_mm = available_side_mm / max_side_levels
                # Ограничиваем максимальную ширину блока, чтобы не становилась слишком большой
                if width_box_mm > 40.0:
                    width_box_mm = 40.0
                # Создаём рисунок
                d = Drawing(timeline_width_mm * mm, timeline_height_mm * mm)
                # Ствол времени по центру
                try:
                    trunk = Line(trunk_x_mm * mm, 0, trunk_x_mm * mm, timeline_height_mm * mm)
                    trunk.strokeColor = colors.black
                    trunk.strokeWidth = 0.8
                    d.add(trunk)
                except Exception:
                    pass
                # Засечки и подписи каждые 2 часа
                for h in range(0, 25, 2):
                    y_mm = (h / 24.0) * timeline_height_mm
                    try:
                        tick = Line((trunk_x_mm - 1.0) * mm, y_mm * mm, (trunk_x_mm + 1.0) * mm, y_mm * mm)
                        tick.strokeColor = colors.grey
                        tick.strokeWidth = 0.3
                        d.add(tick)
                        lbl = String((trunk_x_mm + 2.0) * mm,
                                     (y_mm - 1.5) * mm,
                                     f"{h:02d}:00",
                                     fontSize=5,
                                     fillColor=colors.black,
                                     fontName="DejaVuSans")
                        d.add(lbl)
                    except Exception:
                        pass
                # Рисуем блоки для каждой группы начала
                for start_min, g_blocks in groups.items():
                    # y координата начала для данной группы
                    y_mm = (start_min / 1440.0) * timeline_height_mm
                    # для каждого блока определяем позицию
                    for idx, blk in enumerate(g_blocks):
                        try:
                            # Определяем сторону и уровень
                            side_left = (idx % 2 == 0)
                            level = idx // 2
                            start_min_blk = blk._start_min
                            end_min_blk = blk._end_min
                            rect_height_mm = max(2.0, (end_min_blk - start_min_blk) / 1440.0 * timeline_height_mm)
                            if side_left:
                                x_rect_mm = trunk_x_mm - branch_len_mm - (level + 1) * width_box_mm
                                # линия связи от ствола к правой стороне прямоугольника
                                try:
                                    line = Line(trunk_x_mm * mm, y_mm * mm, (x_rect_mm + width_box_mm) * mm, y_mm * mm)
                                    line.strokeColor = colors.grey
                                    line.strokeWidth = 0.25
                                    d.add(line)
                                except Exception:
                                    pass
                            else:
                                x_rect_mm = trunk_x_mm + branch_len_mm + level * width_box_mm
                                try:
                                    line = Line(trunk_x_mm * mm, y_mm * mm, x_rect_mm * mm, y_mm * mm)
                                    line.strokeColor = colors.grey
                                    line.strokeWidth = 0.25
                                    d.add(line)
                                except Exception:
                                    pass
                            # прямоугольник блока
                            rect = Rect(x_rect_mm * mm,
                                        y_mm * mm,
                                        width_box_mm * mm,
                                        rect_height_mm * mm,
                                        fillColor=colors.HexColor(getattr(blk, "color", "#6fbf73")),
                                        strokeColor=colors.black,
                                        strokeWidth=0.3)
                            d.add(rect)
                            # подпись внутри блока
                            sh = int((start_min_blk // 60) % 24)
                            sm = int(start_min_blk % 60)
                            eh = int((blk._end_min // 60) % 24)
                            em = int(blk._end_min % 60)
                            start_str = f"{sh:02d}:{sm:02d}"
                            end_str = f"{eh:02d}:{em:02d}"
                            title = str(getattr(blk, "title", ""))
                            # Имя колонки
                            try:
                                ci = int(getattr(blk, "col", 0))
                                col_name = column_names[ci] if 0 <= ci < len(column_names) else ""
                            except Exception:
                                col_name = ""
                            import textwrap as twrap
                            title_wrapped = twrap.shorten(title, width=25, placeholder="…")
                            if col_name:
                                label_text = f"{start_str}–{end_str} [{col_name}] {title_wrapped}"
                            else:
                                label_text = f"{start_str}–{end_str} {title_wrapped}"
                            txt = String((x_rect_mm + 0.2) * mm,
                                         (y_mm + rect_height_mm / 2.0 - 1.5) * mm,
                                         label_text,
                                         fontSize=5,
                                         fillColor=colors.white,
                                         fontName="DejaVuSans")
                            d.add(txt)
                        except Exception:
                            logger.error("Ошибка отрисовки блока тайминга (графическое дерево)", exc_info=True)
                            continue
                elements.append(KeepTogether(d))
                elements.append(Spacer(1, 5 * mm))
                if di < len(day_indices) - 1:
                    elements.append(PageBreak())
            return elements
        # Вариант 2: горизонтальная шкала с колонками (второй графический вариант)
        if graphic2:
            from reportlab.graphics.shapes import Drawing, Rect, String, Line
            from reportlab.platypus import KeepTogether, PageBreak
            # Каждый день отображается как набор строк: каждая строка соответствует колонке
            try:
                day_indices = sorted({int(getattr(b, "day_index", 0)) for b in blocks})
            except Exception:
                day_indices = [0]
            for di, day_idx in enumerate(day_indices):
                # Заголовок даты
                day_date = base_date + timedelta(days=day_idx)
                elements.append(Paragraph(day_date.strftime("%d.%m.%Y"), header_style))
                # Выбираем блоки на этот день
                day_blocks: List[Any] = []
                for blk in blocks:
                    try:
                        if int(getattr(blk, "day_index", 0)) == day_idx:
                            day_blocks.append(blk)
                    except Exception:
                        continue
                if not day_blocks:
                    continue
                # Вычисляем параметры
                num_cols = max(1, len(column_names))
                # Размеры: ширина рисунка 260 мм (почти вся ширина листа),
                # высота 150 мм для временной шкалы. Место под заголовки 6 мм.
                total_width_mm = 260.0
                timeline_height_mm = 150.0
                label_width_mm = 30.0
                col_width_mm = (total_width_mm - label_width_mm) / max(1, num_cols)
                header_height_mm = 6.0
                # Определяем начальную точку временной шкалы: за час до первого события
                # Находим минимальное время начала для дня (в минутах)
                min_start_min = None
                for blk in day_blocks:
                    try:
                        st = float(getattr(blk, "row_start", 0.0)) * step_minutes
                        if min_start_min is None or st < min_start_min:
                            min_start_min = st
                    except Exception:
                        continue
                if min_start_min is None:
                    min_start_min = 0.0
                # Округляем до часа вниз и отнимаем один час, чтобы добавить запас
                timeline_start_min = max(0.0, (int(min_start_min // 60) * 60) - 60.0)
                # Завершаем 24 часа после начала
                timeline_end_min = timeline_start_min + 1440.0
                range_min = timeline_end_min - timeline_start_min
                # Создаём рисунок: высота временной шкалы + место для заголовков
                d = Drawing(total_width_mm * mm, (timeline_height_mm + header_height_mm) * mm)
                # Фон белый
                try:
                    bg = Rect(0, 0, total_width_mm * mm, (timeline_height_mm) * mm, fillColor=colors.white, strokeColor=colors.white)
                    d.add(bg)
                except Exception:
                    pass
                # Горизонтальные линии и подписи времени
                for h in range(25):
                    # рассчитываем позицию y с учётом начала шкалы и обратного направления (сверху вниз)
                    ratio = (h * 60.0) / range_min
                    y_mm = timeline_height_mm - (ratio * timeline_height_mm)
                    try:
                        # линия
                        line = Line(label_width_mm * mm, y_mm * mm, total_width_mm * mm, y_mm * mm)
                        line.strokeColor = colors.grey
                        line.strokeWidth = 0.25
                        d.add(line)
                        # подпись времени слева (каждый час)
                        hour_label = int((timeline_start_min // 60 + h) % 24)
                        lbl = String((label_width_mm - 1.0) * mm,
                                     (y_mm - 1.5) * mm,
                                     f"{hour_label:02d}:00",
                                     fontSize=5,
                                     fillColor=colors.black,
                                     textAnchor="end",
                                     fontName="DejaVuSans")
                        d.add(lbl)
                    except Exception:
                        pass
                # Вертикальные линии и заголовки колонок
                for ci, col_name in enumerate(column_names):
                    x_mm = label_width_mm + ci * col_width_mm
                    try:
                        # вертикальная линия разделения (не рисуем для первой колонки)
                        if ci > 0:
                            vline = Line(x_mm * mm, 0, x_mm * mm, timeline_height_mm * mm)
                            vline.strokeColor = colors.grey
                            vline.strokeWidth = 0.25
                            d.add(vline)
                        # название колонки вверху (над временной шкалой)
                        name_lbl = String((x_mm + col_width_mm / 2.0) * mm,
                                           timeline_height_mm * mm + 1.0 * mm,
                                           str(col_name),
                                           fontSize=5,
                                           fillColor=colors.black,
                                           textAnchor="middle",
                                           fontName="DejaVuSans-Bold")
                        d.add(name_lbl)
                    except Exception:
                        pass
                # Последняя вертикальная линия справа
                try:
                    vline = Line((label_width_mm + num_cols * col_width_mm) * mm, 0, (label_width_mm + num_cols * col_width_mm) * mm, timeline_height_mm * mm)
                    vline.strokeColor = colors.grey
                    vline.strokeWidth = 0.25
                    d.add(vline)
                except Exception:
                    pass
                # 9. Рисуем блоки
                for blk in day_blocks:
                    try:
                        ci = int(getattr(blk, "col", 0))
                        start_min = float(getattr(blk, "row_start", 0.0)) * step_minutes
                        dur_min = float(getattr(blk, "duration_units", 1.0)) * step_minutes
                        end_min = start_min + dur_min
                        # Позиция блока по горизонтали
                        x_rect_mm = label_width_mm + ci * col_width_mm
                        rect_width_mm = col_width_mm
                        # Позиции по вертикали с учётом начала шкалы и обратной ориентации
                        ratio_start = (start_min - timeline_start_min) / range_min
                        ratio_end = (end_min - timeline_start_min) / range_min
                        y_top_mm = timeline_height_mm - ratio_start * timeline_height_mm
                        y_bottom_mm = timeline_height_mm - ratio_end * timeline_height_mm
                        rect_height_mm = max(1.5, y_top_mm - y_bottom_mm)
                        y_rect_mm = y_bottom_mm
                        # цвет блока
                        block_color = colors.HexColor(getattr(blk, "color", "#6fbf73"))
                        rect = Rect(x_rect_mm * mm,
                                    y_rect_mm * mm,
                                    rect_width_mm * mm,
                                    rect_height_mm * mm,
                                    fillColor=block_color,
                                    strokeColor=colors.black,
                                    strokeWidth=0.3)
                        d.add(rect)
                        # подпись внутри блока: отображаем только время и заголовок задачи,
                        # без имени колонки. Формируем время начала и конца в формате HH:MM.
                        sh = int((start_min // 60) % 24)
                        sm = int(start_min % 60)
                        eh = int((end_min // 60) % 24)
                        em = int(end_min % 60)
                        start_str = f"{sh:02d}:{sm:02d}"
                        end_str = f"{eh:02d}:{em:02d}"
                        # Заголовок блока может быть длинным, сокращаем его до 30 символов
                        title = str(getattr(blk, "title", ""))
                        import textwrap as twrap
                        title_wrapped = twrap.shorten(title, width=30, placeholder="…")
                        label_text = f"{start_str}–{end_str} {title_wrapped}"
                        txt = String(
                            (x_rect_mm + 0.2) * mm,
                            (y_rect_mm + rect_height_mm / 2.0 - 1.5) * mm,
                            label_text,
                            fontSize=5,
                            fillColor=colors.white,
                            fontName="DejaVuSans",
                        )
                        d.add(txt)
                    except Exception:
                        logger.error("Ошибка отрисовки блока тайминга (графический вариант 2)", exc_info=True)
                        continue
                elements.append(KeepTogether(d))
                elements.append(Spacer(1, 5 * mm))
                if di < len(day_indices) - 1:
                    elements.append(PageBreak())
            return elements
        # Вариант 1 (по умолчанию): дизайнерский вертикальный тайминг с чередованием сторон
        if show_table_flag:
            from reportlab.graphics.shapes import Drawing, Rect, String, Line
            from reportlab.platypus import KeepTogether, PageBreak
            # Каждый день представлен вертикальной шкалой по центру, от которой отходят ветви
            try:
                day_indices = sorted({int(getattr(b, "day_index", 0)) for b in blocks})
            except Exception:
                day_indices = [0]
            for di, day_idx in enumerate(day_indices):
                # Заголовок даты
                day_date = base_date + timedelta(days=day_idx)
                elements.append(Paragraph(day_date.strftime("%d.%m.%Y"), header_style))
                # Выбираем блоки этого дня и сортируем по времени начала
                day_blocks: List[Any] = []
                for blk in blocks:
                    try:
                        if int(getattr(blk, "day_index", 0)) == day_idx:
                            day_blocks.append(blk)
                    except Exception:
                        continue
                day_blocks.sort(key=lambda b: float(getattr(b, "row_start", 0.0)) * step_minutes)
                # Настройки размера: используем всю ширину (180 мм) и высоту не более 150 мм, чтобы
                # рисунок гарантированно помещался на страницу в альбомной ориентации (A4 210×297 мм с полями).
                timeline_width_mm = 180.0
                timeline_height_mm = 150.0
                trunk_x_mm = timeline_width_mm / 2.0
                box_width_mm = 90.0  # увеличенная ширина блока
                branch_len_mm = 8.0  # длина ветви
                box_height_mm = 20.0  # увеличенная высота блока
                d = Drawing(timeline_width_mm * mm, (timeline_height_mm + 10.0) * mm)
                # Ствол времени
                try:
                    trunk = Line(trunk_x_mm * mm, 0, trunk_x_mm * mm, timeline_height_mm * mm)
                    trunk.strokeColor = colors.black
                    trunk.strokeWidth = 0.8
                    d.add(trunk)
                except Exception:
                    pass
                # Засечки и подписи каждые 2 часа
                for h in range(0, 25, 2):
                    y_mm = (h / 24.0) * timeline_height_mm
                    try:
                        tick = Line((trunk_x_mm - 2.0) * mm, y_mm * mm, (trunk_x_mm + 2.0) * mm, y_mm * mm)
                        tick.strokeColor = colors.grey
                        tick.strokeWidth = 0.3
                        d.add(tick)
                        lbl = String((trunk_x_mm + 3.0) * mm,
                                     (y_mm - 1.5) * mm,
                                     f"{h:02d}:00",
                                     fontSize=6,
                                     fillColor=colors.black,
                                     fontName="DejaVuSans")
                        d.add(lbl)
                    except Exception:
                        pass
                # Размещаем блоки чередуя стороны
                for idx, blk in enumerate(day_blocks):
                    try:
                        start_min = float(getattr(blk, "row_start", 0.0)) * step_minutes
                        dur_min = float(getattr(blk, "duration_units", 1.0)) * step_minutes
                        end_min = start_min + dur_min
                    except Exception:
                        continue
                    # Вертикальная позиция — пропорционально началу блока
                    y_mm = (start_min / 1440.0) * timeline_height_mm
                    side_left = (idx % 2 == 0)
                    if side_left:
                        x_rect_mm = trunk_x_mm - branch_len_mm - box_width_mm
                        try:
                            conn = Line(trunk_x_mm * mm, y_mm * mm, (x_rect_mm + box_width_mm) * mm, y_mm * mm)
                            conn.strokeColor = colors.grey
                            conn.strokeWidth = 0.25
                            d.add(conn)
                        except Exception:
                            pass
                    else:
                        x_rect_mm = trunk_x_mm + branch_len_mm
                        try:
                            conn = Line(trunk_x_mm * mm, y_mm * mm, x_rect_mm * mm, y_mm * mm)
                            conn.strokeColor = colors.grey
                            conn.strokeWidth = 0.25
                            d.add(conn)
                        except Exception:
                            pass
                    # Прямоугольник блока
                    try:
                        rect = Rect(x_rect_mm * mm,
                                    (y_mm - box_height_mm / 2.0) * mm,
                                    box_width_mm * mm,
                                    box_height_mm * mm,
                                    fillColor=colors.HexColor(getattr(blk, "color", "#6fbf73")),
                                    strokeColor=colors.black,
                                    strokeWidth=0.3)
                        d.add(rect)
                    except Exception:
                        pass
                    # Подпись внутри блока
                    try:
                        sh = int((start_min // 60) % 24)
                        sm = int(start_min % 60)
                        eh = int((end_min // 60) % 24)
                        em = int(end_min % 60)
                        start_str = f"{sh:02d}:{sm:02d}"
                        end_str = f"{eh:02d}:{em:02d}"
                        title = str(getattr(blk, "title", ""))
                        # Имя колонки
                        try:
                            ci = int(getattr(blk, "col", 0))
                            col_name = column_names[ci] if 0 <= ci < len(column_names) else ""
                        except Exception:
                            col_name = ""
                        import textwrap as twrap
                        title_wrapped = twrap.shorten(title, width=40, placeholder="…")
                        if col_name:
                            label_text = f"{start_str}–{end_str} [{col_name}] {title_wrapped}"
                        else:
                            label_text = f"{start_str}–{end_str} {title_wrapped}"
                        txt = String((x_rect_mm + 0.5) * mm,
                                     (y_mm - box_height_mm / 2.0 + box_height_mm / 2.0 - 1.5) * mm,
                                     label_text,
                                     fontSize=6,
                                     fillColor=colors.white,
                                     fontName="DejaVuSans")
                        d.add(txt)
                    except Exception:
                        logger.error("Ошибка отрисовки блока тайминга (вертикальный)", exc_info=True)
                        continue
                elements.append(KeepTogether(d))
                elements.append(Spacer(1, 5 * mm))
                if di < len(day_indices) - 1:
                    elements.append(PageBreak())
            return elements
        # Если дошли сюда, то ни один вариант не выбран, возвращаем пустой список
        return elements
    except Exception:
        logger.error("Ошибка создания отчёта тайминга", exc_info=True)
        return elements
