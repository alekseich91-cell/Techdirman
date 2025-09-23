"""
Назначение:
    Этот модуль реализует вкладку «Информация» страницы проекта.
    Она содержит форму для ввода базовых сведений о проекте (название,
    даты, заказчик, монтаж/демонтаж и т. п.), место для обложки и
    блок «Сводные показатели».

Принцип работы:
    • При создании вкладки функция `build_info_tab(page, tab)` строит
      интерфейс, привязывая виджеты к атрибутам объекта страницы `page`.
    • Данные формы сохраняются и загружаются в JSON‑файл с помощью
      функций `save_info_json(page)` и `load_info_json(page)`.
    • При загрузке вкладки дополнительно выполняется расчёт
      сводных финансовых показателей (суммы, комиссии, скидки) на
      основе текущих позиций проекта и настроекц вкладки «Бухгалтерия».
      Этот расчёт выполняет функция `update_financial_summary(page)`,
      которая обращается к провайдерам данных из модуля
      :mod:`finance_tab` и обновляет соответствующие метки на вкладке.

Стиль:
    • Код разделён на пронумерованные секции с заголовками.
    • Внутри секций присутствуют краткие комментарии, поясняющие ключевые
      действия.
"""

# 1. Импорт библиотек и модулей
from pathlib import Path
import json
from typing import Any, Dict, List, Optional
import logging
import os

from PySide6 import QtWidgets, QtCore, QtGui

# Импорт внутренних модулей
from .common import DATA_DIR

# 3.0 Дополнительная директория данных
# Помимо локальной папки ``DATA_DIR`` внутри приложения (TechDirRentMan/data)
# существует также глобальная папка данных на уровень выше (A37/data). Именно
# туда сохраняются файлы проекта, когда используется база данных. Использование
# единой директории для информации о проектах устраняет путаницу между
# внутренними шаблонами и пользовательскими данными. Мы рассчитываем путь
# к этой внешней директории, поднявшись на три уровня относительно
# текущего файла (ui/info_tab.py → ui → TechDirRentMan → A37) и добавив папку ``data``.
ROOT_DATA_DIR = Path(__file__).resolve().parents[3] / "data"
from .widgets import ImageDropLabel
from .common import ASSETS_DIR
import shutil

# Импортируем типы и функции из finance_tab для расчёта сводных
try:
    # Lazy import to avoid circular dependencies at import time
    from .finance_tab import (
        DBDataProvider, FileDataProvider,
        compute_client_flow, compute_internal_discount,
        aggregate_by_vendor, VendorSettings, round2,
    )  # type: ignore
except Exception:
    # Если импорт не удался, объявим заглушки; это покрывает
    # ситуацию, когда модуль finance_tab отсутствует или его импорт
    # приводит к ошибке. Логи помогут обнаружить проблему.
    DBDataProvider = None  # type: ignore
    FileDataProvider = None  # type: ignore
    compute_client_flow = None  # type: ignore
    compute_internal_discount = None  # type: ignore
    aggregate_by_vendor = None  # type: ignore
    VendorSettings = None  # type: ignore
    round2 = lambda x: x  # type: ignore


# 0. Настройка логирования
# Создаём директорию для логов и настраиваем логгер для вкладки «Информация»
LOG_DIR = os.path.join(os.path.dirname(__file__), "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "info_tab.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("info_tab")


# 2. Построение вкладки «Информация»
def build_info_tab(page: Any, tab: QtWidgets.QWidget) -> None:
    """Создаёт интерфейс вкладки «Информация».

    Параметры:
        page: экземпляр ProjectPage, в котором будут размещены поля;
        tab: виджет вкладки, на котором разместится интерфейс.
    """
    # Контейнер с горизонтальной компоновкой
    h = QtWidgets.QHBoxLayout(tab)

    # Левая колонка — форма проекта
    left_wrap = QtWidgets.QWidget()
    form = QtWidgets.QFormLayout(left_wrap)
    form.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)

    def ex(w: QtWidgets.QWidget) -> QtWidgets.QWidget:
        """Устанавливает политику расширения по горизонтали."""
        w.setSizePolicy(QtWidgets.QSizePolicy.Expanding, w.sizePolicy().verticalPolicy())
        return w

    # Поля формы: создаём и привязываем к атрибутам страницы
    page.ed_title = ex(QtWidgets.QLineEdit())
    page.ed_date = ex(QtWidgets.QLineEdit())
    page.ed_customer = ex(QtWidgets.QLineEdit())
    page.ed_mount_datetime = ex(QtWidgets.QLineEdit())
    page.ed_site_ready = ex(QtWidgets.QLineEdit())
    page.ed_address = ex(QtWidgets.QLineEdit())
    page.ed_site_ready_dup = ex(QtWidgets.QLineEdit())
    page.ed_dismount_time = ex(QtWidgets.QLineEdit())
    page.ed_floor_elevator = ex(QtWidgets.QLineEdit())
    page.ed_power_capacity = ex(QtWidgets.QLineEdit())
    page.ed_storage_possible = ex(QtWidgets.QLineEdit())

    # Многострочное поле комментариев
    page.ed_comments = QtWidgets.QTextEdit()
    page.ed_comments.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
    page.ed_comments.setMinimumHeight(240)

    # Кнопка сохранения и сигнал
    page.btn_save_info = QtWidgets.QPushButton("Сохранить изменения")
    # По нажатию сохраняем JSON и заново вычисляем сводные показатели
    def on_save_and_update() -> None:
        try:
            save_info_json(page)
            # Пересчитываем сводные показатели после сохранения
            update_financial_summary(page)
        except Exception:
            logger.error("Ошибка при сохранении информации и обновлении сводных показателей", exc_info=True)
    page.btn_save_info.clicked.connect(on_save_and_update)

    # Добавляем строки формы
    for label, w in [
        ("Название:", page.ed_title),
        ("Дата:", page.ed_date),
        ("Заказчик:", page.ed_customer),
        ("Дата и время заезда на монтаж:", page.ed_mount_datetime),
        ("Готовность площадки:", page.ed_site_ready),
        ("Адрес:", page.ed_address),
        ("Готовность площадки (повтор):", page.ed_site_ready_dup),
        ("Время демонтажа:", page.ed_dismount_time),
        ("Этаж и наличие лифта:", page.ed_floor_elevator),
        ("Количество электричества на площадке:", page.ed_power_capacity),
        ("Возможность складирования кофров:", page.ed_storage_possible),
        ("Комментарии:", page.ed_comments),
    ]:
        form.addRow(label, w)
    form.addRow(page.btn_save_info)

    # Скролл слева: позволяет прокручивать форму
    left_scroll = QtWidgets.QScrollArea()
    left_scroll.setWidgetResizable(True)
    left_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
    left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
    left_scroll.setWidget(left_wrap)

    # Правая колонка — обложка и сводные показатели
    right_root = QtWidgets.QWidget()
    right_v = QtWidgets.QVBoxLayout(right_root)

    # Виджет для перетаскивания изображения обложки
    page.cover_label = ImageDropLabel()
    right_v.addWidget(page.cover_label)

    # Группа с показателями
    group = QtWidgets.QGroupBox("Сводные показатели")
    g = QtWidgets.QFormLayout(group)
    g.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)

    # Метки с итогами; при загрузке/сохранении обновляются
    page.lbl_cash_total = QtWidgets.QLabel("0")
    page.lbl_tax_total = QtWidgets.QLabel("0")
    page.lbl_commission_each = QtWidgets.QLabel("—")
    page.lbl_commission_sum = QtWidgets.QLabel("0")
    page.lbl_discount_each = QtWidgets.QLabel("—")
    page.lbl_discount_sum = QtWidgets.QLabel("0")
    page.lbl_power_sum = QtWidgets.QLabel("0")

    # Добавляем строки в группу
    for label, w in [
        ("Общая стоимость (нал):", page.lbl_cash_total),
        ("Общая стоимость с налогом:", page.lbl_tax_total),
        ("Комиссия по каждому подрядчику:", page.lbl_commission_each),
        ("Комиссия суммарно:", page.lbl_commission_sum),
        ("Скидка по подрядчикам:", page.lbl_discount_each),
        ("Скидка суммарно:", page.lbl_discount_sum),
        ("Суммарное потребление:", page.lbl_power_sum),
    ]:
        g.addRow(label, w)

    right_v.addWidget(group)
    # 2.a Добавляем секцию для файлов проекта (исходники смет и сопутствующие материалы)
    try:
        attachments_group = QtWidgets.QGroupBox("Файлы проекта")
        attach_layout = QtWidgets.QVBoxLayout(attachments_group)

        # Внутренний класс для перетаскивания файлов
        class FileDropFrame(QtWidgets.QFrame):
            """
            Простое поле с поддержкой drag‑and‑drop для загрузки файлов.
            При приёме файла копирует его в подпапку проекта ``subfolder``.
            Список добавленных файлов выводится в метке внутри виджета.
            """
            def __init__(self, title: str, subfolder: str, parent: Optional[QtWidgets.QWidget] = None) -> None:
                super().__init__(parent)
                self.setFrameShape(QtWidgets.QFrame.StyledPanel)
                self.setFrameShadow(QtWidgets.QFrame.Sunken)
                self.setAcceptDrops(True)
                self.setMinimumHeight(60)
                self.title = title
                self.subfolder = subfolder
                self.label = QtWidgets.QLabel(title, self)
                self.label.setAlignment(QtCore.Qt.AlignCenter)
                layout = QtWidgets.QVBoxLayout(self)
                layout.addWidget(self.label)
                # Храним список имён файлов для отображения
                self.added_files: List[str] = []

            def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:  # type: ignore
                if event.mimeData().hasUrls():
                    event.acceptProposedAction()
                else:
                    event.ignore()

            def dropEvent(self, event: QtGui.QDropEvent) -> None:  # type: ignore
                try:
                    urls = event.mimeData().urls()
                    if not urls:
                        return
                    proj_id = getattr(page, "project_id", None)
                    if not proj_id:
                        return
                    dest_dir = ASSETS_DIR / f"project_{proj_id}" / self.subfolder
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    for url in urls:
                        src_path = url.toLocalFile()
                        if not src_path:
                            continue
                        if not os.path.isfile(src_path):
                            continue
                        base = os.path.basename(src_path)
                        dest_path = dest_dir / base
                        # При совпадении имён добавляем суффикс
                        counter = 1
                        name, ext = os.path.splitext(base)
                        while dest_path.exists():
                            dest_path = dest_dir / f"{name}_{counter}{ext}"
                            counter += 1
                        shutil.copy2(src_path, dest_path)
                        self.added_files.append(dest_path.name)
                        logger.info("Добавлен файл '%s' в папку %s", dest_path.name, dest_dir)
                    # Обновляем текст метки: показываем количество файлов
                    if self.added_files:
                        self.label.setText(
                            f"Добавлено файлов: {len(self.added_files)}\n" + "\n".join(self.added_files[-3:])
                        )
                    else:
                        self.label.setText(self.title)
                except Exception:
                    logger.error("Ошибка обработки перетаскивания файлов", exc_info=True)

        # Поле для исходников смет
        src_label = QtWidgets.QLabel("Исходники смет:")
        src_drop = FileDropFrame("Перетащите файлы исходников сюда", "info_sources")
        src_open_btn = QtWidgets.QPushButton("Открыть папку")
        def open_src_folder() -> None:
            try:
                proj_id = getattr(page, "project_id", None)
                if not proj_id:
                    return
                path = ASSETS_DIR / f"project_{proj_id}" / "info_sources"
                path.mkdir(parents=True, exist_ok=True)
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(path)))
            except Exception:
                logger.error("Не удалось открыть папку исходников", exc_info=True)
        src_open_btn.clicked.connect(open_src_folder)
        # Компоновка для исходников
        src_layout = QtWidgets.QVBoxLayout()
        src_layout.addWidget(src_label)
        src_layout.addWidget(src_drop)
        src_layout.addWidget(src_open_btn)

        # Поле для сопутствующих материалов
        mat_label = QtWidgets.QLabel("Сопутствующие материалы:")
        mat_drop = FileDropFrame("Перетащите сопутствующие файлы сюда", "info_materials")
        mat_open_btn = QtWidgets.QPushButton("Открыть папку")
        def open_mat_folder() -> None:
            try:
                proj_id = getattr(page, "project_id", None)
                if not proj_id:
                    return
                path = ASSETS_DIR / f"project_{proj_id}" / "info_materials"
                path.mkdir(parents=True, exist_ok=True)
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(path)))
            except Exception:
                logger.error("Не удалось открыть папку сопутствующих материалов", exc_info=True)
        mat_open_btn.clicked.connect(open_mat_folder)
        # Компоновка для материалов
        mat_layout = QtWidgets.QVBoxLayout()
        mat_layout.addWidget(mat_label)
        mat_layout.addWidget(mat_drop)
        mat_layout.addWidget(mat_open_btn)

        # Добавляем обе секции в общий контейнер
        attach_layout.addLayout(src_layout)
        attach_layout.addLayout(mat_layout)
        right_v.addWidget(attachments_group)
    except Exception:
        logger.error("Не удалось создать виджеты для файлов проекта", exc_info=True)

    right_v.addStretch(1)

    # При создании вкладки сразу обновляем сводные финансовые показатели.
    try:
        update_financial_summary(page)
    except Exception:
        # Если расчёт не удался, записываем в лог, но не мешаем загрузке вкладки
        logger.error("Не удалось вычислить сводные показатели при инициализации вкладки", exc_info=True)

    # Скролл справа
    right_scroll = QtWidgets.QScrollArea()
    right_scroll.setWidgetResizable(True)
    right_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
    right_scroll.setWidget(right_root)

    # Компоновка двух колонок
    h.addWidget(left_scroll, 1)
    h.addWidget(right_scroll, 1)


# 3. Пути к JSON
def info_json_path(page: Any) -> Path:
    """
    Возвращает путь к файлу с информацией проекта.

    **Логика выбора директории**
        • Если проект уже имеет идентификатор и связан с базой данных
          (атрибут ``page.db`` не None), то файл информации хранится в
          «глобальной» папке данных (``ROOT_DATA_DIR``). В этой папке лежит
          SQLite‑база и прочие пользовательские файлы проекта. Такой подход
          обеспечивает, что данные, введённые на вкладке «Информация»,
          сохраняются рядом с данными БД и доступны при последующем
          открытии проекта.
        • В противном случае (например, для временных проектов без ID или
          при работе в файловом режиме) используется локальная папка
          ``DATA_DIR`` внутри ``TechDirRentMan/data``. Здесь хранятся
          шаблоны по умолчанию (project_default_*.json) и данные для
          проектов, созданных без подключения к БД.

    :param page: объект страницы проекта, содержащий атрибуты
                 ``project_id`` и ``db``.
    :return: Path к файлу вида ``project_<id>_info.json`` в соответствующей
             директории данных.
    """
    pid = getattr(page, "project_id", None)
    # Подготовим имя файла: «default» для проектов без идентификатора
    pid_str = "default" if pid is None else str(pid)
    # Определяем, следует ли использовать внешнюю папку данных
    use_root = False
    try:
        # Сохраняем в ROOT_DATA_DIR только для проектов, привязанных к БД.
        # Это условие исключает проекты без ID (draft) и файловые режимы.
        if pid is not None and getattr(page, "db", None) is not None:
            # При наличии глобальной директории и проекта в БД — используем её
            if ROOT_DATA_DIR.is_dir():
                use_root = True
    except Exception:
        # В случае ошибки просто оставляем use_root = False
        use_root = False
    base_dir = ROOT_DATA_DIR if use_root else DATA_DIR
    file_name = f"project_{pid_str}_info.json"
    # Убедимся, что целевая директория существует
    try:
        base_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Если не удалось создать папку, будем всё равно возвращать путь,
        # запись в неё может завершиться ошибкой и будет залогирована в save_info_json
        pass
    # Основной путь для файла информации
    info_path = base_dir / file_name
    # Если мы сохраняем в ROOT_DATA_DIR, но файл там не найден, попробуем найти
    # его в локальной папке DATA_DIR. Это обеспечивает совместимость со старой
    # схемой хранения данных и предотвращает потерю сведений при смене директории.
    if use_root and not info_path.exists():
        fallback = DATA_DIR / file_name
        if fallback.exists():
            return fallback
    return info_path


# 4. Загрузка информации из JSON
def load_info_json(page: Any) -> None:
    """
    Заполняет поля формы значениями из JSON‑файла.

    4.1. Если идентификатор проекта отсутствует (None), используется
         файл по умолчанию `project_default_info.json`.
    4.2. Считывает JSON и заполняет все поля формы. При ошибках
         чтения/разбора JSON данные остаются пустыми, а исключение
         записывается в лог.
    4.3. После загрузки инициируется пересчёт сводных показателей.
    """
    p = info_json_path(page)
    data: Dict[str, Any] = {}
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            logger.info("Загружены данные вкладки 'Информация' из %s", p)
        except Exception:
            logger.error("Ошибка чтения JSON информации", exc_info=True)
            data = {}
    else:
        logger.info("Файл информации %s не найден. Используются пустые значения", p)
    # 4.2 Заполняем поля формы данными или пустыми строками
    page.ed_title.setText(data.get("title", ""))
    page.ed_date.setText(data.get("date", ""))
    page.ed_customer.setText(data.get("customer", ""))
    page.ed_mount_datetime.setText(data.get("mount_dt", ""))
    page.ed_site_ready.setText(data.get("site_ready", ""))
    page.ed_address.setText(data.get("address", ""))
    page.ed_site_ready_dup.setText(data.get("site_ready_dup", ""))
    page.ed_dismount_time.setText(data.get("dismount_time", ""))
    page.ed_floor_elevator.setText(data.get("floor_elevator", ""))
    page.ed_power_capacity.setText(data.get("power_capacity", ""))
    page.ed_storage_possible.setText(data.get("storage_possible", ""))
    page.ed_comments.setPlainText(data.get("comments", ""))
    # Обложка: загружаем картинку, если путь сохранён
    cover = data.get("cover_path")
    if cover and Path(cover).exists():
        page.cover_label._stored_path = cover
        try:
            page.cover_label._load_pixmap(Path(cover))
        except Exception:
            logger.error("Не удалось загрузить обложку из %s", cover, exc_info=True)
    else:
        page.cover_label._stored_path = None
        page.cover_label.setText("Перетащите сюда картинку (PNG/JPG)")
    # Сводные показатели
    page.lbl_cash_total.setText(data.get("cash_total", "0"))
    page.lbl_tax_total.setText(data.get("tax_total", "0"))
    page.lbl_commission_each.setText(data.get("commission_each", "—"))
    page.lbl_commission_sum.setText(data.get("commission_sum", "0"))
    page.lbl_discount_each.setText(data.get("discount_each", "—"))
    page.lbl_discount_sum.setText(data.get("discount_sum", "0"))
    page.lbl_power_sum.setText(data.get("power_sum", "0"))
    # 4.3 После загрузки данных пересчитываем сводные показатели
    try:
        update_financial_summary(page)
    except Exception:
        logger.error("Не удалось обновить сводные показатели после загрузки данных", exc_info=True)


# 5. Сохранение информации в JSON
def save_info_json(page: Any) -> None:
    """
    Сохраняет значения из формы в JSON‑файл и выводит сообщение.

    5.1. Сохраняет данные формы в файл, определяемый функцией
         `info_json_path`. Даже если идентификатор проекта не задан
         (`project_id` равен None), данные будут сохранены в файл
         `project_default_info.json`.
    5.2. При возникновении ошибок записи данные логируются, и
         пользователю выводится сообщение об ошибке. В противном
         случае показывается уведомление об успешном сохранении.
    """
    # Сбор значений формы в словарь
    data: Dict[str, Any] = {
        "title": page.ed_title.text().strip(),
        "date": page.ed_date.text().strip(),
        "customer": page.ed_customer.text().strip(),
        "mount_dt": page.ed_mount_datetime.text().strip(),
        "site_ready": page.ed_site_ready.text().strip(),
        "address": page.ed_address.text().strip(),
        "site_ready_dup": page.ed_site_ready_dup.text().strip(),
        "dismount_time": page.ed_dismount_time.text().strip(),
        "floor_elevator": page.ed_floor_elevator.text().strip(),
        "power_capacity": page.ed_power_capacity.text().strip(),
        "storage_possible": page.ed_storage_possible.text().strip(),
        "comments": page.ed_comments.toPlainText().strip(),
        "cover_path": getattr(page.cover_label, "_stored_path", None),
        "cash_total": page.lbl_cash_total.text(),
        "tax_total": page.lbl_tax_total.text(),
        "commission_each": page.lbl_commission_each.text(),
        "commission_sum": page.lbl_commission_sum.text(),
        "discount_each": page.lbl_discount_each.text(),
        "discount_sum": page.lbl_discount_sum.text(),
        "power_sum": page.lbl_power_sum.text(),
    }
    # Путь для сохранения информации
    p = info_json_path(page)
    try:
        # Записываем данные в файл
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        # Логируем успешное сохранение и выводим сообщение пользователю
        if hasattr(page, "_log"):
            page._log("Сохранены данные вкладки «Информация»")
        logger.info("Данные вкладки 'Информация' сохранены: %s", p)
        QtWidgets.QMessageBox.information(page, "Сохранено", "Изменения сохранены.")
    except Exception:
        logger.error("Ошибка сохранения JSON информации", exc_info=True)
        QtWidgets.QMessageBox.critical(page, "Ошибка", "Не удалось сохранить данные. См. логи.")


# 6. Обновление сводных финансовых показателей
def update_financial_summary(page: Any) -> None:
    """Пересчитывает и заполняет сводные финансовые показатели на вкладке.

    Эта функция использует провайдеры данных из :mod:`finance_tab` для
    чтения позиций проекта и настроек подрядчиков. На основании этих
    данных вычисляются суммарные стоимости, налог, скидки и комиссии.
    Результаты записываются в соответствующие метки на странице.
    """
    # Убедимся, что необходимые функции доступны
    if any(x is None for x in (aggregate_by_vendor, compute_client_flow, VendorSettings)):
        logger.warning("Модуль finance_tab недоступен, сводные показатели не будут обновлены")
        return
    try:
        # Определяем подходящий провайдер: из базы или файловый
        provider = None
        proj_id = getattr(page, "project_id", None)
        db = getattr(page, "db", None)
        if db is not None and proj_id is not None and DBDataProvider is not None:
            provider = DBDataProvider(page)
        else:
            # Файловый режим: ищем корень проекта относительно этого файла
            project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
            # Если в page.project_id ничего нет, используем "default"
            pid = str(proj_id) if proj_id is not None else "default"
            if FileDataProvider is not None:
                provider = FileDataProvider(project_root=project_root, project_id=pid)
        if provider is None:
            logger.warning("Не удалось определить провайдера данных для сводных показателей")
            return
        # Загружаем данные
        items = provider.load_items() or []
        vendors, profits, expenses = provider.load_finance()
        # Собираем уникальных подрядчиков из позиций
        unique_vendors = sorted({it.vendor or "(без подрядчика)" for it in items})
        # Заполняем настройки подрядчиков, включая отсутствующих
        vendor_settings = {v: vendors.get(v, VendorSettings()) for v in unique_vendors}
        # Аггрегируем суммы по подрядчикам (по классу equipment и прочим)
        # Используем коэффициенты подрядчиков только если они включены (coeff_enabled).
        preview_coeffs: Dict[str, float] = {}
        for v in unique_vendors:
            s = vendor_settings.get(v)
            try:
                if s and getattr(s, "coeff_enabled", False):
                    preview_coeffs[v] = float(getattr(s, "coeff", 1.0))
            except Exception:
                # В случае ошибок просто пропускаем коэффициент
                continue
        agg = aggregate_by_vendor(items, preview_coeffs)
        # Подготовка итоговых показателей
        total_subtotal = 0.0  # сумма до налога
        total_with_tax = 0.0  # сумма после налога
        total_comm_amount = 0.0
        total_disc_amount = 0.0
        commission_each_parts: list[str] = []
        discount_each_parts: list[str] = []
        for v in unique_vendors:
            data = agg.get(v, {"equip_sum": 0.0, "other_sum": 0.0})
            s = vendor_settings.get(v, VendorSettings())
            equip_sum = data.get("equip_sum", 0.0)
            other_sum = data.get("other_sum", 0.0)
            # Расчёт по клиентскому потоку
            disc_amt, comm_amt, tax_amt, subtotal, total_taxed = compute_client_flow(
                equip_sum, other_sum, s.discount_pct, s.commission_pct, s.tax_pct
            )
            total_subtotal += subtotal
            total_with_tax += total_taxed
            total_comm_amount += comm_amt
            total_disc_amount += disc_amt
            # Формируем текстовые элементы «по каждому»
            # Пропускаем, если скидка/комиссия равна нулю
            if s.commission_pct:
                commission_each_parts.append(
                    f"{v}: {round2(s.commission_pct):,.2f}% ({round2(comm_amt):,.2f} ₽)".replace(",", " ")
                )
            if s.discount_pct:
                discount_each_parts.append(
                    f"{v}: {round2(s.discount_pct):,.2f}% ({round2(disc_amt):,.2f} ₽)".replace(",", " ")
                )
        # Записываем итоги в метки на странице
        # Общая стоимость без налога
        page.lbl_cash_total.setText(f"{round2(total_subtotal):,.2f} ₽".replace(",", " "))
        # Общая стоимость с налогом
        page.lbl_tax_total.setText(f"{round2(total_with_tax):,.2f} ₽".replace(",", " "))
        # Комиссии по каждому
        page.lbl_commission_each.setText(
            "; ".join(commission_each_parts) if commission_each_parts else "—"
        )
        # Суммарная комиссия
        page.lbl_commission_sum.setText(f"{round2(total_comm_amount):,.2f} ₽".replace(",", " "))
        # Скидки по каждому
        page.lbl_discount_each.setText(
            "; ".join(discount_each_parts) if discount_each_parts else "—"
        )
        # Суммарная скидка
        page.lbl_discount_sum.setText(f"{round2(total_disc_amount):,.2f} ₽".replace(",", " "))
        # Суммарная мощность — по умолчанию 0, т.к. Item не содержит power.
        # Можно вычислять через БД, если доступна соответствующая колонка.
        try:
            total_power_w = 0.0
            # Если элемент Item имеет атрибут power_watts, суммируем мощность
            if items and hasattr(items[0], "power_watts"):
                for it in items:
                    # type: ignore[attr-defined]
                    total_power_w += float(getattr(it, "power_watts", 0.0)) * float(getattr(it, "qty", 1.0))
            # Показываем киловатты с двумя знаками
            page.lbl_power_sum.setText(f"{round2(total_power_w / 1000.0):,.2f} кВт".replace(",", " "))
        except Exception:
            logger.warning("Не удалось вычислить суммарную мощность", exc_info=True)
    except Exception:
        logger.error("Ошибка при вычислении сводных финансовых показателей", exc_info=True)
