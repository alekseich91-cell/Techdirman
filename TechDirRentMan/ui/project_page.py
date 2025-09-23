"""
Назначение:
    Этот модуль реализует страницу проекта (ProjectPage) приложения TechDirRentMan.
    Вкладка «Сводная смета» позволяет:
      • отображать и фильтровать позиции по зонам;
      • редактировать qty/coeff/price с пересчётом суммы;
      • добавлять позицию вручную (с полем «Цена/шт») —
        ВАЖНО: теперь ручное добавление также записывает позицию в глобальную Базу данных (каталог);
      • переносить позиции в другую зону;
      • удалять выделенные строки;
      • отменять последнее действие (ручное добавление, удаление, перенос, редактирование).

Принцип работы:
    • Данные проекта пишутся через объект DB (таблица project_items).
    • Для каталога используется DB.catalog_add_or_ignore(...) — пополняет глобальную БД без дублей.
    • Для UNDO используется self._last_action — хранит снимок/метаинформацию.
    • Все значимые действия пишутся в лог через self._log().

Стиль:
    • Код разбит на пронумерованные секции с короткими заголовками.
    • Внутри секций — однострочные комментарии к ключевым операциям.
"""

# 1. Импорт библиотек
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
from datetime import datetime
import json, csv
import traceback  # Для вывода стектрейса при логировании ошибок

from PySide6 import QtWidgets, QtGui, QtCore

# 2. Импорт внутренних модулей
from db import DB
from .common import (
    CLASS_RU2EN, CLASS_EN2RU, WRAP_THRESHOLD, fmt_num, to_float,
    apply_auto_col_resize, setup_auto_col_resize, setup_priority_name, DATA_DIR
)
from .delegates import WrapTextDelegate
from .widgets import ImageDropLabel, SmartDoubleSpinBox, FileDropLabel

# 2.1 Диалоги (MoveDialog обязателен, PowerMismatchDialog опционален)
try:
    from .dialogs import MoveDialog, PowerMismatchDialog
except Exception:  # noqa: E722
    from .dialogs import MoveDialog
    PowerMismatchDialog = None  # type: ignore

# 2.2 Импорт внешних модулей вкладок
#   Эти функции реализуют основную логику вкладок и используются вместо
#   длинного кода в этом файле. См. summary_tab.py и import_tab.py
from .info_tab import (
    build_info_tab as _info_build_tab,
    info_json_path as info_json_path_ext,
    load_info_json as load_info_json_ext,
    save_info_json as save_info_json_ext,
)
from .summary_tab import (
    build_summary_tab as _summary_build_tab,
    init_zone_tabs as init_zone_tabs_ext,
    fill_manual_zone_combo as fill_manual_zone_combo_ext,
    reload_zone_tabs as reload_zone_tabs_ext,
    create_zone as create_zone_ext,
    move_selected_to_zone as move_selected_to_zone_ext,
    add_manual_item as add_manual_item_ext,
    on_summary_item_changed as on_summary_item_changed_ext,
    delete_selected as delete_selected_ext,
    undo_last_summary as undo_last_summary_ext,
)
from .import_tab import (
    build_import_tab as _import_build_tab,
    choose_file as choose_file_ext,
    on_file_dropped as on_file_dropped_ext,
    read_source_file as read_source_file_ext,
    build_mapping_bar as build_mapping_bar_ext,
    fill_src_table as fill_src_table_ext,
    current_mapping as current_mapping_ext,
    rebuild_result as rebuild_result_ext,
    update_import_button_state as update_import_button_state_ext,
    refresh_vendor_dept_zone_lists as refresh_vendor_dept_zone_lists_ext,
    apply_import as apply_import_ext,
    undo_last_import as undo_last_import_ext,
)

# 2.3 Дополнительные вкладки: создание сметы, бухгалтерия, экспорт
# Эти модули содержат базовую логику для соответствующих вкладок. В
# настоящий момент они реализованы как заглушки, но позволяют
# организовать структуру и подключить последующую функциональность.
# Вкладка «Создание сметы» устарела; вместо неё используем вкладку «Тайминг»
from .timing_tab import build_timing_tab as _timing_build_tab
from .finance_tab import build_finance_tab as _finance_build_tab
from .export_tab import build_export_tab as _export_build_tab
# 2.3a Вкладка конвертации PDF→Excel
from .convert_tab import build_convert_tab as _convert_build_tab

# 2.3b Вкладка импорта из Unreal Engine
# Функция build_unreal_tab создаёт интерфейс для загрузки и сопоставления
# таблицы UE5 с позициями каталога. Она располагается в отдельном модуле
# unreal_import_tab.py, аналогично другим вкладкам.
try:
    from .unreal_import_tab import build_unreal_tab as _unreal_build_tab  # type: ignore
except Exception:
    # Если модуль не найден, игнорируем; вкладка UE не будет построена
    _unreal_build_tab = None  # type: ignore


# 3. Класс ProjectPage — основная страница проекта
class ProjectPage(QtWidgets.QWidget):
    # 3.0 Сигналы для синхронизации вкладок
    # Сигнал, излучаемый после изменения сводной сметы (Summary).
    summary_changed = QtCore.Signal()
    # Сигнал, излучаемый после сохранения изменений во вкладке «Бухгалтерия».
    finance_changed = QtCore.Signal()
    # 3.1 Инициализация
    def __init__(self, db: DB, parent=None, log_fn=None):
        super().__init__(parent)
        self.db = db
        self.project_id: Optional[int] = None
        self.project_name: Optional[str] = None
        self.log_fn = log_fn

        # 3.2 Контейнер для отмены последнего действия
        self._last_action: Optional[Dict[str, Any]] = None
        self._last_import_batch: Optional[str] = None

        # 4. Построение интерфейса: вкладки и нижняя строка итога
        root = QtWidgets.QVBoxLayout(self)
        self.tabs = QtWidgets.QTabWidget()

        self.tab_info = QtWidgets.QWidget()
        self.tab_summary = QtWidgets.QWidget()
        self.tab_import = QtWidgets.QWidget()
        # Вкладка тайминга (ранее вкладка создания сметы)
        self.tab_timing = QtWidgets.QWidget()
        self.tab_finance = QtWidgets.QWidget()
        self.tab_export = QtWidgets.QWidget()
        self.tab_convert = QtWidgets.QWidget()

        self.tabs.addTab(self.tab_info, "Информация")
        self.tabs.addTab(self.tab_summary, "Сводная смета")
        self.tabs.addTab(self.tab_import, "Импорт смет")
        self.tabs.addTab(self.tab_timing, "Тайминг")
        self.tabs.addTab(self.tab_finance, "Бухгалтерия")
        self.tabs.addTab(self.tab_export, "Экспорт в PDF")
        # Добавляем вкладку конвертации
        self.tabs.addTab(self.tab_convert, "Конвертация")

        # Создаём вкладку импорта из UE, если доступна функция
        self.tab_unreal = QtWidgets.QWidget()
        if _unreal_build_tab is not None:
            self.tabs.addTab(self.tab_unreal, "Импорт UE")
        else:
            # Если модуль не загрузился, логируем предупреждение
            self._log("Модуль Unreal Import не загружен, вкладка UE недоступна", "error")

        # 4.1 Построение содержимого вкладок
        # Используем вынесенные функции для сборки интерфейса, чтобы уменьшить
        # размер этого файла. В исходном коде эти методы были длинными.
        _info_build_tab(self, self.tab_info)
        _summary_build_tab(self, self.tab_summary)
        _import_build_tab(self, self.tab_import)
        # Построение вкладки тайминга через отдельный модуль
        _timing_build_tab(self, self.tab_timing)
        # Построение вкладки бухгалтерии через отдельный модуль
        _finance_build_tab(self, self.tab_finance)
        # Построение вкладки экспорта в PDF через отдельный модуль
        _export_build_tab(self, self.tab_export)
        # Построение вкладки конвертации через отдельный модуль
        _convert_build_tab(self, self.tab_convert)

        # 4.1a Построение вкладки импорта из UE (если доступно)
        if _unreal_build_tab is not None:
            try:
                _unreal_build_tab(self, self.tab_unreal)
            except Exception as ex:
                # Логируем ошибку построения вкладки, но не останавливаем приложение
                self._log(f"Ошибка построения вкладки UE: {ex}", "error")

        # 4.2 Нижняя панель с итогом по проекту
        bottom = QtWidgets.QHBoxLayout()
        self.label_total = QtWidgets.QLabel("Итого: 0")
        bottom.addStretch(1)
        bottom.addWidget(self.label_total)

        root.addWidget(self.tabs)
        root.addLayout(bottom)

        # 4.3 Переключение вкладок
        # Подключаем обработчик, который будет вызываться при смене вкладки.
        # Он обновляет данные «Бухгалтерии» при переходе на неё и
        # пересчитывает сводные показатели при переходе на «Информацию».
        try:
            self.tabs.currentChanged.connect(self._on_tab_changed)
        except Exception:
            # Если не удалось подключить сигнал, просто логируем ошибку
            self._log("Не удалось подключить обработчик смены вкладок", "error")

        # 4.4 Синхронизация вкладок через сигналы
        # При изменении данных «Бухгалтерии» перезагружаем таблицы сметы и
        # пересчитываем сводные показатели. При изменении сметы обновляем
        # данные бухгалтерии. Эти соединения выполняются здесь один раз.
        try:
            # Связываем сигнал finance_changed с обновлением сметы и инфо
            self.finance_changed.connect(lambda: self._reload_zone_tabs())
            from .info_tab import update_financial_summary  # type: ignore
            self.finance_changed.connect(lambda: update_financial_summary(self))
        except Exception:
            # Ошибка связывания не критична
            self._log("Не удалось связать сигнал finance_changed", "error")

    # 3.3 Логирование
    def _log(self, msg: str, level: str = "info"):
        if callable(self.log_fn):
            self.log_fn(msg, level)

    # 3.4 Пересчёт данных для вкладки «Бухгалтерия»
    def recalc_finance(self) -> None:
        """Обновляет и пересчитывает вкладку «Бухгалтерия».

        Этот метод вызывается из вкладки «Сводная смета» после изменения
        позиций или фильтров. Он загружает актуальные позиции проекта
        напрямую из базы данных (если доступна) или через провайдер
        бухгалтерии, передаёт их в виджет ``FinanceTab`` и инициирует
        пересчёт всех таблиц. После обновления данных также
        переписываются сводные показатели на вкладке «Информация».
        """
        try:
            # 1. Определяем виджет бухгалтерии
            fin = getattr(self, "tab_finance_widget", None)
            if not fin:
                return
            # 2. Собираем текущие позиции проекта
            items = []  # type: List[Item]
            try:
                if self.db is not None and self.project_id is not None:
                    # Читаем строки из базы и создаём объекты Item
                    rows = self.db.list_items(self.project_id)
                    from .finance_tab import Item  # локальный импорт, чтобы избежать циклов
                    for row in rows:
                        try:
                            # sqlite3.Row не поддерживает метод get(), используем доступ по ключу
                            item_id = str(row["id"]) if "id" in row.keys() else ""
                            vendor = ""
                            if "vendor" in row.keys() and row["vendor"] is not None:
                                vendor = str(row["vendor"]).strip()
                            vendor = vendor or "(без подрядчика)"
                            cls_val = "equipment"
                            if "type" in row.keys() and row["type"]:
                                cls_val = row["type"]
                            department = ""
                            if "department" in row.keys() and row["department"]:
                                department = row["department"]
                            zone = ""
                            if "zone" in row.keys() and row["zone"]:
                                zone = row["zone"]
                            name = ""
                            if "name" in row.keys() and row["name"]:
                                name = row["name"]
                            unit_price = 0.0
                            if "unit_price" in row.keys() and row["unit_price"] is not None:
                                unit_price = float(row["unit_price"])
                            qty = 0.0
                            if "qty" in row.keys() and row["qty"] is not None:
                                qty = float(row["qty"])
                            coeff = 1.0
                            if "coeff" in row.keys() and row["coeff"] is not None:
                                coeff = float(row["coeff"])
                            items.append(Item(
                                id=item_id,
                                vendor=vendor,
                                cls=cls_val,
                                department=department,
                                zone=zone,
                                name=name,
                                price=unit_price,
                                qty=qty,
                                coeff=coeff,
                            ))
                        except Exception:
                            # Логируем ошибку и продолжаем со следующей строкой
                            try:
                                from .finance_tab import logger as finance_logger
                                finance_logger.error("project_page.recalc_finance: ошибка обработки строки: %s", traceback.format_exc())
                            except Exception:
                                pass
                            continue
                else:
                    # Если базы нет, используем провайдер из виджета
                    provider = getattr(fin, "provider", None)
                    if provider is not None:
                        try:
                            items = provider.load_items() or []
                        except Exception:
                            items = []
            except Exception:
                items = []
            # 3. Передаём новые позиции в виджет «Бухгалтерия»
            try:
                fin.set_items(items)
                # Запускаем полный пересчёт, чтобы обновить таблицы
                if hasattr(fin, "recalculate_all"):
                    fin.recalculate_all()
            except Exception:
                pass
            # 4. Обновляем вкладку «Информация»
            try:
                from .info_tab import update_financial_summary
                update_financial_summary(self)
            except Exception:
                pass
            # 5. Логируем успешный пересчёт
            try:
                self._log("Вкладка 'Бухгалтерия' пересчитана и обновлена после изменений в смете")
            except Exception:
                pass
        except Exception as ex:
            # Записываем ошибку в лог, не прерываем выполнение
            try:
                self._log(f"Ошибка обновления вкладки 'Бухгалтерия': {ex}", "error")
            except Exception:
                pass

    # 3.5 Обработчик смены вкладок
    def _on_tab_changed(self, index: int) -> None:
        """Вызывается при переключении вкладок в ProjectPage.

        Если пользователь переходит на вкладку «Бухгалтерия», вызываем
        пересчёт финансов для загрузки актуальных данных. Если переходит
        на вкладку «Информация», обновляем сводные показатели.
        """
        try:
            tab_name = self.tabs.tabText(index)
            # При входе на бухгалтерию обновляем данные
            if tab_name == "Бухгалтерия":
                try:
                    if hasattr(self, "recalc_finance"):
                        self.recalc_finance()
                except Exception:
                    pass
            # При входе на вкладку информации обновляем сводные показатели
            elif tab_name == "Информация":
                try:
                    from .info_tab import update_financial_summary  # импорт внутри функции
                    update_financial_summary(self)
                except Exception:
                    pass
        except Exception as ex:
            # Логируем, но не прерываем работу при ошибках
            self._log(f"Ошибка обработки переключения вкладок: {ex}", "error")

    # 5. Вкладка «Информация»
    def _build_info_tab(self, tab: QtWidgets.QWidget):
        # Делегируем построение вкладки функции из модуля info_tab.
        _info_build_tab(self, tab)

    # 6. Вкладка «Сводная смета»
    def _build_summary_tab(self, tab: QtWidgets.QWidget):
        # Делегируем построение вкладки функции из summary_tab, затем прерываем
        _summary_build_tab(self, tab)
        return
        v = QtWidgets.QVBoxLayout(tab)

        # 6.1 Фильтры
        filt = QtWidgets.QHBoxLayout()
        filt.setContentsMargins(4, 4, 4, 2)
        filt.setSpacing(8)

        self.ed_search = QtWidgets.QLineEdit()
        self.ed_search.setPlaceholderText("Поиск по наименованию...")

        self.cmb_f_vendor = QtWidgets.QComboBox()
        self.cmb_f_vendor.addItem("<Все подрядчики>")

        self.cmb_f_department = QtWidgets.QComboBox()
        self.cmb_f_department.addItem("<Все отделы>")

        self.cmb_f_class = QtWidgets.QComboBox()
        self.cmb_f_class.addItem("<Все классы>")
        self.cmb_f_class.addItems(list(CLASS_RU2EN.keys()))

        self.btn_delete_selected = QtWidgets.QPushButton("Удалить выделенные")

        for w in (
            QtWidgets.QLabel("Поиск:"), self.ed_search,
            QtWidgets.QLabel("Подрядчик:"), self.cmb_f_vendor,
            QtWidgets.QLabel("Отдел:"), self.cmb_f_department,
            QtWidgets.QLabel("Класс:"), self.cmb_f_class,
            self.btn_delete_selected,
        ):
            filt.addWidget(w)
        filt.addStretch(1)

        # 6.2 Панель зон
        zbar = QtWidgets.QHBoxLayout()
        zbar.setContentsMargins(4, 2, 4, 2)
        zbar.setSpacing(8)

        self.ed_new_zone = QtWidgets.QLineEdit()
        self.ed_new_zone.setPlaceholderText("Новая зона…")
        self.ed_new_zone.setMinimumWidth(200)

        self.btn_add_zone = QtWidgets.QPushButton("Создать зону")

        self.cmb_move_zone = QtWidgets.QComboBox()
        self.cmb_move_zone.setEditable(True)
        self.cmb_move_zone.setMinimumWidth(180)

        self.btn_move_zone = QtWidgets.QPushButton("Перенести в зону")

        for w in (
            QtWidgets.QLabel("Зоны:"), self.ed_new_zone, self.btn_add_zone,
            QtWidgets.QLabel("→"), self.cmb_move_zone, self.btn_move_zone
        ):
            zbar.addWidget(w)
        zbar.addStretch(1)

        # 6.2.1 Кнопка «Отменить…»
        self.btn_undo_summary = QtWidgets.QPushButton("Отменить последнее действие")
        self.btn_undo_summary.setEnabled(False)
        zbar.addWidget(self.btn_undo_summary)

        # 6.3 Добавление позиций вручную
        addbar1 = QtWidgets.QHBoxLayout()
        addbar1.setContentsMargins(4, 2, 4, 2)
        addbar1.setSpacing(8)

        self.ed_add_name = QtWidgets.QLineEdit()
        self.ed_add_name.setPlaceholderText("Наименование")
        self.ed_add_name.setMinimumWidth(260)

        self.sp_add_qty = SmartDoubleSpinBox()
        self.sp_add_qty.setDecimals(3)
        self.sp_add_qty.setMinimum(0.001)
        self.sp_add_qty.setValue(1.000)

        self.sp_add_coeff = SmartDoubleSpinBox()
        self.sp_add_coeff.setDecimals(3)
        self.sp_add_coeff.setMinimum(0.001)
        self.sp_add_coeff.setValue(1.000)

        self.cmb_add_class = QtWidgets.QComboBox()
        self.cmb_add_class.addItems(list(CLASS_RU2EN.keys()))

        for w in (
            self.ed_add_name,
            QtWidgets.QLabel("Кол-во:"), self.sp_add_qty,
            QtWidgets.QLabel("Коэф.:"), self.sp_add_coeff,
            QtWidgets.QLabel("Класс:"), self.cmb_add_class
        ):
            addbar1.addWidget(w)
        addbar1.addStretch(1)

        addbar2 = QtWidgets.QHBoxLayout()
        addbar2.setContentsMargins(4, 2, 4, 4)
        addbar2.setSpacing(8)

        self.ed_add_vendor = QtWidgets.QLineEdit()
        self.ed_add_vendor.setPlaceholderText("Подрядчик")

        self.ed_add_department = QtWidgets.QLineEdit()
        self.ed_add_department.setPlaceholderText("Отдел")

        self.cmb_add_zone = QtWidgets.QComboBox()
        self.cmb_add_zone.setMinimumWidth(160)

        self.sp_add_power = SmartDoubleSpinBox()
        self.sp_add_power.setDecimals(0)
        self.sp_add_power.setMaximum(10**9)
        self.sp_add_power.setSuffix(" Вт")

        # ВАЖНО: поле «Цена/шт»
        self.sp_add_price = SmartDoubleSpinBox()
        self.sp_add_price.setDecimals(2)
        self.sp_add_price.setMinimum(0.0)
        self.sp_add_price.setMaximum(10**9)
        self.sp_add_price.setValue(0.0)

        self.btn_add_manual = QtWidgets.QPushButton("Добавить позицию")

        addbar2.addWidget(QtWidgets.QLabel("Цена/шт:"))
        addbar2.addWidget(self.sp_add_price)
        addbar2.addWidget(self.ed_add_vendor)
        addbar2.addWidget(self.ed_add_department)
        addbar2.addWidget(QtWidgets.QLabel("Зона:"))
        addbar2.addWidget(self.cmb_add_zone)
        addbar2.addWidget(QtWidgets.QLabel("Потр.:"))
        addbar2.addWidget(self.sp_add_power)
        addbar2.addWidget(self.btn_add_manual)
        addbar2.addStretch(1)

        # 6.4 Табы зон
        self.zone_tabs = QtWidgets.QTabWidget()
        self.zone_tables: Dict[str, QtWidgets.QTableWidget] = {}

        # 6.5 Компоновка
        v.addLayout(filt)
        v.addLayout(zbar)
        v.addLayout(addbar1)
        v.addLayout(addbar2)
        v.addWidget(self.zone_tabs, 1)

        # 6.6 Сигналы
        self.ed_search.textChanged.connect(self._reload_zone_tabs)
        self.cmb_f_vendor.currentTextChanged.connect(self._reload_zone_tabs)
        self.cmb_f_department.currentTextChanged.connect(self._reload_zone_tabs)
        self.cmb_f_class.currentTextChanged.connect(self._reload_zone_tabs)
        self.btn_delete_selected.clicked.connect(self.delete_selected)
        self.btn_add_zone.clicked.connect(self._create_zone)
        self.btn_move_zone.clicked.connect(self._move_selected_to_zone)
        self.btn_add_manual.clicked.connect(self._add_manual_item)
        self.btn_undo_summary.clicked.connect(self._undo_last_summary)

    # 6.7 Таблица зоны
    def _build_zone_table(self) -> QtWidgets.QTableWidget:
        cols = ["Наименование", "Кол-во", "Коэф.", "Цена/шт", "Сумма",
                "Подрядчик", "Отдел", "Зона", "Класс", "Потребление (Вт)"]
        t = QtWidgets.QTableWidget(0, len(cols))
        t.setHorizontalHeaderLabels(cols)
        t.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        t.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        t.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked |
            QtWidgets.QAbstractItemView.SelectedClicked
        )
        t.setWordWrap(False)
        t.setItemDelegateForColumn(0, WrapTextDelegate(t, wrap_threshold=WRAP_THRESHOLD))
        t.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        setup_priority_name(t, name_col=0)
        return t

    # 6.8 Инициализация табов зон
    def _init_zone_tabs(self):
        # Делегируем инициализацию вкладок зон функции из summary_tab
        init_zone_tabs_ext(self)

    # 6.9 Комбо зон для ручного добавления
    def _fill_manual_zone_combo(self, zones: List[str]):
        # Делегируем заполнение списка зон функции из summary_tab
        fill_manual_zone_combo_ext(self, zones)

    # 6.10 Обновление таблиц по фильтрам
    def _reload_zone_tabs(self):
        # Делегируем перезагрузку зон функции из summary_tab
        reload_zone_tabs_ext(self)

    # 6.11 Создание зоны
    def _create_zone(self):
        # Делегируем создание зоны функции из summary_tab
        create_zone_ext(self)

    # 6.12 Перенос выделенных строк в другую зону
    def _move_selected_to_zone(self):
        """Перенос выделенных строк в выбранную зону (делегировано в summary_tab)."""
        move_selected_to_zone_ext(self)

    def _add_manual_item(self):
        """Добавление позиции вручную (делегировано в summary_tab)."""
        add_manual_item_ext(self)

    def _on_summary_item_changed(self, item: QtWidgets.QTableWidgetItem):
        # Делегируем обработку изменения ячейки функции из summary_tab
        on_summary_item_changed_ext(self, item)
        return
        table = item.tableWidget()
        row = item.row()
        col = item.column()

    # 6.15 Удаление выделенных строк (снимок для UNDO)
    def delete_selected(self):
        # Делегируем удаление выбранных строк функции из summary_tab
        delete_selected_ext(self)

    # 7. Вкладка «Импорт смет»
    def _build_import_tab(self, tab: QtWidgets.QWidget):
        # Делегируем построение вкладки функции из import_tab и завершаем
        _import_build_tab(self, tab)
        return
        # 7.1 Переменные импорта
        self._import_file: Optional[Path] = None
        self._src_headers: List[str] = []
        self._src_rows: List[List[Any]] = []
        self._mapping_widgets: List[QtWidgets.QComboBox] = []
        self._result_items: List[Dict[str, Any]] = []

        # 7.2 UI: верхние панели
        root = QtWidgets.QVBoxLayout(tab)

        top1 = QtWidgets.QHBoxLayout()
        self.drop_label = FileDropLabel(accept_exts=(".xlsx", ".csv"), on_file=self._on_file_dropped)
        self.drop_label.setMinimumHeight(48)
        self.btn_choose_file = QtWidgets.QPushButton("Выбрать файл (XLSX/CSV)")
        self.btn_choose_file.clicked.connect(self._choose_file)
        top1.addWidget(self.drop_label, 1)
        top1.addWidget(self.btn_choose_file)

        top2 = QtWidgets.QHBoxLayout()
        self.combo_vendor = QtWidgets.QComboBox()
        self.combo_vendor.setEditable(True)
        self.combo_vendor.setPlaceholderText("Подрядчик (обязательно)")

        self.combo_department = QtWidgets.QComboBox()
        self.combo_department.setEditable(True)
        self.combo_department.setPlaceholderText("Отдел")

        self.combo_zone = QtWidgets.QComboBox()
        self.combo_zone.setEditable(True)
        self.combo_zone.setPlaceholderText("Зона (опционально)")

        top2.addWidget(QtWidgets.QLabel("Подрядчик:"))
        top2.addWidget(self.combo_vendor, 1)
        top2.addWidget(QtWidgets.QLabel("Отдел:"))
        top2.addWidget(self.combo_department, 1)
        top2.addWidget(QtWidgets.QLabel("Зона:"))
        top2.addWidget(self.combo_zone, 1)

        top3 = QtWidgets.QHBoxLayout()
        self.chk_import_power = QtWidgets.QCheckBox("Импортировать потребление")
        self.combo_power_unit = QtWidgets.QComboBox()
        self.combo_power_unit.addItems(["Вт", "кВт", "А"])
        self.chk_filter_itogo = QtWidgets.QCheckBox("Отфильтровать «Итого»")
        self.chk_filter_empty = QtWidgets.QCheckBox("Убрать пустые строки")
        self.chk_filter_no_price_amount = QtWidgets.QCheckBox("Убрать строки без цены и суммы")

        top3.addWidget(self.chk_import_power)
        top3.addWidget(QtWidgets.QLabel("Ед. изм. нагрузки:"))
        top3.addWidget(self.combo_power_unit)
        top3.addStretch(1)
        top3.addWidget(self.chk_filter_itogo)
        top3.addWidget(self.chk_filter_empty)
        top3.addWidget(self.chk_filter_no_price_amount)

        # 7.3 Панель сопоставления столбцов
        self.map_scroll = QtWidgets.QScrollArea()
        self.map_scroll.setWidgetResizable(True)
        self.map_host = QtWidgets.QWidget()
        self.map_layout = QtWidgets.QHBoxLayout(self.map_host)
        self.map_layout.setContentsMargins(6, 4, 6, 4)
        self.map_layout.setSpacing(8)
        self.map_scroll.setWidget(self.map_host)
        self.map_scroll.setFixedHeight(86)

        # 7.4 Таблицы предпросмотра
        mid = QtWidgets.QHBoxLayout()
        left_v = QtWidgets.QVBoxLayout()
        right_v = QtWidgets.QVBoxLayout()

        self.tbl_src = QtWidgets.QTableWidget(0, 0)
        setup_auto_col_resize(self.tbl_src)
        self.tbl_src.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        self.tbl_dst = QtWidgets.QTableWidget(0, 9)
        self.tbl_dst.setHorizontalHeaderLabels([
            "Наименование", "Кол-во", "Коэф.", "Цена/шт", "Сумма",
            "Потребл. (Вт)", "Класс (РУС)", "Подрядчик", "Отдел/Зона"
        ])
        setup_auto_col_resize(self.tbl_dst)

        left_v.addWidget(QtWidgets.QLabel("Исходная таблица"))
        left_v.addWidget(self.tbl_src, 1)
        right_v.addWidget(QtWidgets.QLabel("Результат импорта"))
        right_v.addWidget(self.tbl_dst, 1)

        sum_row = QtWidgets.QHBoxLayout()
        self.lbl_sum_amount = QtWidgets.QLabel("Сумма импорта: 0 ₽")
        self.lbl_sum_power = QtWidgets.QLabel("Потребление: 0 кВт")
        sum_row.addWidget(self.lbl_sum_amount)
        sum_row.addStretch(1)
        sum_row.addWidget(self.lbl_sum_power)
        right_v.addLayout(sum_row)

        mid.addLayout(left_v, 1)
        mid.addLayout(right_v, 1)

        # 7.5 Кнопки действий
        actions = QtWidgets.QHBoxLayout()
        self.btn_prepare = QtWidgets.QPushButton("Обновить предпросмотр")
        self.btn_import = QtWidgets.QPushButton("Импорт в проект")
        self.btn_undo = QtWidgets.QPushButton("Отменить импорт")
        self.btn_import.setEnabled(False)
        self.btn_undo.setEnabled(False)
        actions.addWidget(self.btn_prepare)
        actions.addStretch(1)
        actions.addWidget(self.btn_undo)
        actions.addWidget(self.btn_import)

        root.addLayout(top1)
        root.addLayout(top2)
        root.addLayout(top3)
        root.addWidget(self.map_scroll)
        root.addLayout(mid)
        root.addLayout(actions)

        # 7.6 Сигналы импорта
        self.chk_filter_itogo.toggled.connect(self._rebuild_result)
        self.chk_filter_empty.toggled.connect(self._rebuild_result)
        self.chk_filter_no_price_amount.toggled.connect(self._rebuild_result)
        self.chk_import_power.toggled.connect(self._rebuild_result)
        self.combo_power_unit.currentIndexChanged.connect(self._rebuild_result)
        self.btn_prepare.clicked.connect(self._rebuild_result)
        self.btn_import.clicked.connect(self._apply_import)
        self.btn_undo.clicked.connect(self._undo_last_import)
        self.combo_vendor.editTextChanged.connect(self._update_import_button_state)
        self.combo_vendor.currentTextChanged.connect(self._update_import_button_state)

    # 7.7 Вспомогательные методы импорта
    @staticmethod
    def _to_float(x) -> float:
        return to_float(x, 0.0)

    def _is_itogo(self, name: str) -> bool:
        return "итог" in (name or "").strip().lower()

    def _update_import_button_state(self):
        """Обновляет состояние кнопки импорта, используя ``import_tab.update_import_button_state``.

        Ранее эта функция проверяла наличие подрядчика, присутствие колонки
        «Наименование» и непустоту списка результатов. Теперь эта проверка
        вынесена в модуль ``import_tab``. Вызов внешней функции выполняет
        те же действия и завершает работу, предотвращая исполнение старого кода.
        """
        update_import_button_state_ext(self)
        return

    def _refresh_vendor_dept_zone_lists(self):
        """Обновляет списки подрядчиков/отделов/зон с помощью ``import_tab.refresh_vendor_dept_zone_lists``.

        Внешняя функция выполняет обращение к базе данных, заполнение
        соответствующих комбобоксов и обновление списка зон для ручного
        добавления. Вызов данной функции завершает выполнение текущей.
        """
        refresh_vendor_dept_zone_lists_ext(self)
        return

    # 8. Загрузка проекта
    def load_project(self, project_id: int, project_name: str):
        self.project_id = project_id
        self.project_name = project_name
        self.cover_label.set_project_id(self.project_id)

        self._load_info_json()

        self.cmb_f_vendor.blockSignals(True)
        self.cmb_f_department.blockSignals(True)

        self.cmb_f_vendor.clear()
        self.cmb_f_vendor.addItem("<Все подрядчики>")
        for vnd in self.db.project_distinct_values(self.project_id, "vendor"):
            self.cmb_f_vendor.addItem(vnd)

        self.cmb_f_department.clear()
        self.cmb_f_department.addItem("<Все отделы>")
        for dep in self.db.project_distinct_values(self.project_id, "department"):
            self.cmb_f_department.addItem(dep)

        self.cmb_f_vendor.blockSignals(False)
        self.cmb_f_department.blockSignals(False)

        self._init_zone_tabs()
        self._reload_zone_tabs()
        self._refresh_vendor_dept_zone_lists()

        self._log(f"Открыт проект: id={self.project_id}, name='{self.project_name}'")
        # 8.x Загрузка тайминга проекта, если функция определена
        try:
            if hasattr(self, "load_timing_data") and callable(self.load_timing_data):
                self.load_timing_data()
        except Exception as ex:
            self._log(f"Ошибка загрузки тайминга: {ex}", "error")

        # 8.y Обновление провайдера в «Бухгалтерии»
        # После выбора проекта обязательно переключаем провайдер вкладки «Бухгалтерия»
        # на DBDataProvider, чтобы читать данные из базы вместо файлов. Это
        # обеспечивает корректное отображение актуальной сметы.
        try:
            fin = getattr(self, "tab_finance_widget", None)
            if fin:
                # Импортируем DBDataProvider только здесь, чтобы избежать циклов
                from .finance_tab import DBDataProvider  # type: ignore
                # Передаём текущую страницу в провайдер для доступа к self.db и project_id
                new_provider = DBDataProvider(self)
                # Обновляем провайдер и перезагружаем данные
                if hasattr(fin, "set_provider"):
                    fin.set_provider(new_provider)
                else:
                    # На случай, если метод отсутствует (старые версии)
                    fin.provider = new_provider
                    items = new_provider.load_items() or []
                    fin.set_items(items)
            # Обновляем сводные показатели после смены провайдера
            try:
                from .info_tab import update_financial_summary  # type: ignore
                update_financial_summary(self)
            except Exception:
                pass
        except Exception as ex:
            self._log(f"Ошибка обновления провайдера бухгалтерии: {ex}", "error")

    # 8.1 Информация: пути/загрузка/сохранение JSON
    def _info_json_path(self) -> Path:
        # Используем функцию из info_tab для получения пути к JSON
        return info_json_path_ext(self)

    def _load_info_json(self):
        # Делегируем загрузку данных вкладки «Информация» функции из info_tab
        load_info_json_ext(self)

    def _save_info_json(self):
        # Делегируем сохранение данных вкладки «Информация» функции из info_tab
        save_info_json_ext(self)

    # 9. Импорт смет: выбор, чтение, сбор предпросмотра и запись
    def _choose_file(self):
        """Открывает диалог выбора файла для импорта и делегирует обработку.

        Вместо локальной реализации, использующей ``QFileDialog`` и вызов
        ``_on_file_dropped``, функция теперь обращается к ``choose_file`` из
        модуля ``import_tab``. Внешняя функция уже реализует диалог выбора
        файла и передачу выбранного пути дальше. После вызова возвращаем
        управление, оставляя исходный код ниже в качестве справки (он не
        исполняется из-за ``return``).
        """
        choose_file_ext(self)
        return

    def _on_file_dropped(self, p: Path):
        """Обрабатывает перетаскивание файла, делегируя логику в ``import_tab``.

        Внешняя функция ``on_file_dropped`` в модуле ``import_tab`` выполняет
        установку выбранного файла, чтение данных, формирование панели
        сопоставления, заполнение исходной таблицы и сбор агрегированного
        результата. После вызова возвращаем управление. Старый код
        сохранён ниже для справки.
        """
        on_file_dropped_ext(self, p)
        return

    def _read_source_file(self, path: Path):
        """Считывает файл источника через ``import_tab.read_source_file``.

        В исходной реализации функция содержала подробный алгоритм поиска
        строки заголовков в XLSX/CSV и считывания данных. Эта логика
        перенесена в модуль ``import_tab``. Вызываем соответствующую
        функцию и завершаем выполнение. Код ниже оставлен для справки и не
        исполняется из-за ``return``.
        """
        read_source_file_ext(self, path)
        return

    def _build_mapping_bar(self):
        """Строит панель сопоставления столбцов через ``import_tab.build_mapping_bar``.

        Прежняя реализация вручную создавала виджеты для каждой колонки и
        предзаполняла выбор типа данных. Теперь эта логика перенесена в
        модуль ``import_tab``. Вызов внешней функции выполняет все нужные
        действия, после чего выполнение завершается. Старый код приведён
        ниже для справки.
        """
        build_mapping_bar_ext(self)
        return

    def _fill_src_table(self):
        """Заполняет таблицу исходных данных через ``import_tab.fill_src_table``.

        Раньше таблица заполнялась вручную здесь. Теперь управление отдаётся
        функции ``fill_src_table`` из модуля ``import_tab``. После вызова
        исполнение прекращается.
        """
        fill_src_table_ext(self)
        return

    def _current_mapping(self) -> Dict[int, str]:
        """Возвращает словарь сопоставлений колонок, делегируя в ``import_tab.current_mapping``.

        Вместо локального вычисления словаря на основе текущих значений
        комбобоксов вызывается одноимённая функция из модуля ``import_tab``.
        Это позволяет избежать дублирования кода. Возвращаем результат
        внешней функции.
        """
        return current_mapping_ext(self)

    def _rebuild_result(self):
        """Пересобирает результат импорта, вызывая ``import_tab.rebuild_result``.

        Аггрегирование данных, фильтрация, вычисление сумм и мощностей,
        заполнение таблицы и обновление состояний теперь реализованы во
        внешнем модуле ``import_tab``. Мы вызываем функцию ``rebuild_result``
        оттуда и завершаем выполнение текущей функции. Старый код оставлен
        ниже для справки.
        """
        rebuild_result_ext(self)
        return

    def _apply_import(self):
        """Выполняет импорт через ``import_tab.apply_import``.

        Проверки обязательных полей, формирование записей для базы данных и каталога,
        возможный диалог корректировки мощности, запись и логирование теперь
        инкапсулированы в ``apply_import`` модуля ``import_tab``. Мы просто
        вызываем эту функцию и завершаем выполнение. Старый код оставлен
        ниже для справки, но не исполняется.
        """
        apply_import_ext(self)
        return

    def _undo_last_import(self):
        """Отменяет последний импорт, используя ``import_tab.undo_last_import``.

        Внешняя функция выполняет удаление импортированных позиций, вывод
        сообщений об успехе/ошибке, обновление состояния и очищение
        ``_last_import_batch``. Мы вызываем её и прерываем дальнейшее
        исполнение. Исходный код оставлен ниже для справки.
        """
        undo_last_import_ext(self)
        return

    # 10. Заглушка вкладок
    def _build_placeholder(self, tab: QtWidgets.QWidget, text: str):
        v = QtWidgets.QVBoxLayout(tab)
        label = QtWidgets.QLabel(text)
        label.setWordWrap(True)
        v.addWidget(label)
        v.addStretch(1)

    # 11. Отмена последнего действия (смета)
    def _undo_last_summary(self):
        """Отменяет последнее действие в сводной смете через ``summary_tab.undo_last_summary``.

        Все разновидности отмены (ручное добавление, удаление, перенос,
        редактирование) реализованы во внешней функции ``undo_last_summary``
        модуля ``summary_tab``. Вызываем её и сразу завершаем выполнение,
        сохранив оригинальную реализацию ниже для справки.
        """
        undo_last_summary_ext(self)
        return
