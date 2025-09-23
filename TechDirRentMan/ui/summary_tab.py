"""
Модуль ``summary_tab`` выделяет логику работы вкладки «Сводная смета».

Назначение
------------

Этот файл инкапсулирует создание интерфейса и обработку событий для вкладки
«Сводная смета». Ранее вся логика находилась в ``ProjectPage``, что делало
поддержку сложной. Вынесение кода в отдельный модуль улучшает структуру
проекта и позволяет проще ориентироваться в функциональности.

Вкладка поддерживает полноценный механизм снимков: пользователь может
зафиксировать текущее состояние сметы, сохранить его в файл и в любой
момент включить режим сравнения, чтобы увидеть изменения относительно
сохранённого снимка. Снимок включает список зон и копию каждой позиции
проекта. При включении сравнения таблицы расширяются дополнительными
колонками и отображают разницу по количеству, цене и сумме.

Секции и их назначение
----------------------

* ``build_summary_tab`` — создает интерфейс вкладки и привязывает сигналы к методам страницы.
* ``init_zone_tabs`` — инициализирует список зон и таблиц.
* ``fill_manual_zone_combo`` — обновляет выпадающий список зон для ручного добавления.
* ``fill_manual_dept_combo`` — обновляет выпадающий список отделов для ручного добавления.
* ``reload_zone_tabs`` — перезагружает таблицы зон с учётом фильтров и режима сравнения.
* ``create_zone`` — создает новую зону.
* ``rename_zone`` — переименовывает текущую зону, обновляя данные в базе и интерфейсе.
* ``move_selected_to_zone`` — переносит выделенные позиции в другую зону, дробя строку при частичном переносе.
* ``add_manual_item`` — добавляет позицию вручную в проект и каталог.
* ``on_summary_item_changed`` — обрабатывает изменение количества, коэффициента или цены.
* ``delete_selected`` — удаляет выделенные позиции со снятием снимка для операции отмены.
* ``undo_last_summary`` — отменяет последнее действие (ручное добавление, удаление, перенос или редактирование).
* ``take_snapshot`` — делает временный снимок текущей сметы для последующего сравнения.
* ``toggle_snapshot_compare`` — включает/выключает режим сравнения со снимком, проверяя совместимость зон.
* ``save_snapshot`` — сохраняет снимок в файл и обновляет список сохранённых снимков.
* ``load_snapshot_list`` — загружает список сохранённых снимков для текущего проекта в выпадающий список.
* ``on_snapshot_selected`` — загружает выбранный снимок из файла и сбрасывает режим сравнения.

    Дополнительная логика:

    * При отображении стандартной сводной сметы (без режима сравнения) одинаковые позиции
      (одинаковое наименование и подрядчик внутри одной зоны/отдела) агрегируются: их
      количества и суммы складываются, а коэффициент и цена вычисляются как
      средневзвешенные значения. Это позволяет избежать дублирования строк, если одна
      позиция была добавлена несколько раз (вручную или импортом). Смотрите реализацию
      в ``reload_zone_tabs``.

Стиль
-----

Код разбит на небольшие функции, каждая из которых отвечает за отдельную
задачу. Внутри функций присутствуют краткие комментарии, поясняющие
ключевые операции. Для диагностики используются вызовы ``page._log`` —
информационные сообщения и ошибки отображаются в журнале приложения.
"""

from __future__ import annotations

from typing import List, Dict, Any, Tuple, Set, Optional
from datetime import datetime

from PySide6 import QtWidgets, QtCore, QtGui

from .common import (
    CLASS_RU2EN, CLASS_EN2RU, WRAP_THRESHOLD, fmt_num, fmt_sign, to_float,
    apply_auto_col_resize, setup_priority_name, normalize_case, DATA_DIR,
    # Импортируем функции для канонического поиска
    make_search_key, contains_search
)
from .delegates import WrapTextDelegate
from .widgets import SmartDoubleSpinBox
from .unreal_import_tab import CatalogSelectDialog  # реиспользуем диалог выбора позиции из базы

import json
from pathlib import Path
import logging

def compute_fin_snapshot_data(page: Any) -> Dict[str, Any]:
    """
    Собирает агрегированные данные для финансового отчёта.

    Возвращает словарь, содержащий итоговую сумму с учётом налогов
    по подрядчикам, зонам, отделам и классам, а также общую сумму проекта.

    Структура результата:

        {
            "vendors": {vendor_name: total_with_tax},
            "zones": {zone_name: total_with_tax},
            "departments": {dept_name: total_with_tax},
            "classes": {class_key: total_with_tax},
            "project_total": total_sum_with_tax
        }

    Данные собираются аналогично расчётам в ``_build_fin_report``: суммы
    агрегируются по всем позициям, используя эффективный коэффициент,
    налог берётся из ``preview_tax_pct`` вкладки «Бухгалтерия». Если
    виджет бухгалтерии недоступен, данные берутся напрямую из БД.

    :param page: объект ProjectPage
    :return: словарь агрегированных данных
    """
    try:
        # Получаем доступ к виджету «Бухгалтерия» для чтения настроек
        ft = getattr(page, "tab_finance_widget", None)
        items_for_totals: List[Any] = []
        if ft and hasattr(ft, "items"):
            # Используем текущие элементы вкладки «Бухгалтерия»
            items_for_totals = list(ft.items)
        elif getattr(page, "db", None) and getattr(page, "project_id", None):
            # Загружаем позиции через провайдера при отсутствии виджета
            try:
                from .finance_tab import DBDataProvider  # type: ignore
                prov = DBDataProvider(page)
                items_for_totals = prov.load_items() or []
            except Exception:
                items_for_totals = []
        # Формируем карту налогов по подрядчикам на основе preview_tax_pct (в процентах)
        vendor_tax_map: Dict[str, float] = {}
        try:
            if ft and hasattr(ft, "preview_tax_pct"):
                for v in ft.preview_tax_pct:
                    try:
                        vendor_tax_map[normalize_case(v)] = float(ft.preview_tax_pct.get(v, 0.0)) / 100.0  # type: ignore
                    except Exception:
                        continue
        except Exception:
            vendor_tax_map = {}
        # Инициализируем структуры для сумм
        summary_vendor: Dict[str, float] = {}
        summary_vendor_tax: Dict[str, Tuple[float, float]] = {}
        summary_zone: Dict[str, float] = {}
        summary_zone_tax: Dict[str, Tuple[float, float]] = {}
        summary_dept: Dict[str, float] = {}
        summary_dept_tax: Dict[str, Tuple[float, float]] = {}
        summary_cls: Dict[str, float] = {}
        summary_cls_tax: Dict[str, Tuple[float, float]] = {}
        total_project: float = 0.0
        total_project_tax: float = 0.0
        # Используем агрегированные данные ft._agg_latest для сумм по подрядчикам
        try:
            if ft and hasattr(ft, "_agg_latest") and ft._agg_latest:
                agg = ft._agg_latest  # type: ignore
                for vend, data in agg.items():
                    try:
                        total_sum = float(data.get("equip_sum", 0.0)) + float(data.get("other_sum", 0.0))
                        summary_vendor[vend] = summary_vendor.get(vend, 0.0) + total_sum
                        # Налог и общая сумма с налогом для подрядчика
                        t_pct = vendor_tax_map.get(normalize_case(vend), 0.0)
                        tax_amt = total_sum * t_pct
                        summary_vendor_tax[vend] = (
                            summary_vendor_tax.get(vend, (0.0, 0.0))[0] + tax_amt,
                            summary_vendor_tax.get(vend, (0.0, 0.0))[1] + total_sum + tax_amt,
                        )
                    except Exception:
                        continue
        except Exception:
            pass
        # Проходим по всем позициям для сумм по зонам, отделам и классам
        for it in items_for_totals:
            try:
                # Определяем эффективный коэффициент для equipment
                eff: Optional[float] = None
                if ft:
                    v = it.vendor or ""
                    if getattr(it, "cls", "equipment") == "equipment":
                        try:
                            if ft.preview_coeff_enabled.get(v, True):
                                eff = float(ft._coeff_user_values.get(v, ft.preview_vendor_coeffs.get(v, 1.0)))
                            else:
                                if getattr(it, "original_coeff", None) is not None:
                                    eff = float(getattr(it, "original_coeff"))
                        except Exception:
                            eff = None
                # Сумма позиции без налога
                try:
                    amt = float(it.amount(effective_coeff=eff))
                except Exception:
                    amt = 0.0
                # Накопление суммарных значений
                total_project += amt
                zone_name = (it.zone or "Без зоны").strip() or "Без зоны"
                dept_name = (it.department or "Без отдела").strip() or "Без отдела"
                cls_en = getattr(it, "cls", "equipment")
                summary_zone[zone_name] = summary_zone.get(zone_name, 0.0) + amt
                summary_dept[dept_name] = summary_dept.get(dept_name, 0.0) + amt
                summary_cls[cls_en] = summary_cls.get(cls_en, 0.0) + amt
                # Рассчитываем налог для позиции исходя из ставки подрядчика
                try:
                    tax_pct = vendor_tax_map.get(normalize_case(it.vendor or ""), 0.0)
                except Exception:
                    tax_pct = 0.0
                tax_amt = amt * tax_pct
                total_project_tax += tax_amt
                # Обновляем суммарные налоги и суммы с налогом для зон, отделов и классов
                prev_tax, prev_total = summary_zone_tax.get(zone_name, (0.0, 0.0))
                summary_zone_tax[zone_name] = (prev_tax + tax_amt, prev_total + amt + tax_amt)
                prev_tax, prev_total = summary_dept_tax.get(dept_name, (0.0, 0.0))
                summary_dept_tax[dept_name] = (prev_tax + tax_amt, prev_total + amt + tax_amt)
                prev_tax, prev_total = summary_cls_tax.get(cls_en, (0.0, 0.0))
                summary_cls_tax[cls_en] = (prev_tax + tax_amt, prev_total + amt + tax_amt)
            except Exception:
                continue
        # Формируем результат
        result: Dict[str, Any] = {
            "vendors": {},
            "zones": {},
            "departments": {},
            "classes": {},
            "project_total": total_project + total_project_tax,
        }
        # Заполняем суммы по подрядчикам (с учётом налога)
        for vend in summary_vendor:
            try:
                if vend in summary_vendor_tax:
                    result["vendors"][vend] = summary_vendor_tax[vend][1]
                else:
                    result["vendors"][vend] = summary_vendor[vend]
            except Exception:
                continue
        # Заполняем суммы по зонам, отделам и классам
        for z, (_, total_with_tax) in summary_zone_tax.items():
            result["zones"][z] = total_with_tax
        for d, (_, total_with_tax) in summary_dept_tax.items():
            result["departments"][d] = total_with_tax
        for cls, (_, total_with_tax) in summary_cls_tax.items():
            result["classes"][cls] = total_with_tax
        return result
    except Exception as ex:
        # При ошибке логируем подробности и возвращаем пустую структуру
        logging.getLogger(__name__).error("Ошибка расчёта снимка финансового отчёта: %s", ex, exc_info=True)
        return {}

# 0. Персистенция списка зон
# ----------------------------
# При создании или переименовании зоны необходимо сохранять её название,
# чтобы оно было доступно после перезапуска приложения даже если в зоне
# пока нет позиций. Сохраняем список зон в JSON-файл, расположенный в
# корневой папке данных (A37/data) рядом с другими проектными файлами.
# Для проектов без идентификатора используется файл project_default_zones.json.

def _zones_json_path(page: Any) -> Path:
    """
    Возвращает путь к JSON-файлу со списком зон для текущего проекта.

    Используем корневую директорию 'data' (на три уровня выше ui), чтобы
    сохранить информацию на уровне проекта, аналогично файлам info_json.
    Если project_id не задан, используется имя 'default'.
    """
    pid = getattr(page, "project_id", None)
    pid_str = "default" if pid is None else str(pid)
    # Находим корневую папку A37: three parents up from this file
    root_data = Path(__file__).resolve().parents[3] / "data"
    return root_data / f"project_{pid_str}_zones.json"



# == FIX A68: Canonical zone handling =========================================
# 1) "Без зоны" хранится только как пустой ключ "" в БД/внутри приложения.
# 2) В JSON-персистентности НЕ храним пустую зону и строку "Без зоны".
# 3) Дедупликация зон ведётся без учёта регистра.
def _is_no_zone(name: str) -> bool:
    if name is None:
        return True
    s = str(name).strip()
    return s == "" or s.lower() == "без зоны"

def _canon_zone(name: str) -> str:
    """Нормализует имя зоны для хранения/сравнения."""
    if _is_no_zone(name):
        return ""
    return str(name).strip()

def _canonize_list(zones: List[str]) -> List[str]:
    """Очищает список зон: убирает пустые и 'Без зоны', удаляет дубликаты (case-insensitive)."""
    out: List[str] = []
    seen: Set[str] = set()
    for z in zones or []:
        c = _canon_zone(z)
        if not c:
            continue
        k = c.lower()
        if k not in seen:
            seen.add(k)
            out.append(c)
    return out
# == /FIX A68 ================================================================



def _load_persisted_zones(page: Any) -> List[str]:
    """
    Загружает список сохранённых зон из JSON-файла и нормализует их.
    Возвращает список БЕЗ пустой зоны и без строки "Без зоны".
    В случае наличия лишних значений выполняется авто-санитизация файла.
    """
    p = _zones_json_path(page)
    try:
        raw: List[str] = []
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(data, list):
                raw = [str(z) for z in data]
        cleaned = _canonize_list(raw)
        # Автоматически санитизируем файл при расхождении
        try:
            if cleaned != raw:
                _save_persisted_zones(page, cleaned)
        except Exception:
            logging.getLogger(__name__).warning("Не удалось санитизировать файл зон %s", p, exc_info=True)
        return cleaned
    except Exception:
        logging.getLogger(__name__).error("Не удалось прочитать файл зон %s", p, exc_info=True)
        return []



def _save_persisted_zones(page: Any, zones: List[str]) -> None:
    """
    Сохраняет список зон в JSON-файл.
    Храним только НЕпустые зоны, без строки "Без зоны".
    Дубликаты убираем без учёта регистра.
    """
    p = _zones_json_path(page)
    try:
        z = _canonize_list(zones)
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(z, f, ensure_ascii=False, indent=2)
    except Exception as ex:
        logging.getLogger(__name__).error("Не удалось сохранить файл зон %s: %s", p, ex, exc_info=True)

try:
    # Диалог переноса зон. PowerMismatchDialog здесь не используется
    from .dialogs import MoveDialog  # type: ignore
except Exception:  # noqa: B902
    MoveDialog = None  # type: ignore


# 1. Построение вкладки «Сводная смета»
def build_summary_tab(page: Any, tab: QtWidgets.QWidget) -> None:
    """Создаёт интерфейс вкладки «Сводная смета».

    Все виджеты и сигналы регистрируются на объекте ``page``. Реальная
    обработка событий происходит в методах страницы, которые будут
    переназначены в ``ProjectPage`` для вызова соответствующих функций из
    этого модуля.

    :param page: экземпляр ProjectPage, куда помещаются атрибуты и методы
    :param tab: виджет вкладки, который будет заполняться элементами
    """
    v = QtWidgets.QVBoxLayout(tab)

    # 1.1 Фильтры
    filt = QtWidgets.QHBoxLayout()
    filt.setContentsMargins(4, 4, 4, 2)
    filt.setSpacing(8)

    page.ed_search = QtWidgets.QLineEdit()
    page.ed_search.setPlaceholderText("Поиск по наименованию…")

    page.cmb_f_vendor = QtWidgets.QComboBox()
    page.cmb_f_vendor.addItem("<Все подрядчики>")

    page.cmb_f_department = QtWidgets.QComboBox()
    page.cmb_f_department.addItem("<Все отделы>")

    page.cmb_f_class = QtWidgets.QComboBox()
    page.cmb_f_class.addItem("<Все классы>")
    page.cmb_f_class.addItems(list(CLASS_RU2EN.keys()))

    page.btn_delete_selected = QtWidgets.QPushButton("Удалить выделенные")

    for w in (
        QtWidgets.QLabel("Поиск:"), page.ed_search,
        QtWidgets.QLabel("Подрядчик:"), page.cmb_f_vendor,
        QtWidgets.QLabel("Отдел:"), page.cmb_f_department,
        QtWidgets.QLabel("Класс:"), page.cmb_f_class,
        page.btn_delete_selected,
    ):
        filt.addWidget(w)
    # Полоса снимка: кнопка и чекбокс, размещаем справа
    # Кнопка создания снимка, переключатель сравнения и выпадающий список сохранённых снимков
    page.btn_snapshot = QtWidgets.QPushButton("Сохранить снимок")
    page.chk_snapshot_compare = QtWidgets.QCheckBox("Сравнение")
    page.cmb_snapshot = QtWidgets.QComboBox()
    page.cmb_snapshot.setMinimumWidth(160)
    # Объединяем элементы в отдельный фрейм для наглядности
    snap_frame = QtWidgets.QFrame()
    snap_layout = QtWidgets.QHBoxLayout(snap_frame)
    snap_layout.setContentsMargins(4, 2, 4, 2)
    snap_layout.setSpacing(4)
    snap_layout.addWidget(page.btn_snapshot)
    snap_layout.addWidget(page.chk_snapshot_compare)
    snap_layout.addWidget(page.cmb_snapshot)
    snap_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    snap_frame.setStyleSheet("background-color: #444444; border-radius: 4px;")
    filt.addStretch(1)
    filt.addWidget(snap_frame)

    # 1.2 Панель зон
    zbar = QtWidgets.QHBoxLayout()
    zbar.setContentsMargins(4, 2, 4, 2)
    zbar.setSpacing(8)

    page.ed_new_zone = QtWidgets.QLineEdit()
    page.ed_new_zone.setPlaceholderText("Новая зона…")
    page.ed_new_zone.setMinimumWidth(200)

    page.btn_add_zone = QtWidgets.QPushButton("Создать зону")

    # 1.2.0 Кнопка переименования зоны
    page.btn_rename_zone = QtWidgets.QPushButton("Переименовать зону")
    page.btn_delete_zone = QtWidgets.QPushButton("Удалить зону")


    page.cmb_move_zone = QtWidgets.QComboBox()
    page.cmb_move_zone.setEditable(True)
    page.cmb_move_zone.setMinimumWidth(180)

    page.btn_move_zone = QtWidgets.QPushButton("Перенести в зону")

    for w in (
        QtWidgets.QLabel("Зоны:"), page.ed_new_zone, page.btn_add_zone,
        page.btn_rename_zone, page.btn_delete_zone,
        QtWidgets.QLabel("→"), page.cmb_move_zone, page.btn_move_zone
    ):
        zbar.addWidget(w)
    zbar.addStretch(1)

    # 1.2.1 Кнопка «Отменить…»
    page.btn_undo_summary = QtWidgets.QPushButton("Отменить последнее действие")
    page.btn_undo_summary.setEnabled(False)
    zbar.addWidget(page.btn_undo_summary)

    # 1.3 Добавление позиций вручную/из базы
    # Блок добавления помещаем в отдельный фрейм с фоном, чтобы визуально
    # отделить его от остальных элементов интерфейса
    manual_frame = QtWidgets.QFrame()
    manual_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
    # Панель ручного ввода отделяется от остального интерфейса светлым фоном.
    # Темно‑серый фон (#444444) заменён на светло‑серый (#cccccc),
    # чтобы элементы формы были лучше видны на экране и не «сливались»
    # с остальным интерфейсом. Радиус и отступы сохранены.
    # Оттенок фона для панели ручного ввода немного затемнённый относительно
    # системного цвета. Ранее использовался цвет #cccccc (светло‑серый),
    # теперь выбран #bbbbbb для лучшей контрастности.
    manual_frame.setStyleSheet(
        "background-color: #bbbbbb; border-radius: 4px; padding: 4px;"
    )
    manual_layout = QtWidgets.QVBoxLayout(manual_frame)
    manual_layout.setContentsMargins(4, 2, 4, 2)
    manual_layout.setSpacing(4)

    addbar1 = QtWidgets.QHBoxLayout()
    addbar1.setContentsMargins(0, 0, 0, 0)
    addbar1.setSpacing(8)

    # Поле наименования для ручного ввода
    page.ed_add_name = QtWidgets.QLineEdit()
    page.ed_add_name.setPlaceholderText("Наименование")
    page.ed_add_name.setMinimumWidth(260)

    page.sp_add_qty = SmartDoubleSpinBox()
    page.sp_add_qty.setDecimals(3)
    page.sp_add_qty.setMinimum(0.001)
    page.sp_add_qty.setValue(1.000)

    page.sp_add_coeff = SmartDoubleSpinBox()
    page.sp_add_coeff.setDecimals(3)
    page.sp_add_coeff.setMinimum(0.001)
    page.sp_add_coeff.setValue(1.000)

    page.cmb_add_class = QtWidgets.QComboBox()
    page.cmb_add_class.addItems(list(CLASS_RU2EN.keys()))

    # 1.3.1 Добавляем элементы первой строки: наименование, количество, коэффициент и класс
    for w in (
        page.ed_add_name,
        QtWidgets.QLabel("Кол-во:"), page.sp_add_qty,
        QtWidgets.QLabel("Коэф.:",), page.sp_add_coeff,
        QtWidgets.QLabel("Класс:"), page.cmb_add_class
    ):
        addbar1.addWidget(w)
    addbar1.addStretch(1)

    addbar2 = QtWidgets.QHBoxLayout()
    addbar2.setContentsMargins(0, 0, 0, 0)
    addbar2.setSpacing(8)

    # Поля подрядчика и отдела (ручной режим)
    page.ed_add_vendor = QtWidgets.QLineEdit()
    page.ed_add_vendor.setPlaceholderText("Подрядчик")
    page.cmb_add_department = QtWidgets.QComboBox()
    page.cmb_add_department.setEditable(True)
    try:
        le = page.cmb_add_department.lineEdit(); le.setPlaceholderText("Отдел")
    except Exception:
        pass
    # Фильтры подрядчика и отдела (использовались в режиме базы) удалены

    page.cmb_add_zone = QtWidgets.QComboBox()
    page.cmb_add_zone.setMinimumWidth(160)

    page.sp_add_power = SmartDoubleSpinBox()
    page.sp_add_power.setDecimals(0)
    page.sp_add_power.setMaximum(10 ** 9)
    page.sp_add_power.setSuffix(" Вт")

    # Поле «Цена/шт»
    page.sp_add_price = SmartDoubleSpinBox()
    page.sp_add_price.setDecimals(2)
    page.sp_add_price.setMinimum(0.0)
    page.sp_add_price.setMaximum(10 ** 9)
    page.sp_add_price.setValue(0.0)

    page.btn_add_manual = QtWidgets.QPushButton("Добавить позицию")

    # Добавляем элементы второй строки: цена, подрядчик/фильтр, отдел/фильтр, зона, потребление, кнопка
    addbar2.addWidget(QtWidgets.QLabel("Цена/шт:"))
    addbar2.addWidget(page.sp_add_price)
    # Поле подрядчика и отдел для ручного ввода
    addbar2.addWidget(page.ed_add_vendor)
    addbar2.addWidget(QtWidgets.QLabel("Отдел"))
    addbar2.addWidget(page.cmb_add_department)
    # Зона
    addbar2.addWidget(QtWidgets.QLabel("Зона:"))
    addbar2.addWidget(page.cmb_add_zone)
    # Потребление
    addbar2.addWidget(QtWidgets.QLabel("Потр.:",))
    addbar2.addWidget(page.sp_add_power)
    # Кнопка ручного добавления
    addbar2.addWidget(page.btn_add_manual)
    # Кнопка добавления из базы данных
    page.btn_add_from_db = QtWidgets.QPushButton("Из базы…")
    # Задаём красный цвет фона и белый цвет текста для кнопки добавления из базы
    page.btn_add_from_db.setStyleSheet("background-color: #b00020; color: #ffffff;")
    addbar2.addWidget(page.btn_add_from_db)
    addbar2.addStretch(1)

    # 1.3.1 Добавляем оба ряда в фрейм
    manual_layout.addLayout(addbar1)
    manual_layout.addLayout(addbar2)

    # 1.4 Табы зон
    page.zone_tabs = QtWidgets.QTabWidget()
    page.zone_tables: Dict[str, QtWidgets.QTableWidget] = {}

    # 1.5 Компоновка
    v.addLayout(filt)
    v.addLayout(zbar)
    v.addWidget(manual_frame)
    v.addWidget(page.zone_tabs, 1)

    # 1.5a Кнопки мастеров: экран и колонки. Размещаем внизу вкладки.
    # 1.5a Кнопка «Мастер добавления»
    #
    # Внизу вкладки располагается одна кнопка, открывающая мастер добавления.
    # Ранее здесь были отдельные кнопки «Добавить экран» и «Добавить колонки».
    # Теперь они переехали внутрь мастера, который также включает
    # добавление коммутации, сценического подиума и технического директора.
    master_bar = QtWidgets.QHBoxLayout()
    master_bar.setContentsMargins(4, 4, 4, 4)
    master_bar.setSpacing(8)
    master_bar.addStretch(1)
    page.btn_master_add = QtWidgets.QPushButton("Мастер добавления")
    master_bar.addWidget(page.btn_master_add)
    # Создаём скрытые заглушки для устаревших кнопок (экран, редактировать экран, колонки),
    # чтобы сохранить совместимость с кодом, который может обращаться к ним. Эти
    # элементы не добавляются в компоновку и остаются невидимыми.
    page.btn_add_screen = QtWidgets.QPushButton()
    page.btn_add_screen.setVisible(False)
    page.btn_edit_screen = QtWidgets.QPushButton()
    page.btn_edit_screen.setVisible(False)
    page.btn_add_column = QtWidgets.QPushButton()
    page.btn_add_column.setVisible(False)
    v.addLayout(master_bar)

    # 1.6 Сигналы: делегируем на методы ProjectPage
    page.ed_search.textChanged.connect(page._reload_zone_tabs)
    page.cmb_f_vendor.currentTextChanged.connect(page._reload_zone_tabs)
    page.cmb_f_department.currentTextChanged.connect(page._reload_zone_tabs)
    page.cmb_f_class.currentTextChanged.connect(page._reload_zone_tabs)
    page.btn_delete_selected.clicked.connect(page.delete_selected)
    page.btn_add_zone.clicked.connect(page._create_zone)
    page.btn_rename_zone.clicked.connect(lambda: rename_zone(page))
    page.btn_delete_zone.clicked.connect(lambda: delete_zone(page))
    page.btn_move_zone.clicked.connect(page._move_selected_to_zone)
    page.btn_add_manual.clicked.connect(page._add_manual_item)
    page.btn_undo_summary.clicked.connect(page._undo_last_summary)

    # 1.6a Переключение режима базы удалено

    # 1.7 Сигналы для механизма снимков: создание и включение сравнения
    # Кнопка сохраняет снимок с именем
    page.btn_snapshot.clicked.connect(lambda: save_snapshot(page))
    page.chk_snapshot_compare.toggled.connect(lambda state: toggle_snapshot_compare(page))
    page.cmb_snapshot.currentIndexChanged.connect(lambda idx: on_snapshot_selected(page))
    # 1.8 Сигналы поиска удалены, поскольку режим базы заменён отдельным диалогом

    # 1.9 Сигнал кнопки «Из базы…» для открытия диалога выбора позиции
    page.btn_add_from_db.clicked.connect(lambda: show_catalog_dialog(page))

    # 1.10 Сигнал мастера добавления
    #
    # Одна кнопка «Мастер добавления» открывает универсальный диалог,
    # который содержит кнопки для добавления экрана, колонок, коммутации,
    # сценического подиума и технического директора. Отдельные кнопки
    # экрана/колонок больше не используются.
    page.btn_master_add.clicked.connect(lambda: open_master_addition(page))


# 2. Таблица зоны
def build_zone_table(page: Any) -> QtWidgets.QTableWidget:
    """Создаёт пустую таблицу для отображения позиций в зоне."""
    cols = [
        "Наименование", "Кол-во", "Коэф.", "Цена/шт", "Сумма",
        "Подрядчик", "Отдел", "Зона", "Класс", "Потребление (Вт)"
    ]
    t = QtWidgets.QTableWidget(0, len(cols))
    t.setHorizontalHeaderLabels(cols)
    t.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
    t.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
    t.setEditTriggers(
        QtWidgets.QAbstractItemView.DoubleClicked |
        QtWidgets.QAbstractItemView.SelectedClicked
    )
    t.setWordWrap(False)
    # Для наименований используем делегат переноса
    t.setItemDelegateForColumn(0, WrapTextDelegate(t, wrap_threshold=WRAP_THRESHOLD))
    t.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
    setup_priority_name(t, name_col=0)
    return t


# 3. Инициализация табов зон
def init_zone_tabs(page: Any) -> None:
    """Очищает и пересоздаёт вкладки зон в соответствии с данными проекта.

    В зависимости от наличия позиций без зоны (``zone`` равна ``NULL`` или
    пустой строке) создаёт вкладку «Без зоны» или использует первую
    существующую зону в качестве зоны по умолчанию. Также вычисляет
    ``page.default_zone`` для использования при добавлении новых позиций.
    """
    # Удаляем старые вкладки
    while page.zone_tabs.count() > 0:
        w = page.zone_tabs.widget(0)
        page.zone_tabs.removeTab(0)
        w.deleteLater()
    page.zone_tables.clear()

    # Сбрасываем значение зоны по умолчанию
    page.default_zone = ""

    if page.project_id is None:
        return

    # 3.1 Получаем список зон из БД
    try:
        zones: List[str] = page.db.project_distinct_values(page.project_id, "zone") or []
    except Exception:
        zones = []
    # 3.2 Подмешиваем сохранённые зоны, чтобы отображать вкладки без позиций
    try:
        zones += _load_persisted_zones(page)
    except Exception:
        pass
    # 3.3 Определяем наличие позиций без зоны (``None`` или пустая строка)
    no_zone_exists = False
    for z in zones:
        if z is None or str(z).strip() == "":
            no_zone_exists = True
            break
    # 3.4 Формируем список уникальных зон без учёта регистра и пустых значений
    unique_zones: List[str] = []
    seen_lower: Set[str] = set()
    for z in zones:
        if not z:
            continue
        norm = str(z).strip()
        key = norm.lower()
        if key not in seen_lower:
            unique_zones.append(norm)
            seen_lower.add(key)
    # 3.5 Если есть позиции без зоны, добавляем пустую зону; иначе нет
    zones_clean: List[str] = []
    if no_zone_exists:
        zones_clean.append("")
        # зона без имени остаётся зоной по умолчанию
        page.default_zone = ""
    else:
        # иначе зоной по умолчанию станет первая существующая
        if unique_zones:
            page.default_zone = unique_zones[0]
    zones_clean.extend(unique_zones)

    # 3.6 Создаём вкладки и таблицы для каждой зоны
    for z in zones_clean:
        # метка для вкладки
        label = "Без зоны" if not z else z
        table = build_zone_table(page)
        table.itemChanged.connect(page._on_summary_item_changed)
        # Разрешаем пользовательское контекстное меню для группирования
        table.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        # Передаём ключ зоны через lambda, чтобы обработать меню для нужной таблицы
        table.customContextMenuRequested.connect(lambda pos, z_key=z: on_zone_table_context_menu(page, z_key, pos))
        page.zone_tabs.addTab(table, label)
        page.zone_tables[z] = table

    # 3.7 Обновляем список зон для переноса
    page.cmb_move_zone.blockSignals(True)
    page.cmb_move_zone.clear()
    if no_zone_exists:
        page.cmb_move_zone.addItem("Без зоны", "")
    for z in unique_zones:
        # для комбобокса используем нормализованное отображение
        page.cmb_move_zone.addItem(normalize_case(z), z)
    page.cmb_move_zone.blockSignals(False)

    # 3.8 Обновляем список зон для ручного добавления
    fill_manual_zone_combo(page, unique_zones)


# 4. Комбо зон для ручного добавления
def fill_manual_zone_combo(page: Any, zones: List[str]) -> None:
    """Обновляет выпадающий список зон на панели ручного добавления.

    Если ``page.default_zone`` равна пустой строке, то показываем пункт
    «Без зоны» в качестве первого варианта. В противном случае выводим
    только существующие зоны, чтобы пользователь по умолчанию добавлял
    позиции в первую зону из списка.
    """
    page.cmb_add_zone.blockSignals(True)
    page.cmb_add_zone.clear()
    # Включаем вариант "Без зоны" только если зона по умолчанию действительно пустая
    if getattr(page, "default_zone", "") == "":
        page.cmb_add_zone.addItem("Без зоны", "")
    # Добавляем остальные зоны
    for z in zones:
        if z:
            # Показываем нормализованный вариант, но храним исходный ключ
            page.cmb_add_zone.addItem(normalize_case(z), z)
    page.cmb_add_zone.blockSignals(False)


# 4.a Контекстное меню таблицы зон: группирование и разъединение
def on_zone_table_context_menu(page: Any, zone_key: str, pos: QtCore.QPoint) -> None:
    """
    Обрабатывает вызов контекстного меню для таблицы зоны.

    Позволяет пользователю объединить выбранные позиции в группу или
    разъединить существующую группу. Группы применяются только в рамках
    сводной сметы и не влияют на расчёты. Для группировки применяется
    поле ``group_name`` в таблице ``items``. После изменения группы
    таблицы перезагружаются.

    :param page: Экземпляр ProjectPage
    :param zone_key: Ключ зоны (пустая строка для «Без зоны»)
    :param pos: Координаты клика, переданные сигналом
    """
    try:
        table = page.zone_tables.get(zone_key)
    except Exception:
        table = None
    if not isinstance(table, QtWidgets.QTableWidget):
        return
    # Определяем выбранные строки
    selected_indexes = sorted({idx.row() for idx in table.selectedIndexes()})
    if not selected_indexes:
        return
    # Создаём меню
    menu = QtWidgets.QMenu(table)
    act_group = menu.addAction("Собрать в группу")
    act_ungroup = menu.addAction("Разъединить группу")
    action = menu.exec(table.viewport().mapToGlobal(pos))
    if action == act_group:
        group_selected_items(page, zone_key)
    elif action == act_ungroup:
        ungroup_selected_items(page, zone_key)


def group_selected_items(page: Any, zone_key: str) -> None:
    """
    Объединяет выделенные строки таблицы зоны в одну группу.

    У пользователя запрашивается имя группы. При отмене диалога
    никаких изменений не производится. После назначения group_name
    перезагружаются вкладки зон и выводится сообщение в лог.

    :param page: Экземпляр ProjectPage
    :param zone_key: Ключ зоны, в которой выполняется группирование
    """
    table = page.zone_tables.get(zone_key)
    if not isinstance(table, QtWidgets.QTableWidget):
        return
    selected_rows = sorted({idx.row() for idx in table.selectedIndexes()})
    if not selected_rows:
        return
    # Запрашиваем имя группы у пользователя. Предлагаем вариант по умолчанию
    # в формате «Группа n».
    default_name = ""
    try:
        # Если все выделенные элементы уже принадлежат одной группе, предлагаем её имя
        group_names: set[str] = set()
        for r in selected_rows:
            itm = table.item(r, 0)
            if itm is None:
                continue
            item_id_data = itm.data(QtCore.Qt.ItemDataRole.UserRole)
            try:
                item_id = int(item_id_data)
            except Exception:
                continue
            row_db = page.db.get_item_by_id(item_id)
            if row_db is not None:
                g = normalize_case(row_db["group_name"] or "")
                if g:
                    group_names.add(g)
        if len(group_names) == 1:
            default_name = next(iter(group_names))
    except Exception:
        default_name = ""
    name, ok = QtWidgets.QInputDialog.getText(
        page,
        "Создание группы",
        "Введите название группы:",
        text=default_name or "Группа"
    )
    if not ok:
        return
    group_name = normalize_case(name.strip())
    if not group_name:
        return
    updated_count = 0
    for r in selected_rows:
        itm = table.item(r, 0)
        if itm is None:
            continue
        item_id_data = itm.data(QtCore.Qt.ItemDataRole.UserRole)
        try:
            item_id = int(item_id_data)
        except Exception:
            continue
        try:
            page.db.update_item_fields(item_id, {"group_name": group_name})
            updated_count += 1
        except Exception:
            continue
    # Логируем и перезагружаем данные
    try:
        if updated_count > 0 and hasattr(page, "_log"):
            page._log(f"Группирование: {updated_count} позиций объединены в группу «{group_name}».")
    except Exception:
        pass
    try:
        page._reload_zone_tabs()
    except Exception:
        pass


def ungroup_selected_items(page: Any, zone_key: str) -> None:
    """
    Разъединяет выделенные позиции, устанавливая для них уникальный group_name.

    Для каждой выбранной строки group_name обновляется до её собственного
    наименования позиции, что приводит к тому, что позиции больше не
    связываются в одну группу. После обновления данные перезагружаются.

    :param page: Экземпляр ProjectPage
    :param zone_key: Ключ зоны, в которой выполняется разбиение
    """
    table = page.zone_tables.get(zone_key)
    if not isinstance(table, QtWidgets.QTableWidget):
        return
    selected_rows = sorted({idx.row() for idx in table.selectedIndexes()})
    if not selected_rows:
        return
    updated_count = 0
    for r in selected_rows:
        itm = table.item(r, 0)
        if itm is None:
            continue
        item_id_data = itm.data(QtCore.Qt.ItemDataRole.UserRole)
        try:
            item_id = int(item_id_data)
        except Exception:
            continue
        # Получаем наименование для новой группы
        try:
            row_db = page.db.get_item_by_id(item_id)
        except Exception:
            row_db = None
        new_group = ""
        try:
            if row_db is not None:
                nm = normalize_case(row_db["name"] or "")
                # Чтобы избежать повторного группирования нескольких одинаковых элементов,
                # используем уникальный идентификатор в имени группы
                new_group = f"{nm} #{item_id}"
        except Exception:
            new_group = ""
        # Обновляем group_name
        try:
            page.db.update_item_fields(item_id, {"group_name": new_group})
            updated_count += 1
        except Exception:
            continue
    try:
        if updated_count > 0 and hasattr(page, "_log"):
            page._log(f"Разъединение: {updated_count} позиций разъединены на отдельные группы.")
    except Exception:
        pass
    try:
        page._reload_zone_tabs()
    except Exception:
        pass


# 4.1 Комбо отделов для ручного добавления
def fill_manual_dept_combo(page: Any, departments: List[str]) -> None:
    """Обновляет выпадающий список отделов на панели ручного добавления.

    Пользователь может выбрать отдел из списка существующих отделов проекта
    или ввести свой. Для сохранения текущего текста комбобокс запоминает
    выбранное значение, очищает список, снова заполняет его и восстанавливает
    введённый текст.
    """
    # Текущий текст до перезаполнения
    current = page.cmb_add_department.currentText() if hasattr(page, 'cmb_add_department') else ""
    page.cmb_add_department.blockSignals(True)
    page.cmb_add_department.clear()
    page.cmb_add_department.setEditable(True)
    # Заполняем список существующих отделов
    for d in departments:
        if d:
            page.cmb_add_department.addItem(d)
    # Восстанавливаем введённый текст
    page.cmb_add_department.setEditText(current)
    page.cmb_add_department.blockSignals(False)


# 5. Обновление таблиц по фильтрам
def reload_zone_tabs(page: Any) -> None:
    """Перестраивает таблицы зон согласно текущим фильтрам.

    Если активирован режим сравнения со снимком (page._snapshot_compare_enabled),
    таблицы расширяются дополнительными колонками и заполняются с учётом
    изменений относительно сохранённого состояния.
    """
    if page.project_id is None:
        return
    # При смене проекта обновляем список сохранённых снимков
    if getattr(page, "_snapshots_loaded_project_id", None) != page.project_id:
        load_snapshot_list(page)
        page._snapshots_loaded_project_id = page.project_id
    # При первом вызове инициализируем вкладки зон
    if not page.zone_tables:
        init_zone_tabs(page)

    # Получаем текущие фильтры
    # Текст поиска (сырой) для дальнейшей фильтрации по канону
    search_raw = page.ed_search.text() if hasattr(page, "ed_search") else ""
    vendor = page.cmb_f_vendor.currentText()
    if vendor == "<Все подрядчики>":
        vendor = "<ALL>"
    department = page.cmb_f_department.currentText()
    if department == "<Все отделы>":
        department = "<ALL>"
    class_ru = page.cmb_f_class.currentText()
    class_en = CLASS_RU2EN.get(class_ru) if class_ru and class_ru != "<Все классы>" else "<ALL>"

    # Общая сумма по всем зонам (не отображается пользователю),
    # мы будем отдельно рассчитывать сумму для активной зоны.
    total_amount = 0.0

    # Проверяем, активирован ли режим сравнения
    snap_mode = getattr(page, "_snapshot_compare_enabled", False) and hasattr(page, "_snapshot_data")

    for zone_key, table in page.zone_tables.items():
        # Получаем строки по фильтрам для текущей зоны. Поиск по наименованию
        # выполняем позже в Python, чтобы игнорировать регистр и учитывать хомоглифы.
        rows = page.db.list_items_filtered(
            project_id=page.project_id,
            vendor=vendor,
            department=department,
            zone=zone_key,
            class_en=class_en,
            name_like=None,
        )
        # Фильтруем по строке поиска вручную, если указано non-empty
        if search_raw:
            filtered: List[Any] = []
            for r in rows:
                try:
                    # r — sqlite3.Row, доступ по ключу
                    nm = r["name"] if r["name"] is not None else ""
                    ven = r["vendor"] if r["vendor"] is not None else ""
                    dep = r["department"] if r["department"] is not None else ""
                    zn = r["zone"] if r["zone"] is not None else ""
                    if (
                        contains_search(nm, search_raw)
                        or contains_search(ven, search_raw)
                        or contains_search(dep, search_raw)
                        or contains_search(zn, search_raw)
                    ):
                        filtered.append(r)
                except Exception:
                    continue
            rows = filtered
            try:
                page._log(f"Сводная смета: поиск '{search_raw}' отфильтровал {len(rows)} элементов (зона {zone_key}).")
            except Exception:
                pass

        table.blockSignals(True)
        table.setRowCount(0)

        if snap_mode:
            # В режиме сравнения расширяем таблицу и заголовки
            headers = [
                "Наименование", "Состояние", "Кол-во", "ΔКол-во", "Коэф.",
                "Цена/шт", "ΔЦена", "Сумма", "ΔСумма",
                "Подрядчик", "Отдел", "Зона", "Класс", "Потребление (Вт)"
            ]
            if table.columnCount() != len(headers):
                table.setColumnCount(len(headers))
                table.setHorizontalHeaderLabels(headers)
            else:
                table.setHorizontalHeaderLabels(headers)
        else:
            # Стандартный набор столбцов
            headers = [
                "Наименование", "Кол-во", "Коэф.", "Цена/шт", "Сумма",
                "Подрядчик", "Отдел", "Зона", "Класс", "Потребление (Вт)"
            ]
            if table.columnCount() != len(headers):
                table.setColumnCount(len(headers))
                table.setHorizontalHeaderLabels(headers)
            else:
                table.setHorizontalHeaderLabels(headers)

        if snap_mode:
            # Словарь snapshot для быстрого доступа
            snap_items = page._snapshot_data.get("items", {})
            used_snapshot_ids = set()
            # Заполняем текущие строки со сравнением
            for r in rows:
                i = table.rowCount()
                table.insertRow(i)
                item_id = int(r["id"])
                snap = snap_items.get(item_id)
                # Текущие значения
                cur_qty = float(r["qty"] or 0)
                cur_coeff = float(r["coeff"] or 0)
                cur_price = float(r["unit_price"] or 0)
                cur_amount = float(r["amount"] or 0)
                cur_class = CLASS_EN2RU.get((r["type"] or "equipment"), "Оборудование")
                # Разница по умолчанию для новых позиций
                diff_qty = cur_qty
                diff_price = cur_price
                diff_amount = cur_amount
                state = "добавлено"
                # Цвет текста для добавленных строк (тёмно‑зелёный)
                color = QtGui.QColor(0, 150, 0)
                if snap:
                    used_snapshot_ids.add(item_id)
                    # Расчёт разницы
                    diff_qty = cur_qty - snap["qty"]
                    diff_price = cur_price - snap["unit_price"]
                    diff_amount = cur_amount - snap["amount"]
                    # Определяем состояние и цвет
                    if abs(diff_qty) < 1e-6 and abs(diff_price) < 1e-6:
                        state = "не изменилось"
                        color = QtGui.QColor(0, 0, 0, 0)  # прозрачный
                    else:
                        state = "изменилось"
                        # Если изменилась и цена и количество — жёлтый
                        if abs(diff_qty) >= 1e-6 and abs(diff_price) >= 1e-6:
                            # Оранжевый текст для изменения и цены, и количества
                            color = QtGui.QColor(255, 165, 0)
                        # Если изменилась только цена: красный для увеличения, зелёный для уменьшения
                        elif abs(diff_price) >= 1e-6:
                            color = QtGui.QColor(200, 0, 0) if diff_price > 0 else QtGui.QColor(0, 150, 0)
                        # Если изменилось только количество: красный для увеличения, зелёный для уменьшения
                        elif abs(diff_qty) >= 1e-6:
                            color = QtGui.QColor(200, 0, 0) if diff_qty > 0 else QtGui.QColor(0, 150, 0)
                # Формируем значения строк
                # Нормализуем регистр отображаемых полей (наименование, подрядчик, отдел, зона)
                # Имя всегда присутствует в строке sqlite3.Row, используем значение или пустую строку
                name_norm = normalize_case(r["name"] or "")
                vendor_norm = normalize_case(r["vendor"] or "")
                department_norm = normalize_case(r["department"] or "")
                zone_norm = normalize_case(r["zone"] or "")
                # При сравнении отображаем имя группы, если оно задано и отличается от имени позиции.
                try:
                    gname_s = normalize_case(r.get("group_name", ""))
                except Exception:
                    gname_s = ""
                display_name_snap = name_norm
                if gname_s and gname_s not in ("", "аренда оборудования", name_norm.lower()):
                    display_name_snap = f"{normalize_case(gname_s)}: {name_norm}"
                vals = [
                    display_name_snap,
                    state,
                    fmt_num(cur_qty, 3),
                    fmt_sign(diff_qty, 3),
                    fmt_num(cur_coeff, 3),
                    fmt_num(cur_price, 2),
                    fmt_sign(diff_price, 2),
                    fmt_num(cur_amount, 2),
                    fmt_sign(diff_amount, 2),
                    vendor_norm,
                    department_norm,
                    zone_norm,
                    cur_class,
                    fmt_num(float(r["power_watts"] or 0), 0),
                ]
                for c, v in enumerate(vals):
                    item = QtWidgets.QTableWidgetItem(str(v))
                    if c == 0:
                        item.setData(QtCore.Qt.UserRole, item_id)
                    # Разрешаем редактирование только для кол-ва, коэффициента и цены (столбцы 2, 4, 5)
                    if c not in (2, 4, 5):
                        item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                    # Если задан цвет — применяем его к цвету текста, а не к фону
                    if color.alpha() > 0:
                        item.setForeground(QtGui.QBrush(color))
                    table.setItem(i, c, item)
                total_amount += cur_amount
            # Добавляем удалённые строки
            for sid, snap in snap_items.items():
                if snap.get("zone", "") != zone_key:
                    continue
                if sid in used_snapshot_ids:
                    continue
                i = table.rowCount()
                table.insertRow(i)
                # Значения из снимка
                snap_qty = float(snap.get("qty", 0))
                snap_price = float(snap.get("unit_price", 0))
                snap_amount = float(snap.get("amount", 0))
                cur_class = CLASS_EN2RU.get((snap.get("class") or "equipment"), "Оборудование")
                diff_qty = -snap_qty
                diff_price = -snap_price
                diff_amount = -snap_amount
                # Нормализуем поля из снимка для отображения
                name_norm = normalize_case(snap.get("name", ""))
                vendor_norm = normalize_case(snap.get("vendor", ""))
                department_norm = normalize_case(snap.get("department", ""))
                zone_norm = normalize_case(snap.get("zone", ""))
                vals = [
                    name_norm,
                    "удалено",
                    fmt_num(0, 3),
                    fmt_sign(diff_qty, 3),
                    fmt_num(float(snap.get("coeff", 0)), 3),
                    fmt_num(0, 2),
                    fmt_sign(diff_price, 2),
                    fmt_num(0, 2),
                    fmt_sign(diff_amount, 2),
                    vendor_norm,
                    department_norm,
                    zone_norm,
                    cur_class,
                    fmt_num(float(snap.get("power_watts", 0)), 0),
                ]
                for c, v in enumerate(vals):
                    item = QtWidgets.QTableWidgetItem(str(v))
                    # Не позволяем редактировать удалённые записи
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                    # Для удалённых строк используем тёмно‑красный цвет текста для выделения
                    item.setForeground(QtGui.QBrush(QtGui.QColor(200, 0, 0)))
                    table.setItem(i, c, item)
        else:
            # 5.3 Агрегация дубликатов (без режима сравнения)
            # Для стандартного режима без сравнения объединяем записи с одинаковыми
            # наименованием, подрядчиком и зоной (и отделом) внутри одной зоны. Это позволяет
            # суммировать позиции, добавленные вручную или импортом, если они относятся
            # к одному подрядчику и имеют одно название. Суммируются количество и сумма,
            # коэффициент и цена рассчитываются как взвешенные значения.
            agg: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
            for r in rows:
                # Нормализуем ключевые поля для устойчивого сравнения
                name_norm = normalize_case(r["name"] or "")
                vendor_norm = normalize_case(r["vendor"] or "")
                zone_norm = normalize_case(r["zone"] or "")
                department_norm = normalize_case(r["department"] or "")
                # Нормализуем имя группы. Пустая строка или 'Аренда оборудования'
                # считаются отсутствием группы, чтобы не путать с произвольными названиями.
                group_name_norm = normalize_case(r["group_name"] or "")
                # Конструируем ключ агрегирования так, чтобы позиции разных групп не
                # объединялись. Это позволяет отображать позиции по группам. Если
                # группы нет, используем пустую строку, чтобы поведение осталось прежним.
                key = (name_norm, vendor_norm, zone_norm, department_norm, group_name_norm)
                try:
                    qty = float(r["qty"] or 0)
                except Exception:
                    qty = 0.0
                try:
                    coeff = float(r["coeff"] or 0)
                except Exception:
                    coeff = 0.0
                try:
                    amount = float(r["amount"] or 0)
                except Exception:
                    amount = 0.0
                try:
                    unit_price = float(r["unit_price"] or 0)
                except Exception:
                    unit_price = 0.0
                try:
                    power = float(r["power_watts"] or 0)
                except Exception:
                    power = 0.0
                if key not in agg:
                    # Инициализируем запись агрегата. Сохраняем group_name для отображения.
                    agg[key] = {
                        "id": r["id"],
                        "name": name_norm,
                        "vendor": vendor_norm,
                        "department": department_norm,
                        "zone": zone_norm,
                        "group_name": group_name_norm,
                        "type": r["type"] or "equipment",
                        "qty_sum": qty,
                        "coeff_sum": coeff * qty,
                        "coeff_values": {coeff},
                        "amount_sum": amount,
                        "power_sum": power * qty,
                        # Сумма (цена * количество) для вычисления средней цены без учёта коэффициента
                        "price_qty_sum": unit_price * qty,
                    }
                else:
                    rec = agg[key]
                    rec["qty_sum"] += qty
                    rec["coeff_sum"] += coeff * qty
                    rec["amount_sum"] += amount
                    rec["power_sum"] += power * qty
                    rec["price_qty_sum"] += unit_price * qty
                    # Сохраняем все уникальные коэффициенты
                    rec.setdefault("coeff_values", set()).add(coeff)
            # Преобразуем агрегированные данные в строки таблицы
            for rec in agg.values():
                i = table.rowCount()
                table.insertRow(i)
                qty_total = rec["qty_sum"]
                coeff_val: float = 0.0
                # Если у позиции несколько различных коэффициентов, спрашиваем пользователя, какой выбрать
                coeff_values: Set[float] = rec.get("coeff_values", set())
                if qty_total > 0 and len(coeff_values) > 1:
                    # Подготавливаем варианты для выбора, сортируем и форматируем
                    coeff_options = sorted({round(v, 6) for v in coeff_values})
                    opt_strings = [fmt_num(c, 3) for c in coeff_options]
                    try:
                        choice, ok = QtWidgets.QInputDialog.getItem(
                            page,
                            "Выбор коэффициента",
                            f"Для позиции '{rec['name']}' (подрядчик: {rec['vendor']}, зона: {rec['zone']}) "
                            f"обнаружены разные коэффициенты. Выберите нужный коэффициент:",
                            opt_strings,
                            0,
                            False,
                        )
                        if ok and choice:
                            coeff_val = float(choice.replace(" ", "").replace(",", "."))
                        else:
                            # Отмена или ошибка: используем средневзвешенное значение
                            coeff_val = rec["coeff_sum"] / qty_total if qty_total > 0 else 0.0
                    except Exception:
                        coeff_val = rec["coeff_sum"] / qty_total if qty_total > 0 else 0.0
                else:
                    # Если коэффициент один, используем его
                    if qty_total > 0:
                        if coeff_values:
                            # берем единственный коэффициент
                            coeff_val = list(coeff_values)[0]
                        else:
                            coeff_val = rec["coeff_sum"] / qty_total
                # Вычисляем цену за единицу. Цена рассчитывается как средневзвешенная по количеству
                # без учёта коэффициента: price_avg = sum(price_i * qty_i) / sum(qty_i)
                unit_price_val = 0.0
                if qty_total > 0:
                    # Если сумма price_qty_sum определена, используем её, иначе fallback к amount_sum/qty_total/coeff
                    pq_sum = rec.get("price_qty_sum")
                    if pq_sum is not None:
                        unit_price_val = pq_sum / qty_total
                    elif coeff_val > 0:
                        unit_price_val = rec["amount_sum"] / (qty_total * coeff_val)
                # Среднее значение мощности (Вт) на единицу
                power_avg = 0.0
                if qty_total > 0:
                    power_avg = rec["power_sum"] / qty_total
                # Русское отображение класса
                class_ru_show = CLASS_EN2RU.get((rec.get("type") or "equipment"), "Оборудование")
                # Для отображения имени учитываем группу: если имя группы задано и оно не
                # совпадает с именем позиции, добавляем его в начале. Это позволяет
                # визуально отделять элементы, принадлежащие одной группе.
                display_name = rec["name"]
                try:
                    gname = rec.get("group_name", "")
                except Exception:
                    gname = ""
                if gname and gname not in ("", "аренда оборудования", display_name.lower()):
                    # отображаем нормализованный вариант группы для читаемости
                    display_name = f"{normalize_case(gname)}: {display_name}"
                # Формируем список отображаемых значений
                vals = [
                    display_name,
                    fmt_num(qty_total, 3),
                    fmt_num(coeff_val, 3),
                    fmt_num(unit_price_val, 2),
                    fmt_num(rec["amount_sum"], 2),
                    rec["vendor"],
                    rec["department"],
                    rec["zone"],
                    class_ru_show,
                    fmt_num(power_avg, 0),
                ]
                for c, v in enumerate(vals):
                    item = QtWidgets.QTableWidgetItem(str(v))
                    # Сохраняем id для первой колонки
                    if c == 0:
                        item.setData(QtCore.Qt.UserRole, int(rec["id"]))
                    # Разрешаем редактирование только для количества, коэффициента и цены (столбцы 1,2,3)
                    if c not in (1, 2, 3):
                        item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                    table.setItem(i, c, item)
                total_amount += rec["amount_sum"]
        table.blockSignals(False)
        apply_auto_col_resize(table)
    # Сумма зависит не только от фильтров, но и от выбранной зоны.
    # Определяем активную зону: индекс вкладки и соответствующий ключ.
    try:
        cur_idx = page.zone_tabs.currentIndex() if hasattr(page, "zone_tabs") else -1
    except Exception:
        cur_idx = -1
    cur_sum = 0.0
    if cur_idx >= 0:
        # Получаем ключ зоны по порядку из zone_tables (порядок соответствует вкладкам)
        try:
            keys = list(page.zone_tables.keys())
            if 0 <= cur_idx < len(keys):
                cur_zone_key = keys[cur_idx]
                # Получаем те же фильтры, что и в начале функции
                vendor_f = page.cmb_f_vendor.currentText()
                if vendor_f == "<Все подрядчики>":
                    vendor_f = "<ALL>"
                department_f = page.cmb_f_department.currentText()
                if department_f == "<Все отделы>":
                    department_f = "<ALL>"
                class_ru_f = page.cmb_f_class.currentText()
                class_en_f = CLASS_RU2EN.get(class_ru_f) if class_ru_f and class_ru_f != "<Все классы>" else "<ALL>"
                # Запрашиваем позиции только для выбранной зоны и фильтров
                rows_cur = page.db.list_items_filtered(
                    project_id=page.project_id,
                    vendor=vendor_f,
                    department=department_f,
                    zone=cur_zone_key,
                    class_en=class_en_f,
                    name_like=None,
                )
                # Если есть строка поиска, применяем её
                search_raw_f = page.ed_search.text() if hasattr(page, "ed_search") else ""
                if search_raw_f:
                    filtered_cur = []
                    for r in rows_cur:
                        try:
                            nm = r["name"] or ""
                            ven = r["vendor"] or ""
                            dep = r["department"] or ""
                            zn = r["zone"] or ""
                            if (
                                contains_search(nm, search_raw_f)
                                or contains_search(ven, search_raw_f)
                                or contains_search(dep, search_raw_f)
                                or contains_search(zn, search_raw_f)
                            ):
                                filtered_cur.append(r)
                        except Exception:
                            continue
                    rows_cur = filtered_cur
                # Суммируем amount
                for r in rows_cur:
                    try:
                        amt = float(r["amount"] or 0.0)
                        cur_sum += amt
                    except Exception:
                        continue
        except Exception:
            cur_sum = 0.0
    # Устанавливаем сумму выбранной зоны
    page.label_total.setText(f"Итого: {fmt_num(cur_sum, 2)}")
    # Обновляем данные бухгалтерии, если доступна
    try:
        if hasattr(page, "recalc_finance"):
            page.recalc_finance()
    except Exception:
        pass


# 6. Создание новой зоны
def create_zone(page: Any) -> None:
    """Создаёт новую зону, добавляя вкладку и обновляя списки."""
    # Нормализуем название зоны
    name = normalize_case(page.ed_new_zone.text())
    if not name:
        QtWidgets.QMessageBox.information(page, "Внимание", "Введите название зоны.")
        return
    # Проверяем дубликат с учётом нормализованного регистра
    if name in page.zone_tables:
        QtWidgets.QMessageBox.information(page, "Инфо", "Такая зона уже есть.")
        return

    table = build_zone_table(page)
    page.zone_tables[name] = table
    page.zone_tabs.addTab(table, name)
    page.cmb_move_zone.addItem(name, name)

    zones = page.db.project_distinct_values(page.project_id, "zone") if page.project_id else []
    fill_manual_zone_combo(page, list(zones) + [name])

    page._log(f"Создана новая зона: «{name}»")
    page.ed_new_zone.clear()

    # 6.a Сохраняем зону в файл, чтобы она оставалась после перезапуска
    try:
        persisted = _load_persisted_zones(page)
        persisted.append(name)
        _save_persisted_zones(page, persisted)
    except Exception:
        # Логируем ошибку сохранения зон
        logging.getLogger(__name__).error("Не удалось сохранить новую зону", exc_info=True)


# 6. Переименование зоны
def rename_zone(page: Any) -> None:
    """Переименовывает текущую зону.

    Запрашивает у пользователя новое имя для выбранной вкладки зоны. После
    проверки на валидность имя обновляется в базе данных (через метод
    ``DB.rename_zone``) и в пользовательском интерфейсе: переименовывается
    вкладка, ключ в ``page.zone_tables`` и выпадающие списки. Также
    обновляется файл персистентности зон, чтобы новое имя сохранялось
    между сеансами.

    Логика работы:

    * Определяем текущую зону по выбранной вкладке. Если имя пустое,
      интерпретируем его как "Без зоны" (ключ ``""``).
    * Запрашиваем новое имя с помощью ``QInputDialog.getText``.
    * Нормализуем регистр и убираем пробелы с помощью ``normalize_case``.
    * Проверяем, что имя не пустое и не дублирует существующую зону.
    * Вызываем ``page.db.rename_zone`` для обновления таблицы ``items``.
    * Обновляем структуры ``page.zone_tables``, ``page.zone_tabs`` и
      выпадающие списки ``cmb_move_zone`` и ``cmb_add_zone``.
    * Обновляем JSON-файл с зонами: удаляем старое имя и добавляем новое.
    * Логируем операцию через ``page._log``.

    При ошибках записи в базу или работе с файлами ошибка выводится в
    лог приложения, а пользователю показывается информационное сообщение.
    """
    # Получаем индекс и имя текущей вкладки
    cur_index = page.zone_tabs.currentIndex()
    if cur_index < 0:
        return
    old_label = page.zone_tabs.tabText(cur_index)
    # 6.1.1 Приводим имя текущей зоны к каноничному виду.
    # строка "Без зоны" и пустые/None значения считаются пустой зоной ("").
    old_zone = _canon_zone(old_label)
    # Логируем начало операции для диагностики.
    logging.getLogger(__name__).info(
        "Начато переименование зоны: текущая=%s (canon=%s)",
        old_label or "<Без зоны>", old_zone or "<без зоны>",
    )
    # Запрашиваем новое имя
    new_name, ok = QtWidgets.QInputDialog.getText(
        page,
        "Переименование зоны",
        f"Введите новое имя для зоны «{old_label}»:",
        text=old_label,
    )
    if not ok:
        return
    # Нормализуем введённый текст
    new_name_norm = normalize_case(new_name)
    # Приводим новое имя к каноничному виду. Пустая строка или "Без зоны"
    # интерпретируется как зона без имени.
    new_zone = _canon_zone(new_name_norm)
    # Если имя не изменилось — ничего не делаем
    if new_zone == old_zone:
        logging.getLogger(__name__).info(
            "Переименование зоны отменено: новое имя совпадает с текущим."
        )
        return
    # Проверяем дубликат (с учётом регистра) среди уже созданных зон.
    existing_lower = {str(k).lower() for k in page.zone_tables.keys()}
    if new_zone.lower() in existing_lower:
        QtWidgets.QMessageBox.information(
            page, "Информация", "Зона с таким названием уже существует."
        )
        logging.getLogger(__name__).info(
            "Переименование зоны отменено: дубликат %s", new_zone
        )
        return
    # Переименовываем в БД
    try:
        if page.project_id is not None:
            # old_zone или new_zone могут быть пустой строкой, что означает NULL
            page.db.rename_zone(
                page.project_id, old_zone or None, new_zone or None
            )
    except Exception as ex:
        # Логируем и выводим информацию
        logging.getLogger(__name__).error(
            "Ошибка обновления зоны в базе данных: %s", ex, exc_info=True
        )
        QtWidgets.QMessageBox.critical(
            page, "Ошибка", f"Не удалось переименовать зону в базе: {ex}"
        )
        return
    # Обновляем структуру zone_tables: переносим таблицу на новый ключ
    table = page.zone_tables.get(old_zone)
    if table is None:
        return
    # Удаляем старый ключ и присваиваем новый каноничный ключ.
    page.zone_tables.pop(old_zone, None)
    page.zone_tables[new_zone] = table
    # Обновляем название вкладки: для пустой зоны отображаем "Без зоны",
    # иначе используем нормализованную форму для отображения.
    new_label = "Без зоны" if new_zone == "" else normalize_case(new_name_norm)
    page.zone_tabs.setTabText(cur_index, new_label)
    # Обновляем выпадающий список зон для переноса
    page.cmb_move_zone.blockSignals(True)
    page.cmb_move_zone.clear()
    page.cmb_move_zone.addItem("Без зоны", "")
    for z_key in page.zone_tables.keys():
        if z_key:
            # Отображаем пользователю нормализованное название,
            # но сохраняем канон как данные.
            page.cmb_move_zone.addItem(normalize_case(z_key), z_key)
    page.cmb_move_zone.blockSignals(False)
    # Обновляем комбобокс для ручного добавления
    # Получаем зоны из БД (не учитывая пустую строку)
    zones_db: List[str] = []
    try:
        if page.project_id is not None:
            zones_db = page.db.project_distinct_values(page.project_id, "zone") or []
    except Exception:
        pass
    # Загружаем текущие сохранённые зоны
    try:
        persisted = _load_persisted_zones(page)
    except Exception:
        persisted = []
    # Обновляем список: удаляем старую зону и добавляем новую
    updated_persisted: List[str] = []
    for z in persisted:
        zn = z.strip()
        if not _is_no_zone(zn) and zn.strip().lower() != old_zone.lower():
            updated_persisted.append(z)
    if new_zone:
        updated_persisted.append(new_zone)
    try:
        _save_persisted_zones(page, updated_persisted)
    except Exception:
        logging.getLogger(__name__).error(
            "Не удалось обновить файл зон при переименовании", exc_info=True
        )
    # Формируем список для заполнения комбобокса ручного добавления
    combined_zones: List[str] = []
    # Добавляем зоны из БД и из persist (без пустой строки) с учётом регистра
    for z in zones_db:
        c = _canon_zone(z)
        if c:
            disp = normalize_case(c)
            if disp not in combined_zones:
                combined_zones.append(disp)
    for z in updated_persisted:
        c = _canon_zone(z)
        if c:
            disp = normalize_case(c)
            if disp not in combined_zones:
                combined_zones.append(disp)
    fill_manual_zone_combo(page, combined_zones)
    # Логируем переименование
    try:
        page._log(f"Зона «{old_label}» переименована в «{new_label}»")
    except Exception:
        pass
    logging.getLogger(__name__).info(
        "Зона '%s' переименована в '%s'",
        old_label or "<Без зоны>", new_label,
    )
    # Перезагружаем таблицы, чтобы обновить данные в колонке зоны
    try:
        # reload_zone_tabs_ext импортируется динамически в ProjectPage
        # Здесь вызываем метод ProjectPage через page._reload_zone_tabs
        page._reload_zone_tabs()
    except Exception:
        pass



# 6.b Удаление зоны
def delete_zone(page: Any) -> None:
    """Удаляет текущую зону со всеми позициями.

    Спрашивает подтверждение, удаляет все позиции выбранной зоны из базы,
    убирает зону из JSON-персистентности, перестраивает вкладки и списки.
    """
    cur_index = page.zone_tabs.currentIndex()
    if cur_index < 0:
        return
    zone_label = page.zone_tabs.tabText(cur_index).strip()
    # 6.b.1 Канонический ключ зоны: пустая строка для любых вариантов "Без зоны".
    zone_key = _canon_zone(zone_label)
    # Логируем начало операции удаления
    logging.getLogger(__name__).info(
        "Запрошено удаление зоны: %s (canon=%s)",
        zone_label or "<Без зоны>", zone_key or "<без зоны>",
    )
    # Подтверждение
    reply = QtWidgets.QMessageBox.question(
        page, "Подтверждение",
        f"Удалить зону «{zone_label or 'Без зоны'}» со всеми позициями?",
        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
    )
    if reply != QtWidgets.QMessageBox.Yes:
        return
    # Удаляем позиции из базы
    try:
        ids: List[int] = []
        if page.project_id is not None:
            # Используем канонический ключ зоны для выборки элементов.
            cur = page.db._conn.cursor()
            if not zone_key:
                # Для пустой зоны выбираем NULL или ''.
                cur.execute(
                    "SELECT id FROM items WHERE project_id=? AND COALESCE(zone,'')=''",
                    (page.project_id,),
                )
            else:
                cur.execute(
                    "SELECT id FROM items WHERE project_id=? AND COALESCE(zone,'')=?",
                    (page.project_id, zone_key),
                )
            ids = [int(r[0]) for r in cur.fetchall()]
        if ids:
            page.db.delete_items(ids)
    except Exception as ex:
        logging.getLogger(__name__).error(
            "Ошибка удаления зоны '%s': %s", zone_label, ex, exc_info=True
        )
        QtWidgets.QMessageBox.critical(
            page, "Ошибка", f"Не удалось удалить зону: {ex}"
        )
        return
    # Обновляем JSON-файл зон
    try:
        zones = _load_persisted_zones(page)
        # Удаляем удалённую зону по каноничному ключу
        zones = [z for z in zones if _canon_zone(z).lower() != zone_key.lower()]
        _save_persisted_zones(page, zones)
    except Exception:
        logging.getLogger(__name__).error(
            "Не удалось обновить файл зон при удалении", exc_info=True
        )
    # Перестраиваем вкладки
    try:
        page._reload_zone_tabs()
    except Exception:
        pass
    # Лог
    try:
        page._log(
            f"Смета: удалена зона «{zone_label or 'Без зоны'}» "
            f"(удалено позиций: {len(ids) if 'ids' in locals() else 0})."
        )
    except Exception:
        pass
    logging.getLogger(__name__).info(
        "Удалена зона '%s' (canon=%s), удалено позиций: %d",
        zone_label or "<Без зоны>", zone_key or "<без зоны>", len(ids) if 'ids' in locals() else 0
    )
# 7. Перенос выделенных строк в другую зону
def move_selected_to_zone(page: Any) -> None:
    """Перемещает выделенные позиции из текущей вкладки в выбранную зону."""
    if MoveDialog is None:
        QtWidgets.QMessageBox.critical(page, "Ошибка", "Диалог MoveDialog не найден.")
        return

    # Получаем целевую зону: либо данные, либо текст, затем нормализуем
    target_text = (page.cmb_move_zone.currentText() or "").strip()
    target_data = page.cmb_move_zone.currentData()
    raw = target_data if target_data is not None else target_text
    target = normalize_case(raw) if raw is not None else ""
    # Если выбрана пустая зона, но есть зона по умолчанию, используем её
    if not target and getattr(page, "default_zone", ""):
        target = page.default_zone

    # Если выбранной зоны нет среди существующих — создаём её
    if target and target not in page.zone_tables:
        table = build_zone_table(page)
        page.zone_tables[target] = table
        page.zone_tabs.addTab(table, target)
        page.cmb_move_zone.addItem(target, target)
        zones = page.db.project_distinct_values(page.project_id, "zone") if page.project_id else []
        fill_manual_zone_combo(page, unique_zones)

    cur_table = page.zone_tabs.currentWidget()
    if not isinstance(cur_table, QtWidgets.QTableWidget):
        return

    rows = cur_table.selectionModel().selectedRows()
    if not rows:
        QtWidgets.QMessageBox.information(page, "Внимание", "Выберите позиции для переноса.")
        return

    items_data: List[Tuple[int, str, float, int]] = []
    for m in rows:
        r = m.row()
        name_item = cur_table.item(r, 0)
        if not name_item:
            continue
        item_id = int(name_item.data(QtCore.Qt.UserRole))
        qty = to_float(cur_table.item(r, 1).text(), 0.0)
        items_data.append((item_id, name_item.text(), qty, r))

    dlg = MoveDialog(items_data, parent=page)
    if dlg.exec() != QtWidgets.QDialog.Accepted:
        return

    moves = dlg.result_moves  # dict: item_id -> qty_to_move
    if not moves:
        return

    undo_batch = f"__undo_move__{datetime.utcnow().isoformat()}"
    originals: List[Tuple[int, float, float]] = []  # (item_id, old_qty, old_amount)

    created = 0
    updated = 0

    for item_id, move_qty in moves.items():
        if move_qty <= 0:
            continue
        row = page.db.get_item_by_id(item_id)
        if not row:
            continue

        qty_old = float(row["qty"] or 0)
        unit_price = float(row["unit_price"] or 0)
        coeff = float(row["coeff"] or 1.0)

        if move_qty >= qty_old - 1e-9:
            page.db.update_item_fields(item_id, {"zone": target})
            updated += 1
        else:
            qty_left = qty_old - move_qty
            amount_left = unit_price * qty_left * coeff

            originals.append((item_id, qty_old, float(row["amount"] or 0)))

            page.db.update_item_fields(item_id, {"qty": qty_left, "amount": amount_left})

            amount_new = unit_price * move_qty * coeff
            page.db.add_items_bulk([
                {
                    "project_id": row["project_id"],
                    "type": row["type"],
                    "group_name": row["group_name"],
                    "name": row["name"],
                    "qty": move_qty,
                    "coeff": coeff,
                    "amount": amount_new,
                    "unit_price": unit_price,
                    "source_file": row["source_file"],
                    "vendor": row["vendor"] or "",
                    "department": row["department"] or "",
                    "zone": target,
                    "power_watts": row["power_watts"] or 0,
                    "import_batch": undo_batch,
                }
            ])
            created += 1

    # Сохраняем информацию для UNDO
    page._last_action = {
        "type": "move",
        "project_id": page.project_id,
        "batch": undo_batch,
        "original": originals,
    }
    page.btn_undo_summary.setEnabled(True)

    page._log(
        f"Перенос в зону «{target or 'Без зоны'}»: обновлено {updated}, создано новых {created}."
    )
    reload_zone_tabs(page)


# 8. Добавление позиции вручную
def add_manual_item(page: Any) -> None:
    """Добавляет позицию вручную в проект и каталог, поддерживает UNDO.

    В режиме добавления из базы данных (``page._db_mode_enabled``)
    делегирует действие функции ``add_catalog_item``.
    """
    # Если активирован режим базы данных — вызываем соответствующий обработчик
    if getattr(page, "_db_mode_enabled", False):
        add_catalog_item(page)
        return
    if page.project_id is None:
        return

    name = normalize_case(page.ed_add_name.text())
    # Проверяем на пустую строку после нормализации
    if not name:
        QtWidgets.QMessageBox.information(page, "Внимание", "Введите наименование.")
        return

    qty = float(page.sp_add_qty.value() or 1.0)
    coeff = float(page.sp_add_coeff.value() or 1.0)
    price = float(page.sp_add_price.value() or 0.0)
    amount = qty * coeff * price

    class_en = CLASS_RU2EN.get(page.cmb_add_class.currentText(), "equipment")
    # Нормализуем подрядчика и отдел
    vendor = normalize_case(page.ed_add_vendor.text())
    # Получаем выбранный или введённый отдел из комбобокса и нормализуем
    department = normalize_case(page.cmb_add_department.currentText())

    zone_data = page.cmb_add_zone.currentData()
    # Выбираем текст зоны или данные и нормализуем
    zone_raw = zone_data if zone_data is not None else (page.cmb_add_zone.currentText() or "")
    zone = normalize_case(zone_raw)
    # Если пользователь выбрал «Без зоны», но зона по умолчанию задана, используем её
    if not zone and getattr(page, "default_zone", ""):
        zone = page.default_zone

    power_w = float(page.sp_add_power.value() or 0)
    if power_w <= 0:
        power_w = page.db.catalog_max_power_by_name(name) or 0

    undo_batch = f"__undo_manual_add__{datetime.utcnow().isoformat()}"

    # Проверяем наличие дубликата среди существующих позиций.
    duplicate_manual = None
    try:
        existing_rows = page.db.list_items_filtered(
            project_id=page.project_id,
            vendor="<ALL>",
            department="<ALL>",
            zone=zone,
            class_en=class_en,
            name_like=name
        )
        for r in existing_rows:
            try:
                if normalize_case(r["name"] or "") == name \
                   and normalize_case(r["vendor"] or "") == vendor \
                   and normalize_case(r["department"] or "") == department \
                   and (r["type"] or "equipment") == class_en \
                   and abs(float(r["unit_price"] or 0) - price) < 1e-6 \
                   and abs(float(r["coeff"] or 0) - coeff) < 1e-6:
                    duplicate_manual = r
                    break
            except Exception:
                continue
    except Exception as ex:
        page._log(f"Ошибка поиска дубликатов при ручном добавлении: {ex}", "error")

    if duplicate_manual:
        # Увеличиваем количество существующей строки
        try:
            old_qty_m = float(duplicate_manual["qty"] or 0)
            new_qty_m = old_qty_m + qty
            coeff_old_m = float(duplicate_manual["coeff"] or 1)
            price_old_m = float(duplicate_manual["unit_price"] or 0)
            new_amount_m = new_qty_m * coeff_old_m * price_old_m
            page.db.update_item_fields(duplicate_manual["id"], {"qty": new_qty_m, "amount": new_amount_m})
            # Фиксируем действие для UNDO как редактирование
            page._last_action = {
                "type": "edit",
                "item_id": duplicate_manual["id"],
                "old": {
                    "qty": old_qty_m,
                    "coeff": coeff_old_m,
                    "unit_price": price_old_m,
                    "amount": float(duplicate_manual["amount"] or 0),
                },
            }
            page.btn_undo_summary.setEnabled(True)
            page._log(
                f"Позиция «{name}» уже есть в зоне '{zone or 'Без зоны'}', количество увеличено на {fmt_num(qty,3)}."
            )
        except Exception as ex:
            page._log(f"Ошибка обновления позиции: {ex}", "error")
            QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось обновить позицию: {ex}")
            return
    else:
        # 8.1 Сохраняем в проект новую позицию
        page.db.add_items_bulk([
            {
                "project_id": page.project_id,
                "type": class_en,
                "group_name": "Аренда оборудования",
                "name": name,
                "qty": qty,
                "coeff": coeff,
                "amount": amount,
                "unit_price": price,
                "source_file": None,
                "vendor": vendor,
                "department": department,
                "zone": zone,
                "power_watts": power_w,
                "import_batch": undo_batch,
            }
        ])

        # 8.2 Сохраняем в каталог (глобальная БД)
        catalog_row = [
            {
                "name": name,
                "unit_price": price,
                "class": class_en,
                "vendor": vendor,
                "power_watts": power_w,
                "department": department,
            }
        ]
        if hasattr(page.db, "catalog_add_or_ignore"):
            try:
                page.db.catalog_add_or_ignore(catalog_row)
                page._log(
                    f"Каталог: добавлена/проверена позиция «{name}» (vendor='{vendor}', class='{class_en}')."
                )
            except Exception as ex:
                page._log(
                    f"Каталог: не удалось добавить позицию «{name}»: {ex}", "error"
                )
        else:
            page._log(
                "Каталог: метод catalog_add_or_ignore отсутствует — пропускаем добавление.",
                "error",
            )

        # 8.3 Готовим UNDO и лог
        page._last_action = {
            "type": "manual_add",
            "project_id": page.project_id,
            "batch": undo_batch,
        }
        page.btn_undo_summary.setEnabled(True)

        page._log(
            f"Добавлена позиция вручную (проект+каталог): «{name}», qty={fmt_num(qty,3)}, "
            f"coeff={fmt_num(coeff,3)}, price={fmt_num(price,2)}, power={fmt_num(power_w,0)} Вт, зона='{zone or 'Без зоны'}'."
        )

    # 8.4 Сбрасываем форму и обновляем таблицы
    page.ed_add_name.clear()
    page.sp_add_qty.setValue(1.000)
    page.sp_add_coeff.setValue(1.000)
    page.ed_add_vendor.clear()
    page.cmb_add_department.setCurrentIndex(-1)
    page.cmb_add_zone.setCurrentIndex(0)
    page.sp_add_power.setValue(0)
    page.sp_add_price.setValue(0.0)

    reload_zone_tabs(page)


# 9. Изменение ячейки (qty/coeff/price)
def on_summary_item_changed(page: Any, item: QtWidgets.QTableWidgetItem) -> None:
    """Обрабатывает изменение количества, коэффициента или цены в таблице зоны."""
    table = item.tableWidget()
    row = item.row()

    name_item = table.item(row, 0)
    if not name_item:
        return

    item_id = int(name_item.data(QtCore.Qt.UserRole))

    try:
        qty = to_float(table.item(row, 1).text(), 0.0)
        coeff = to_float(table.item(row, 2).text(), 0.0)
        price = to_float(table.item(row, 3).text(), 0.0)
        amount = price * qty * coeff
    except Exception:
        return

    old_row = page.db.get_item_by_id(item_id)
    old_fields: Dict[str, float] = {}
    if old_row:
        # sqlite3.Row поддерживает доступ к полям по ключу, но не метод get()
        # Используем значение поля или 0.0, затем приводим к float
        old_fields = {
            "qty": float(old_row["qty"] or 0.0),
            "coeff": float(old_row["coeff"] or 0.0),
            "unit_price": float(old_row["unit_price"] or 0.0),
            "amount": float(old_row["amount"] or 0.0),
        }

    fields = {"qty": qty, "coeff": coeff, "unit_price": price, "amount": amount}

    try:
        page.db.update_item_fields(item_id, fields)
        table.blockSignals(True)
        table.item(row, 4).setText(fmt_num(amount, 2))
        table.blockSignals(False)

        # После изменения строки пересчитываем сумму только для активной зоны с учётом текущих фильтров
        try:
            cur_idx = page.zone_tabs.currentIndex() if hasattr(page, "zone_tabs") else -1
        except Exception:
            cur_idx = -1
        cur_sum = 0.0
        if cur_idx >= 0:
            try:
                keys = list(page.zone_tables.keys())
                if 0 <= cur_idx < len(keys):
                    cur_zone_key = keys[cur_idx]
                    # Фильтры идентичны применяемым в reload_zone_tabs
                    vendor_f = page.cmb_f_vendor.currentText()
                    if vendor_f == "<Все подрядчики>": vendor_f = "<ALL>"
                    department_f = page.cmb_f_department.currentText()
                    if department_f == "<Все отделы>": department_f = "<ALL>"
                    class_ru_f = page.cmb_f_class.currentText()
                    class_en_f = CLASS_RU2EN.get(class_ru_f) if class_ru_f and class_ru_f != "<Все классы>" else "<ALL>"
                    rows_cur = page.db.list_items_filtered(
                        project_id=page.project_id,
                        vendor=vendor_f,
                        department=department_f,
                        zone=cur_zone_key,
                        class_en=class_en_f,
                        name_like=None,
                    )
                    search_raw_f = page.ed_search.text() if hasattr(page, "ed_search") else ""
                    if search_raw_f:
                        tmp = []
                        for r in rows_cur:
                            try:
                                nm = r["name"] or ""
                                ven = r["vendor"] or ""
                                dep = r["department"] or ""
                                zn = r["zone"] or ""
                                if (
                                    contains_search(nm, search_raw_f)
                                    or contains_search(ven, search_raw_f)
                                    or contains_search(dep, search_raw_f)
                                    or contains_search(zn, search_raw_f)
                                ):
                                    tmp.append(r)
                            except Exception:
                                continue
                        rows_cur = tmp
                    for r in rows_cur:
                        try:
                            cur_sum += float(r["amount"] or 0.0)
                        except Exception:
                            continue
            except Exception:
                cur_sum = 0.0
        page.label_total.setText(f"Итого: {fmt_num(cur_sum, 2)}")
        # Обновляем бухгалтерию при изменении позиции
        try:
            if hasattr(page, "recalc_finance"):
                page.recalc_finance()
        except Exception:
            pass

        page._last_action = {
            "type": "edit",
            "project_id": page.project_id,
            "item_id": item_id,
            "old": old_fields,
        }
        page.btn_undo_summary.setEnabled(True)

        page._log(
            f"Изменена строка id={item_id}: qty={qty}, coeff={coeff}, price={price}, amount={amount}."
        )
    except Exception as ex:
        page._log(f"Ошибка обновления позиции в смете: {ex}", "error")


# 10. Удаление выделенных строк
def delete_selected(page: Any) -> None:
    """Удаляет выделенные позиции из проекта, сохраняя снимок для UNDO."""
    cur_table = page.zone_tabs.currentWidget()
    if not isinstance(cur_table, QtWidgets.QTableWidget):
        return

    rows = cur_table.selectionModel().selectedRows()
    if not rows:
        QtWidgets.QMessageBox.information(page, "Внимание", "Выберите строки для удаления.")
        return

    ids: List[int] = []
    for m in rows:
        r = m.row()
        name_item = cur_table.item(r, 0)
        if not name_item:
            continue
        rid = name_item.data(QtCore.Qt.UserRole)
        if rid is not None:
            ids.append(int(rid))

    if not ids:
        return

    if QtWidgets.QMessageBox.question(
        page,
        "Подтверждение",
        f"Удалить {len(ids)} позиций?",
    ) != QtWidgets.QMessageBox.Yes:
        return

    snapshot: List[Dict[str, Any]] = []
    for _id in ids:
        row = page.db.get_item_by_id(int(_id))
        if row:
            # sqlite3.Row не поддерживает метод get(); используем доступ по ключу и проверяем наличие
            created_at = row["created_at"] if "created_at" in row.keys() else None
            import_batch = row["import_batch"] if "import_batch" in row.keys() else None
            snapshot.append({
                "project_id": row["project_id"],
                "type": row["type"],
                "group_name": row["group_name"],
                "name": row["name"],
                "qty": float(row["qty"] or 0),
                "coeff": float(row["coeff"] or 1),
                "amount": float(row["amount"] or 0),
                "unit_price": float(row["unit_price"] or 0),
                "source_file": row["source_file"],
                "created_at": created_at,
                "vendor": row["vendor"] or "",
                "department": row["department"] or "",
                "zone": row["zone"] or "",
                "power_watts": float(row["power_watts"] or 0),
                "import_batch": import_batch,
            })

    try:
        page.db.delete_items(ids)
        page._last_action = {
            "type": "delete",
            "project_id": page.project_id,
            "rows": snapshot,
        }
        page.btn_undo_summary.setEnabled(True)
        page._log(f"Смета: удалено позиций {len(ids)}.")
    except Exception as ex:
        page._log(f"Ошибка удаления позиций: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось удалить: {ex}")
        return

    reload_zone_tabs(page)


# 11. Отмена последнего действия
def undo_last_summary(page: Any) -> None:
    """Отменяет последнее действие: ручное добавление, удаление, перенос или редактирование."""
    if not page._last_action:
        QtWidgets.QMessageBox.information(page, "Отмена", "Нет действий для отмены.")
        return

    act = page._last_action
    try:
        if act.get("type") == "manual_add":
            # Удаляем все записи данного batch
            page.db.delete_items_by_import_batch(act["project_id"], act["batch"])
            page._log("Отменено добавление позиции вручную (позиции удалены из проекта).")

        elif act.get("type") == "delete":
            rows = act.get("rows", [])
            if rows:
                page.db.add_items_bulk(rows)
                page._log(f"Отменено удаление: восстановлено {len(rows)} записей.")

        elif act.get("type") == "move":
            batch = act.get("batch")
            if batch:
                page.db.delete_items_by_import_batch(act["project_id"], batch)
            for item_id, qty, amount in act.get("original", []):
                page.db.update_item_fields(item_id, {"qty": qty, "amount": amount})
            page._log("Отменён перенос по зонам.")

        elif act.get("type") == "edit":
            item_id = act.get("item_id")
            old = act.get("old", {})
            if item_id and old:
                page.db.update_item_fields(
                    item_id,
                    {
                        "qty": old.get("qty", 0.0),
                        "coeff": old.get("coeff", 1.0),
                        "unit_price": old.get("unit_price", 0.0),
                        "amount": old.get("amount", 0.0),
                    },
                )
                page._log(f"Отменено редактирование строки id={item_id}.")

    finally:
        page._last_action = None
        page.btn_undo_summary.setEnabled(False)
        reload_zone_tabs(page)


# 12. Создание снимка и сравнение смет
def take_snapshot(page: Any) -> None:
    """Сохраняет текущую сводную смету для последующего сравнения.

    Снимок включает список зон и копию всех позиций проекта на данный
    момент. Используется для анализа изменений между двумя состояниями.
    После создания снимка пользователь может включить режим сравнения
    через чекбокс ``chk_snapshot_compare``. Ограничение: сравнение
    допускается только если набор зон совпадает.
    """
    if page.project_id is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Проект не выбран.")
        return
    # Получаем все уникальные зоны текущего проекта
    zones = page.db.project_distinct_values(page.project_id, "zone") or []
    # Создаем структуру снимка. Помимо зон и элементов включаем идентификатор проекта
    # и имя снимка (оставляем пустым, будет заполнено при сохранении в файл).
    page._snapshot_data = {
        "project_id": page.project_id,
        "name": "",
        "zones": list(zones),
        "items": {},  # item_id -> dict(row data)
    }
    # Берём все позиции проекта без фильтрации
    all_rows = page.db.list_items(page.project_id)
    for row in all_rows:
        page._snapshot_data["items"][row["id"]] = {
            "id": row["id"],
            "name": row["name"],
            "qty": float(row["qty"] or 0),
            "coeff": float(row["coeff"] or 0),
            "unit_price": float(row["unit_price"] or 0),
            "amount": float(row["amount"] or 0),
            "vendor": row["vendor"] or "",
            "department": row["department"] or "",
            "zone": row["zone"] or "",
            "class": row["type"] or "equipment",
            "power_watts": float(row["power_watts"] or 0),
        }
    # Сбрасываем состояние сравнения
    page._snapshot_compare_enabled = False
    page.chk_snapshot_compare.blockSignals(True)
    page.chk_snapshot_compare.setChecked(False)
    page.chk_snapshot_compare.blockSignals(False)
    # Вычисляем снимок финансового отчёта и сохраняем его в структуру снимка
    try:
        fin_snap = compute_fin_snapshot_data(page)
        page._snapshot_data["fin_snapshot"] = fin_snap
    except Exception as ex:
        # Ошибку записываем в лог, но не прерываем создание снимка сметы
        page._log(f"Ошибка создания снимка финансового отчёта: {ex}", "error")
    page._log("Снимок сводной сметы создан.")
    QtWidgets.QMessageBox.information(page, "Снимок создан", "Текущая сводная смета сохранена для сравнения.")


def toggle_snapshot_compare(page: Any) -> None:
    """Включает или отключает режим сравнения с сохранённым снимком.

    При включении проверяет наличие снимка и совпадение набора зон.
    Если сравнение возможно, обновляет таблицы с дополнительными колонками.
    Иначе выводит предупреждение и отключает чекбокс.
    """
    # Проверяем, что снимок существует
    has_snap = hasattr(page, "_snapshot_data") and bool(getattr(page, "_snapshot_data", None))
    if not has_snap:
        QtWidgets.QMessageBox.information(page, "Нет снимка", "Снимок для сравнения не создан.")
        page.chk_snapshot_compare.blockSignals(True)
        page.chk_snapshot_compare.setChecked(False)
        page.chk_snapshot_compare.blockSignals(False)
        return
    if page.chk_snapshot_compare.isChecked():
        # Сравниваем набор зон
        snap_zones = set(page._snapshot_data.get("zones", []))
        cur_zones = set(page.db.project_distinct_values(page.project_id, "zone") or [])
        snap_zones = {z or "" for z in snap_zones}
        cur_zones = {z or "" for z in cur_zones}
        if snap_zones != cur_zones:
            QtWidgets.QMessageBox.warning(
                page,
                "Несовместимые зоны",
                "Режим сравнения для этого снимка недоступен, так как была другая компоновка зон."
            )
            page.chk_snapshot_compare.blockSignals(True)
            page.chk_snapshot_compare.setChecked(False)
            page.chk_snapshot_compare.blockSignals(False)
            page._snapshot_compare_enabled = False
            return
        page._snapshot_compare_enabled = True
        page._log("Режим сравнения включён.")
    else:
        page._snapshot_compare_enabled = False
        page._log("Режим сравнения выключен.")
    reload_zone_tabs(page)


# 13. Работа с сохранёнными снимками
def snapshots_dir_for_project(page: Any) -> Path:
    """
    Возвращает путь к каталогу для сохранения снимков текущего проекта.

    Снимки хранятся в папке ``assets/project_<id>/snapshots``. При первом
    обращении каталог создаётся автоматически. Это обеспечивает
    привязку снимков к конкретному проекту и сохранение их вместе
    с другими ресурсами проекта (картинки, логотипы).

    :param page: текущая страница ProjectPage
    :return: путь к каталогу снимков для проекта
    """
    from .common import ASSETS_DIR  # импорт здесь, чтобы избежать циклов при импорте
    # Если идентификатор проекта отсутствует, используем общий каталог в DATA_DIR/snapshots
    if not getattr(page, "project_id", None):
        root = DATA_DIR / "snapshots"
        root.mkdir(parents=True, exist_ok=True)
        return root
    proj_id = page.project_id
    # Каталог вида assets/project_<id>/snapshots
    root = ASSETS_DIR / f"project_{proj_id}" / "snapshots"
    root.mkdir(parents=True, exist_ok=True)
    return root


def save_snapshot(page: Any) -> None:
    """
    Сохраняет текущую сводную смету в файл и предлагает пользователю
    присвоить имя снимку. После сохранения список снимков обновляется.
    """
    if page.project_id is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Проект не выбран.")
        return
    # Запрашиваем у пользователя имя снимка
    name, ok = QtWidgets.QInputDialog.getText(page, "Имя снимка", "Введите имя для снимка:")
    if not ok or not name.strip():
        return
    name = name.strip()
    # Делаем временную копию текущего состояния
    take_snapshot(page)
    snap = getattr(page, "_snapshot_data", None)
    if not snap:
        page._log("Не удалось создать снимок.", "error")
        return
    # Создаём структуру для сохранения: копируем уже сформированный снимок,
    # присваиваем имя и проверяем наличие project_id. Если снимок был
    # сформирован в ``take_snapshot``, он уже содержит project_id. Иначе
    # устанавливаем его явно.
    snap_data = {
        "project_id": snap.get("project_id", page.project_id),
        "name": name,
        "zones": snap.get("zones", []),
        "items": snap.get("items", {}),
        # Добавляем снимок финансового отчёта, если он присутствует в временной структуре
        "fin_snapshot": snap.get("fin_snapshot", {})
    }
    # Обновляем имя во временном снимке, чтобы последующее сравнение
    # использовало читаемое имя в логах и отладке.
    try:
        page._snapshot_data["name"] = name
    except Exception:
        pass
    # Формируем путь к файлу: project_<id>_<timestamp>.json
    timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    filename = f"project_{page.project_id}_{timestamp}.json"
    path = snapshots_dir_for_project(page) / filename
    try:
        # Сохраняем данные в файл JSON
        with open(path, "w", encoding="utf-8") as f:
            json.dump(snap_data, f, ensure_ascii=False, indent=2)
        # Информируем пользователя через лог
        page._log(f"Снимок «{name}» сохранён.")
        # После сохранения обновляем список снимков для проекта
        load_snapshot_list(page)
        # Автоматически выбираем только что сохранённый снимок в списке
        try:
            cmb = getattr(page, "cmb_snapshot", None)
            if cmb is not None:
                # Ищем индекс элемента, чей userData совпадает с путём сохранённого файла
                for i in range(cmb.count()):
                    if str(cmb.itemData(i)) == str(path):
                        # Устанавливаем индекс комбобокса на новый снимок
                        cmb.setCurrentIndex(i)
                        # Явно вызываем обработчик выбора снимка, чтобы загрузить данные
                        try:
                            on_snapshot_selected(page)
                        except Exception:
                            pass
                        break
        except Exception:
            # Игнорируем ошибки выбора
            pass
    except Exception as ex:
        page._log(f"Ошибка сохранения снимка: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось сохранить снимок: {ex}")


def load_snapshot_list(page: Any) -> None:
    """
    Загружает список снимков для текущего проекта и заполняет выпадающий
    список ``cmb_snapshot``. Список обновляется только если он
    относится к текущему проекту.
    """
    # Если проект ещё не создан или комбобокса нет — ничего не делаем
    if page.project_id is None or not hasattr(page, "cmb_snapshot"):
        return
    # Очищаем список
    page.cmb_snapshot.blockSignals(True)
    page.cmb_snapshot.clear()
    # Добавляем элемент по умолчанию
    page.cmb_snapshot.addItem("<Выберите снимок>", None)
    snap_dir = snapshots_dir_for_project(page)
    entries = []
    # Находим все файлы этого проекта по маске project_<id>_*.json
    pattern = f"project_{page.project_id}_*.json"
    for f in snap_dir.glob(pattern):
        try:
            with open(f, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            # Сохраняем кортеж: отображаемое имя, путь
            name = data.get("name") or f.stem
            entries.append((name, f))
        except Exception as ex:
            # Логируем ошибку чтения конкретного файла, но не прерываем сбор списка
            page._log(f"Ошибка чтения файла снимка {f}: {ex}", "error")
            continue
    # Сортируем снимки по имени (можно изменить при необходимости)
    entries.sort(key=lambda t: t[0].lower())
    for name, path in entries:
        page.cmb_snapshot.addItem(name, str(path))
    page.cmb_snapshot.blockSignals(False)
    # Сбрасываем текущий выбор
    page.cmb_snapshot.setCurrentIndex(0)
    # Информируем пользователя в логах, что список снимков обновлён
    page._log("Список снимков обновлён.")


def on_snapshot_selected(page: Any) -> None:
    """
    Загружает выбранный снимок из файла и устанавливает его как
    текущий для сравнения. Если выбран пункт по умолчанию, снимок
    сбрасывается.
    """
    if page.project_id is None:
        return
    idx = page.cmb_snapshot.currentIndex() if hasattr(page, "cmb_snapshot") else 0
    data = page.cmb_snapshot.itemData(idx) if hasattr(page, "cmb_snapshot") else None
    if not data:
        # Выбран пункт по умолчанию — сбрасываем снимок
        if hasattr(page, "_snapshot_data"):
            delattr(page, "_snapshot_data")
        page._snapshot_compare_enabled = False
        page.chk_snapshot_compare.blockSignals(True)
        page.chk_snapshot_compare.setChecked(False)
        page.chk_snapshot_compare.blockSignals(False)
        reload_zone_tabs(page)
        # Логируем, что снимок не выбран
        page._log("Снимок не выбран, режим сравнения отключён.")
        return
    # Загружаем файл
    try:
        path = Path(data)
        with open(path, "r", encoding="utf-8") as f:
            snap_data = json.load(f)
        # Проверяем, что снимок принадлежит текущему проекту
        if snap_data.get("project_id") != page.project_id:
            QtWidgets.QMessageBox.warning(page, "Несовместимый снимок", "Этот снимок относится к другому проекту.")
            return
        # Приводим ключи словаря items к целочисленному типу, так как после
        # загрузки из JSON ключи становятся строками. Без этого сравнение
        # по идентификаторам элементов будет некорректно.
        items_dict = snap_data.get("items", {}) or {}
        converted: Dict[int, Any] = {}
        for k, v in items_dict.items():
            try:
                converted[int(k)] = v
            except Exception:
                # Если ключ не преобразуется в int, пропускаем его
                continue
        snap_data["items"] = converted
        # Устанавливаем данные и отключаем режим сравнения
        page._snapshot_data = snap_data
        page._snapshot_compare_enabled = False
        page.chk_snapshot_compare.blockSignals(True)
        page.chk_snapshot_compare.setChecked(False)
        page.chk_snapshot_compare.blockSignals(False)
        # Записываем в лог имя снимка
        page._log(f"Снимок «{snap_data.get('name', '')}» загружен.")
        reload_zone_tabs(page)
    except Exception as ex:
        page._log(f"Ошибка загрузки снимка: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось загрузить снимок: {ex}")


# 14. Переключение режима ручного добавления / из базы данных
def toggle_db_mode(page: Any) -> None:
    """Переключает панель ручного добавления между режимом ввода вручную и
    добавлением из каталога базы данных.

    В режиме базы данных отображается поле поиска по каталогу и
    фильтры по подрядчику и отделу. Поля подрядчика, отдела и
    выбора класса скрываются, цена/шт и потребление берутся из каталога
    автоматически и редактирование блокируется.
    """
    # Если на странице нет переключателя режима БД, значит данный код уже не используется.
    # Добавлен защитный код, чтобы функции, оставшиеся для обратной совместимости,
    # не вызывали исключений при доступе к отсутствующим атрибутам.
    if not hasattr(page, 'chk_db_mode'):
        # Старый переключатель удалён, просто выходим из функции.
        return
    state = page.chk_db_mode.isChecked()
    # Запоминаем состояние режима базы
    page._db_mode_enabled = bool(state)
    # Поле наименования: видим либо строку ввода, либо комбобокс поиска
    page.ed_add_name.setVisible(not state)
    page.cmb_search_name.setVisible(state)
    # Поле подрядчика и комбобоксы отдела: переключаем видимость
    page.ed_add_vendor.setVisible(not state)
    page.cmb_add_department.setVisible(not state)
    page.cmb_filter_vendor.setVisible(state)
    page.cmb_filter_department.setVisible(state)
    # Блокируем редактирование цены, класса и потребления в режиме базы
    page.sp_add_price.setReadOnly(state)
    page.cmb_add_class.setEnabled(not state)
    page.sp_add_power.setReadOnly(state)
    # При включении режима базы подгружаем доступные фильтры и обновляем подсказки.
    # Соединение сигналов выполняется один раз при построении вкладки.
    if state:
        load_catalog_filters(page)
        update_catalog_suggestions(page)
    else:
        # При выходе из режима базы очищаем подсказки и восстанавливаем поля
        page.cmb_search_name.blockSignals(True)
        page.cmb_search_name.clear()
        page.cmb_search_name.blockSignals(False)
        # Восстанавливаем возможность ввода цены, класса и потребления
        page.sp_add_price.setReadOnly(False)
        page.cmb_add_class.setEnabled(True)
        page.sp_add_power.setReadOnly(False)
        # Сбрасываем поля ввода
        page.ed_add_name.clear()
        page.ed_add_vendor.clear()
        page.cmb_add_department.setCurrentIndex(-1)
        page.sp_add_price.setValue(0.0)
        page.cmb_add_class.setCurrentIndex(0)
        page.sp_add_power.setValue(0)


# 15. Загрузка фильтров каталога (подрядчики и отделы)
def load_catalog_filters(page: Any) -> None:
    """Заполняет фильтры подрядчика и отдела из глобального каталога.

    Использует метод DB.catalog_distinct_values для получения уникальных
    значений. В начало списков добавляется элемент «<Любой>».
    """
    # Если фильтры отсутствуют, значит режим базы не используется – выходим.
    if not (hasattr(page, "cmb_filter_vendor") and hasattr(page, "cmb_filter_department")):
        return
    try:
        vendors = page.db.catalog_distinct_values("vendor")
        departments = page.db.catalog_distinct_values("department")
    except Exception as ex:
        page._log(f"Ошибка загрузки фильтров каталога: {ex}", "error")
        vendors = []
        departments = []
    # Заполняем комбобоксы фильтров
    page.cmb_filter_vendor.blockSignals(True)
    page.cmb_filter_vendor.clear()
    page.cmb_filter_vendor.addItem("<Любой>", None)
    for v in vendors:
        page.cmb_filter_vendor.addItem(normalize_case(v), v)
    page.cmb_filter_vendor.setCurrentIndex(0)
    page.cmb_filter_vendor.blockSignals(False)

    page.cmb_filter_department.blockSignals(True)
    page.cmb_filter_department.clear()
    page.cmb_filter_department.addItem("<Любой>", None)
    for d in departments:
        page.cmb_filter_department.addItem(normalize_case(d), d)
    page.cmb_filter_department.setCurrentIndex(0)
    page.cmb_filter_department.blockSignals(False)


# 16. Обновление списка подсказок каталога
def update_catalog_suggestions(page: Any) -> None:
    """Обновляет выпадающий список ``cmb_search_name`` в зависимости от текста
    поиска и выбранных фильтров подрядчика и отдела.

    Отображаемый текст включает наименование и текущую цену, чтобы
    пользователь видел ориентировочную стоимость. Для каждого элемента
    записываются данные каталога в ``itemData``, чтобы затем можно было
    быстро заполнить поля при выборе.
    """
    # Если режим базы данных не активен или нужные виджеты отсутствуют, не обновляем подсказки.
    if not getattr(page, "_db_mode_enabled", False):
        return
    # Защитимся от случая, когда модуль базы данных отключён и поля поиска отсутствуют.
    if not (hasattr(page, "cmb_search_name") and hasattr(page, "cmb_filter_vendor") and hasattr(page, "cmb_filter_department")):
        return
    text = ""
    try:
        text = page.cmb_search_name.lineEdit().text().strip()
    except Exception:
        pass
    vendor_filter = page.cmb_filter_vendor.currentData()
    dept_filter = page.cmb_filter_department.currentData()
    filters: Dict[str, Any] = {}
    if text:
        filters["name"] = text
    # Передаём фильтр, если выбран конкретный подрядчик/отдел
    if vendor_filter:
        filters["vendor"] = vendor_filter
    if dept_filter:
        filters["department"] = dept_filter
    # Получаем подходящие позиции каталога
    rows: List[Any] = []
    try:
        rows = page.db.catalog_list(filters)
    except Exception as ex:
        page._log(f"Ошибка запроса каталога: {ex}", "error")
    # Обновляем комбобокс подсказок
    # Сохраняем текущий текст, который ввёл пользователь, чтобы восстановить его
    current_text = ""
    try:
        current_text = page.cmb_search_name.lineEdit().text()
    except Exception:
        pass
    page.cmb_search_name.blockSignals(True)
    page.cmb_search_name.clear()
    # Добавляем элементы подсказок: показываем имя и цену, записываем данные в itemData
    for r in rows:
        try:
            name_norm = normalize_case(r["name"] or "")
            price = float(r["unit_price"] or 0)
            display = f"{name_norm} (цена: {fmt_num(price,2)})"
            page.cmb_search_name.addItem(display, dict(r))
        except Exception:
            continue
    # Восстанавливаем текст, который вводил пользователь, и не выбираем ни один элемент
    try:
        # Восстанавливаем текст непосредственно через lineEdit, чтобы
        # символы, вводимые пользователем, появлялись в поле ввода
        le = page.cmb_search_name.lineEdit()
        le.setText(current_text)
        # Ставим курсор в конец текста
        le.setCursorPosition(len(current_text))
    except Exception:
        pass
    page.cmb_search_name.setCurrentIndex(-1)
    page.cmb_search_name.blockSignals(False)
    # Если есть найденные элементы и введён текст, показываем выпадающий список подсказок
    try:
        if rows and current_text:
            page.cmb_search_name.showPopup()
    except Exception:
        pass


# 17. Обработка выбора позиции из каталога
def on_catalog_item_selected(page: Any) -> None:
    """Заполняет поля ручного добавления выбранной позицией из каталога.

    После выбора позиции из списка подсказок наименование, цена,
    класс, подрядчик, отдел и потребление подставляются в
    соответствующие поля, причём редактирование цены и питания
    блокируется до выхода из режима базы.
    """
    # Если режим базы данных не активен или комбобокс поиска отсутствует — игнорируем выбор
    if not getattr(page, "_db_mode_enabled", False) or not hasattr(page, "cmb_search_name"):
        return
    idx = page.cmb_search_name.currentIndex()
    if idx < 0:
        return
    data = page.cmb_search_name.itemData(idx)
    if not data:
        return
    row = data
    try:
        # Заполняем внутренние поля для последующего добавления
        name_norm = normalize_case(row.get("name", ""))
        vendor_norm = normalize_case(row.get("vendor", ""))
        dept_norm = normalize_case(row.get("department", ""))
        class_en = row.get("class", "equipment")
        class_ru = CLASS_EN2RU.get(class_en, "Оборудование")
        price = float(row.get("unit_price", 0.0) or 0.0)
        power = float(row.get("power_watts", 0.0) or 0.0)
        # Устанавливаем значения в скрытые поля
        page.ed_add_name.setText(name_norm)
        page.ed_add_vendor.setText(vendor_norm)
        page.cmb_add_department.setEditText(dept_norm)
        page.cmb_add_class.setCurrentText(class_ru)
        page.sp_add_price.setValue(price)
        page.sp_add_power.setValue(power)
        # Блокируем редактирование цены и потребления на всякий случай
        page.sp_add_price.setReadOnly(True)
        page.sp_add_power.setReadOnly(True)
    except Exception as ex:
        page._log(f"Ошибка заполнения позиции каталога: {ex}", "error")


# 18. Добавление позиции из каталога в смету
def add_catalog_item(page: Any) -> None:
    """Добавляет выбранную позицию из каталога в проект.

    При отсутствии выбранной позиции выводит уведомление. Использует
    количество, коэффициент и выбранную зону из панели добавления.
    """
    # Делаем проверку существования поля поиска из каталога. Если его нет, функция неактуальна.
    if not hasattr(page, "cmb_search_name"):
        return
    if page.project_id is None:
        return
    idx = page.cmb_search_name.currentIndex()
    if idx < 0:
        QtWidgets.QMessageBox.information(page, "Внимание", "Выберите позицию из каталога.")
        return
    data = page.cmb_search_name.itemData(idx)
    if not data:
        QtWidgets.QMessageBox.information(page, "Внимание", "Выберите позицию из каталога.")
        return
    row = data
    # Собираем данные
    name = normalize_case(row.get("name", ""))
    qty = float(page.sp_add_qty.value() or 1.0)
    coeff = float(page.sp_add_coeff.value() or 1.0)
    price = float(row.get("unit_price", 0.0) or 0.0)
    amount = qty * coeff * price
    vendor = normalize_case(row.get("vendor", ""))
    department = normalize_case(row.get("department", ""))
    class_en = row.get("class", "equipment")
    zone_data = page.cmb_add_zone.currentData()
    zone_raw = zone_data if zone_data is not None else (page.cmb_add_zone.currentText() or "")
    zone = normalize_case(zone_raw)
    # Если пользователь выбрал «Без зоны», но зона по умолчанию задана, используем её
    if not zone and getattr(page, "default_zone", ""):
        zone = page.default_zone
    power = float(row.get("power_watts", 0.0) or 0.0)
    # Проверяем, существует ли уже в смете позиция с теми же параметрами (имя, подрядчик, отдел,
    # класс, цена, коэффициент и зона). Если такая найдена, увеличиваем её количество,
    # иначе создаём новую запись.
    duplicate = None
    try:
        # Получаем все позиции в заданной зоне данного проекта
        existing_rows = page.db.list_items_filtered(
            project_id=page.project_id,
            vendor="<ALL>",
            department="<ALL>",
            zone=zone,
            class_en=class_en,
            name_like=name
        )
        for r in existing_rows:
            try:
                # Сравниваем по нормализованным полям
                if normalize_case(r["name"] or "") == name \
                   and normalize_case(r["vendor"] or "") == vendor \
                   and normalize_case(r["department"] or "") == department \
                   and (r["type"] or "equipment") == class_en \
                   and abs(float(r["unit_price"] or 0) - price) < 1e-6 \
                   and abs(float(r["coeff"] or 0) - coeff) < 1e-6:
                    duplicate = r
                    break
            except Exception:
                continue
    except Exception as ex:
        page._log(f"Ошибка поиска дубликатов при добавлении из каталога: {ex}", "error")

    if duplicate:
        # Увеличиваем количество существующей позиции
        try:
            old_qty = float(duplicate["qty"] or 0)
            new_qty = old_qty + qty
            # Сумма рассчитывается на основе нового количества, исходного коэффициента и цены
            coeff_old = float(duplicate["coeff"] or 1)
            price_old = float(duplicate["unit_price"] or 0)
            new_amount = new_qty * coeff_old * price_old
            page.db.update_item_fields(duplicate["id"], {"qty": new_qty, "amount": new_amount})
            # Фиксируем действие для UNDO как редактирование
            page._last_action = {
                "type": "edit",
                "item_id": duplicate["id"],
                "old": {
                    "qty": old_qty,
                    "coeff": coeff_old,
                    "unit_price": price_old,
                    "amount": float(duplicate["amount"] or 0),
                },
            }
            page.btn_undo_summary.setEnabled(True)
            page._log(
                f"Позиция «{name}» уже есть в зоне '{zone or 'Без зоны'}', количество увеличено на {fmt_num(qty,3)}."
            )
        except Exception as ex:
            page._log(f"Ошибка обновления количества позиции: {ex}", "error")
            QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось обновить позицию: {ex}")
            return
    else:
        # Записываем в проект новую позицию
        undo_batch = f"__undo_catalog_add__{datetime.utcnow().isoformat()}"
        try:
            page.db.add_items_bulk([
                {
                    "project_id": page.project_id,
                    "type": class_en,
                    "group_name": "Каталог",
                    "name": name,
                    "qty": qty,
                    "coeff": coeff,
                    "amount": amount,
                    "unit_price": price,
                    "source_file": None,
                    "vendor": vendor,
                    "department": department,
                    "zone": zone,
                    "power_watts": power,
                    "import_batch": undo_batch,
                }
            ])
            # Обновляем undo-данные
            page._last_action = {
                "type": "manual_add",
                "project_id": page.project_id,
                "batch": undo_batch,
            }
            page.btn_undo_summary.setEnabled(True)
            page._log(
                f"Добавлена позиция из каталога: «{name}», qty={fmt_num(qty,3)}, coeff={fmt_num(coeff,3)}, "
                f"price={fmt_num(price,2)}, power={fmt_num(power,0)} Вт, зона='{zone or 'Без зоны'}'."
            )
        except Exception as ex:
            page._log(f"Ошибка добавления позиции из каталога: {ex}", "error")
            QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось добавить позицию: {ex}")
            return
    # После добавления очищаем форму и отключаем выбор
    page.cmb_search_name.setCurrentIndex(-1)
    page.sp_add_qty.setValue(1.000)
    page.sp_add_coeff.setValue(1.000)
    page.sp_add_price.setValue(0.0)
    page.sp_add_price.setReadOnly(True)
    page.sp_add_power.setValue(0.0)
    page.sp_add_power.setReadOnly(True)
    page.ed_add_name.clear()
    page.ed_add_vendor.clear()
    page.cmb_add_department.setCurrentIndex(-1)
    page.cmb_add_class.setCurrentIndex(0)
    reload_zone_tabs(page)


# 19. Открытие диалога выбора позиции из базы данных
def show_catalog_dialog(page: Any) -> None:
    """Открывает модальное окно для выбора позиции из каталога базы данных.

    Диалог предоставляет фильтры (поиск по наименованию, подрядчик, отдел),
    таблицу каталога и поля для указания количества, коэффициента и зоны.
    После выбора строки и нажатия кнопки «Добавить» позиция будет
    добавлена в проект с указанными параметрами.

    :param page: объект ProjectPage для доступа к базе данных и логированию
    """
    # Если проект не выбран, ничего не делаем
    if page.project_id is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return

    class CatalogDialog(QtWidgets.QDialog):
        """Внутренний класс диалога выбора позиции из каталога."""
        def __init__(self, parent_page: Any):  # type: ignore
            super().__init__(parent_page)
            self.page = parent_page
            self.setWindowTitle("Выбор позиции из базы данных")
            self.resize(800, 500)
            # Основная вертикальная компоновка
            v_layout = QtWidgets.QVBoxLayout(self)
            v_layout.setContentsMargins(8, 8, 8, 8)
            v_layout.setSpacing(6)
            # 19.1 Фильтры
            filt = QtWidgets.QHBoxLayout()
            filt.setSpacing(6)
            # Поле поиска по наименованию
            self.ed_search = QtWidgets.QLineEdit()
            self.ed_search.setPlaceholderText("Поиск по наименованию…")
            self.ed_search.setMinimumWidth(200)
            filt.addWidget(QtWidgets.QLabel("Поиск:"))
            filt.addWidget(self.ed_search)
            # Комбо подрядчика
            self.cmb_vendor = QtWidgets.QComboBox()
            self.cmb_vendor.setMinimumWidth(150)
            filt.addWidget(QtWidgets.QLabel("Подрядчик:"))
            filt.addWidget(self.cmb_vendor)
            # Комбо отдела
            self.cmb_department = QtWidgets.QComboBox()
            self.cmb_department.setMinimumWidth(150)
            filt.addWidget(QtWidgets.QLabel("Отдел:"))
            filt.addWidget(self.cmb_department)
            filt.addStretch(1)
            v_layout.addLayout(filt)
            # 19.2 Таблица каталога
            self.tbl = QtWidgets.QTableWidget(0, 6)
            self.tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
            self.tbl.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            self.tbl.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
            headers = ["Наименование", "Класс", "Подрядчик", "Цена", "Потр. (Вт)", "Отдел"]
            self.tbl.setHorizontalHeaderLabels(headers)
            self.tbl.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            v_layout.addWidget(self.tbl, 1)
            # 19.3 Панель добавления
            add_panel = QtWidgets.QHBoxLayout()
            add_panel.setSpacing(6)
            # Количество
            self.sp_qty = SmartDoubleSpinBox()
            self.sp_qty.setDecimals(3)
            self.sp_qty.setMinimum(0.001)
            self.sp_qty.setValue(1.0)
            add_panel.addWidget(QtWidgets.QLabel("Кол-во:"))
            add_panel.addWidget(self.sp_qty)
            # Коэффициент
            self.sp_coeff = SmartDoubleSpinBox()
            self.sp_coeff.setDecimals(3)
            self.sp_coeff.setMinimum(0.001)
            self.sp_coeff.setValue(1.0)
            add_panel.addWidget(QtWidgets.QLabel("Коэф.:"))
            add_panel.addWidget(self.sp_coeff)
            # Зона
            self.cmb_zone = QtWidgets.QComboBox()
            self.cmb_zone.setMinimumWidth(160)
            add_panel.addWidget(QtWidgets.QLabel("Зона:"))
            add_panel.addWidget(self.cmb_zone)
            add_panel.addStretch(1)
            # Кнопки
            self.btn_add = QtWidgets.QPushButton("Добавить")
            self.btn_cancel = QtWidgets.QPushButton("Отмена")
            add_panel.addWidget(self.btn_add)
            add_panel.addWidget(self.btn_cancel)
            v_layout.addLayout(add_panel)
            # 19.4 Заполняем фильтры и зону
            self._fill_filters()
            self._fill_zones()
            # 19.5 Подключаем сигналы
            self.ed_search.textChanged.connect(self._update_table)
            self.cmb_vendor.currentIndexChanged.connect(self._update_table)
            self.cmb_department.currentIndexChanged.connect(self._update_table)
            self.btn_add.clicked.connect(self._on_add)
            self.btn_cancel.clicked.connect(self.reject)
            # 19.6 Загружаем таблицу
            self._update_table()

        def _fill_filters(self) -> None:
            """Заполняет комбобоксы подрядчиков и отделов."""
            try:
                vendors = self.page.db.catalog_distinct_values("vendor")
                departments = self.page.db.catalog_distinct_values("department")
            except Exception as ex:
                self.page._log(f"Ошибка загрузки фильтров каталога: {ex}", "error")
                vendors = []
                departments = []
            # Подрядчики
            self.cmb_vendor.blockSignals(True)
            self.cmb_vendor.clear()
            self.cmb_vendor.addItem("<Любой>", None)
            for v in vendors:
                if v:
                    self.cmb_vendor.addItem(normalize_case(v), v)
            self.cmb_vendor.setCurrentIndex(0)
            self.cmb_vendor.blockSignals(False)
            # Отделы
            self.cmb_department.blockSignals(True)
            self.cmb_department.clear()
            self.cmb_department.addItem("<Любой>", None)
            for d in departments:
                if d:
                    self.cmb_department.addItem(normalize_case(d), d)
            self.cmb_department.setCurrentIndex(0)
            self.cmb_department.blockSignals(False)

        def _fill_zones(self) -> None:
            """Заполняет список зон из текущего проекта."""
            zones = self.page.db.project_distinct_values(self.page.project_id, "zone") or []
            self.cmb_zone.clear()
            self.cmb_zone.addItem("Без зоны", "")
            for z in zones:
                if z:
                    self.cmb_zone.addItem(z, z)
            self.cmb_zone.setCurrentIndex(0)

        def _update_table(self) -> None:
            """
            Обновляет таблицу каталога согласно фильтрам.

            Поиск по наименованию осуществляется без учёта регистра: введённый
            текст нормализуется с помощью ``normalize_case`` перед передачей
            в ``catalog_list``, что обеспечивает регистронезависимый поиск
            даже при наличии символов кириллицы или латиницы. Это улучшает
            работу фильтра «Поиск» в окне выбора позиции из базы.
            """
            # Приводим поисковую строку к каноническому виду для нечувствительного поиска
            raw_name = self.ed_search.text().strip()
            name = normalize_case(raw_name) if raw_name else ""
            vendor = self.cmb_vendor.currentData()
            dept = self.cmb_department.currentData()
            filters: Dict[str, Any] = {}
            if name:
                filters["name"] = name
            if vendor:
                filters["vendor"] = vendor
            if dept:
                filters["department"] = dept
            rows: List[Any] = []
            try:
                rows = self.page.db.catalog_list(filters)
            except Exception as ex:
                # Логируем ошибку запроса каталога
                try:
                    self.page._log(f"Ошибка запроса каталога: {ex}", "error")
                except Exception:
                    pass
                rows = []
            # Заполняем таблицу
            self.tbl.setRowCount(0)
            for r in rows:
                row_idx = self.tbl.rowCount()
                self.tbl.insertRow(row_idx)
                # Нормализуем строки для отображения без учёта регистра
                name_norm = normalize_case(r["name"] or "")
                class_ru = CLASS_EN2RU.get((r["class"] or "equipment"), "Оборудование")
                vendor_norm = normalize_case(r["vendor"] or "")
                price = float(r["unit_price"] or 0.0)
                power = float(r["power_watts"] or 0.0)
                dept_norm = normalize_case(r["department"] or "")
                values = [
                    name_norm,
                    class_ru,
                    vendor_norm,
                    fmt_num(price, 2),
                    fmt_num(power, 0),
                    dept_norm,
                ]
                for col, val in enumerate(values):
                    item = QtWidgets.QTableWidgetItem(str(val))
                    if col == 0:
                        # Сохраняем оригинальные данные в первом столбце
                        item.setData(QtCore.Qt.UserRole, dict(r))
                    self.tbl.setItem(row_idx, col, item)
            # Подстраиваем ширину колонок под содержимое
            try:
                self.tbl.resizeColumnsToContents()
            except Exception:
                pass

        def _on_add(self) -> None:
            """Добавляет выбранную строку из каталога в проект.

            Эта версия сначала пытается найти в сводной смете существующую позицию
            с теми же параметрами (имя, подрядчик, отдел, класс, цена, коэффициент
            и зона). Если дубликат найден, увеличивает количество и сумму. Иначе
            сохраняет новую запись через add_items_bulk. В обоих случаях выводит
            информацию в лог и фиксирует действие для UNDO.
            """
            row_idx = self.tbl.currentRow()
            if row_idx < 0:
                QtWidgets.QMessageBox.information(self, "Внимание", "Выберите позицию для добавления.")
                return
            item = self.tbl.item(row_idx, 0)
            if not item:
                return
            # 19.7 Получаем данные выбранной строки каталога
            data = item.data(QtCore.Qt.UserRole)
            if not data:
                return
            try:
                # Нормализуем имя, подрядчика, отдел и класс
                name = normalize_case(data.get("name", ""))
                vendor = normalize_case(data.get("vendor", ""))
                department = normalize_case(data.get("department", ""))
                class_en = data.get("class", "equipment")
                price = float(data.get("unit_price", 0.0) or 0.0)
                power = float(data.get("power_watts", 0.0) or 0.0)
            except Exception as ex:
                # Логируем и информируем пользователя об ошибке разбора
                self.page._log(f"Ошибка чтения данных позиции: {ex}", "error")
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось прочитать данные: {ex}")
                return
            # 19.8 Читаем количество и коэффициент из спинов
            qty = float(self.sp_qty.value() or 1.0)
            coeff = float(self.sp_coeff.value() or 1.0)
            # Зона
            zone_data = self.cmb_zone.currentData()
            zone_raw = zone_data if zone_data is not None else (self.cmb_zone.currentText() or "")
            zone = normalize_case(zone_raw)
            # По умолчанию предполагаем создание новой записи
            duplicate = None
            try:
                # Получаем существующие позиции в этой зоне и классе
                existing_rows = self.page.db.list_items_filtered(
                    project_id=self.page.project_id,
                    vendor="<ALL>",
                    department="<ALL>",
                    zone=zone,
                    class_en=class_en,
                    name_like=name
                )
                for r in existing_rows:
                    try:
                        # Сравниваем по нормализованным полям. Используем индексированный доступ,
                        # поскольку sqlite3.Row не поддерживает метод get(). Пустые значения
                        # приводим к пустой строке или нулю.
                        if normalize_case(r["name"] or "") == name \
                           and normalize_case(r["vendor"] or "") == vendor \
                           and normalize_case(r["department"] or "") == department \
                           and (r["type"] or "equipment") == class_en \
                           and abs(float((r["unit_price"] or 0)) - price) < 1e-6 \
                           and abs(float((r["coeff"] or 0)) - coeff) < 1e-6:
                            duplicate = r
                            break
                    except Exception:
                        continue
            except Exception as ex:
                self.page._log(f"Ошибка поиска дубликатов в сводной смете: {ex}", "error")
            if duplicate:
                # 19.9 Нашли дубликат: увеличиваем количество и сумму
                try:
                    old_qty = float(duplicate["qty"] or 0.0)
                    new_qty = old_qty + qty
                    coeff_old = float(duplicate["coeff"] or 1.0)
                    price_old = float(duplicate["unit_price"] or 0.0)
                    new_amount = new_qty * coeff_old * price_old
                    self.page.db.update_item_fields(duplicate["id"], {"qty": new_qty, "amount": new_amount})
                    # Сохраняем данные для UNDO как редактирование
                    self.page._last_action = {
                        "type": "edit",
                        "item_id": duplicate["id"],
                        "old": {
                            "qty": old_qty,
                            "coeff": coeff_old,
                            "unit_price": price_old,
                            "amount": float(duplicate["amount"] or 0.0),
                        },
                    }
                    self.page.btn_undo_summary.setEnabled(True)
                    # Выводим информацию в лог
                    self.page._log(
                        f"Позиция «{name}» уже есть в зоне '{zone or 'Без зоны'}', количество увеличено на {fmt_num(qty,3)}."
                    )
                except Exception as ex:
                    self.page._log(f"Ошибка обновления количества позиции: {ex}", "error")
                    QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось обновить позицию: {ex}")
                    return
            else:
                # 19.10 Дубликат не найден — создаём новую запись
                amount = qty * coeff * price
                undo_batch = f"__undo_catalog_add__{datetime.utcnow().isoformat()}"
                try:
                    self.page.db.add_items_bulk([
                        {
                            "project_id": self.page.project_id,
                            "type": class_en,
                            "group_name": "Каталог",
                            "name": name,
                            "qty": qty,
                            "coeff": coeff,
                            "amount": amount,
                            "unit_price": price,
                            "source_file": None,
                            "vendor": vendor,
                            "department": department,
                            "zone": zone,
                            "power_watts": power,
                            "import_batch": undo_batch,
                        }
                    ])
                    # Фиксируем действие для UNDO как добавление
                    self.page._last_action = {
                        "type": "manual_add",
                        "project_id": self.page.project_id,
                        "batch": undo_batch,
                    }
                    self.page.btn_undo_summary.setEnabled(True)
                    self.page._log(
                        f"Добавлена позиция из каталога: «{name}», qty={fmt_num(qty,3)}, coeff={fmt_num(coeff,3)}, "
                        f"price={fmt_num(price,2)}, power={fmt_num(power,0)} Вт, зона='{zone or 'Без зоны'}'."
                    )
                except Exception as ex:
                    self.page._log(f"Ошибка добавления позиции: {ex}", "error")
                    QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось добавить позицию: {ex}")
                    return
            # 19.11 Обновляем таблицы сводной сметы и закрываем диалог
            reload_zone_tabs(self.page)
            self.accept()

    # Создаём и отображаем диалог
    dlg = CatalogDialog(page)
    dlg.exec()


# 20. Мастер добавления LED‑экрана
def open_screen_master(page: Any) -> None:
    """Открывает мастер создания LED‑экрана.

    Диалог позволяет ввести размеры экрана в метрах, выбрать цену за
    квадратный метр (вручную либо из каталога), определить подрядчика и
    отдел, а также автоматически рассчитывает площадь, количество
    кабинетов, разрешение, общее число пикселей и требуемое количество
    витых пар. Опционально пользователь может добавить к экрану
    соответствующие аксессуары: витую пару и видеопроцессор.

    После подтверждения создаются записи в проекте и (при необходимости)
    пополняется каталог. Мастер использует :class:`CatalogSelectDialog`
    для выбора позиции экрана из базы, чтобы подставить цену и метаданные.
    """
    # 20.1 Проверяем, открыт ли проект
    if getattr(page, "project_id", None) is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return

    # 20.2 Логируем открытие мастера
    try:
        if hasattr(page, "_log"):
            page._log("Мастер LED‑экрана: открыт диалог.")
    except Exception:
        pass

    class ScreenMasterDialog(QtWidgets.QDialog):
        """Внутренний диалог для ввода параметров экрана."""

        def __init__(self, parent_page: Any):  # type: ignore
            super().__init__(parent_page)
            self.page = parent_page
            self.setWindowTitle("Добавление LED‑экрана")
            self.resize(400, 350)
            layout = QtWidgets.QVBoxLayout(self)
            layout.setContentsMargins(8, 8, 8, 8)
            layout.setSpacing(6)

            form = QtWidgets.QFormLayout()
            form.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter)

            # Размеры экрана
            self.ed_width = QtWidgets.QDoubleSpinBox()
            self.ed_width.setDecimals(2)
            self.ed_width.setRange(0.1, 1000)
            self.ed_width.setValue(1.0)
            self.ed_width.setSuffix(" м")
            self.ed_height = QtWidgets.QDoubleSpinBox()
            self.ed_height.setDecimals(2)
            self.ed_height.setRange(0.1, 1000)
            self.ed_height.setValue(1.0)
            self.ed_height.setSuffix(" м")
            form.addRow("Ширина:", self.ed_width)
            form.addRow("Высота:", self.ed_height)
            # ID экрана (номер) — позволяет различать несколько экранов в смете
            self.spin_id = QtWidgets.QSpinBox()
            self.spin_id.setRange(1, 99999)
            self.spin_id.setValue(1)
            form.addRow("Номер экрана:", self.spin_id)

            # Цена за м²
            self.spin_price = QtWidgets.QDoubleSpinBox()
            self.spin_price.setDecimals(2)
            self.spin_price.setRange(0.0, 1_000_000.0)
            self.spin_price.setValue(0.0)
            self.spin_price.setSuffix(" ₽/м²")
            form.addRow("Цена за м²:", self.spin_price)

            # Разрешение одного модуля (по умолчанию 128×128 пикселей)
            self.spin_mod_w = QtWidgets.QSpinBox()
            self.spin_mod_w.setRange(1, 4096)
            self.spin_mod_w.setValue(128)
            self.spin_mod_w.setSuffix(" px")
            self.spin_mod_h = QtWidgets.QSpinBox()
            self.spin_mod_h.setRange(1, 4096)
            self.spin_mod_h.setValue(128)
            self.spin_mod_h.setSuffix(" px")
            form.addRow("Пикселей в модуле (ширина):", self.spin_mod_w)
            form.addRow("Пикселей в модуле (высота):", self.spin_mod_h)

            # Подрядчик и отдел (editable comboboxes с существующими значениями)
            self.cmb_vendor = QtWidgets.QComboBox()
            self.cmb_vendor.setEditable(True)
            self.cmb_department = QtWidgets.QComboBox()
            self.cmb_department.setEditable(True)
            # Заполняем списки подрядчиков и отделов из каталога
            try:
                vendors = self.page.db.catalog_distinct_values("vendor")
                departments = self.page.db.catalog_distinct_values("department")
            except Exception:
                vendors, departments = [], []
            self.cmb_vendor.addItem("")
            for v in vendors:
                if v:
                    self.cmb_vendor.addItem(normalize_case(v))
            self.cmb_department.addItem("")
            for d in departments:
                if d:
                    self.cmb_department.addItem(normalize_case(d))
            form.addRow("Подрядчик:", self.cmb_vendor)
            form.addRow("Отдел:", self.cmb_department)

            layout.addLayout(form)

            # Вычисляемые параметры
            self.lbl_area = QtWidgets.QLabel()
            self.lbl_cabinets = QtWidgets.QLabel()
            self.lbl_resolution = QtWidgets.QLabel()
            self.lbl_pixels = QtWidgets.QLabel()
            self.lbl_cables = QtWidgets.QLabel()
            # Обновить надписи
            self._update_labels()
            layout.addWidget(self.lbl_area)
            layout.addWidget(self.lbl_cabinets)
            layout.addWidget(self.lbl_resolution)
            layout.addWidget(self.lbl_pixels)
            layout.addWidget(self.lbl_cables)

            # Чекбоксы для аксессуаров
            self.chk_cable = QtWidgets.QCheckBox("Добавить витую пару")
            self.chk_cable.setChecked(True)
            self.chk_vp = QtWidgets.QCheckBox("Добавить видеопроцессор")
            self.chk_vp.setChecked(True)
            layout.addWidget(self.chk_cable)
            layout.addWidget(self.chk_vp)
            # Чекбокс и цена для конструктива (рама/конструктив для установки экрана)
            self.chk_structure = QtWidgets.QCheckBox("Добавить конструктив")
            self.chk_structure.setChecked(False)
            self.spin_structure_price = QtWidgets.QDoubleSpinBox()
            self.spin_structure_price.setDecimals(2)
            self.spin_structure_price.setRange(0.0, 1_000_000.0)
            self.spin_structure_price.setValue(0.0)
            self.spin_structure_price.setSuffix(" ₽")
            # Размещаем конструктив и его цену в одной строке
            h_struct = QtWidgets.QHBoxLayout()
            h_struct.addWidget(self.chk_structure)
            h_struct.addWidget(self.spin_structure_price)
            layout.addLayout(h_struct)

            # Кнопка выбора позиции из базы (экран)
            self.btn_select = QtWidgets.QPushButton("Выбрать экран из базы…")
            layout.addWidget(self.btn_select)

            # Кнопки OK/Cancel
            btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            layout.addWidget(btns)

            # Связи сигналов
            self.ed_width.valueChanged.connect(self._update_labels)
            self.ed_height.valueChanged.connect(self._update_labels)
            self.spin_mod_w.valueChanged.connect(self._update_labels)
            self.spin_mod_h.valueChanged.connect(self._update_labels)
            self.btn_select.clicked.connect(self._choose_from_catalog)
            btns.accepted.connect(self.accept)
            btns.rejected.connect(self.reject)

            # Переменные для выбранного элемента каталога
            self.selected_catalog: Optional[Dict[str, Any]] = None

        def _update_labels(self) -> None:
            """Пересчитывает и отображает площадь, количество кабинетов, разрешение и пиксели.

            Расчёты выполняются согласно заданной геометрии:

            * Площадь = ширина × высота.
            * Количество кабинетов = ceil(ширина×2) × ceil(высота×2). Здесь
              предполагается, что один кабинет имеет габариты 0.5×0.5 метра, поэтому
              вдоль каждой стороны помещается ``ceil(dimension × 2)`` кабинетов.
            * Разрешение по ширине/высоте = количество кабинетов вдоль стороны × 128
              пикселей (каждый кабинет 128×128 пикселей).
            * Общее число пикселей = res_x × res_y.
            * Витых пар требуется ``ceil(total_pixels / 650000)`` (округляем вверх).
            """
            w = float(self.ed_width.value() or 0)
            h = float(self.ed_height.value() or 0)
            # Площадь
            area = w * h
            import math
            # Количество кабинетов: считаем по каждой стороне отдельно
            # Шаг модуля = 0.5 м (два модуля на метр)
            cab_w = max(1, math.ceil(w * 2))
            cab_h = max(1, math.ceil(h * 2))
            cabinets = cab_w * cab_h
            # Разрешение модуля (ширина×высота) задаёт разрешение каждого кабинета
            mod_px_w = int(self.spin_mod_w.value())
            mod_px_h = int(self.spin_mod_h.value())
            # Разрешение экрана = количество кабинетов × разрешение модуля
            res_x = cab_w * mod_px_w
            res_y = cab_h * mod_px_h
            total_pixels = res_x * res_y
            # Количество витых пар (минимум 1)
            cables = max(1, math.ceil(total_pixels / 650_000))
            # Отображаем
            self.lbl_area.setText(f"Площадь: {fmt_num(area, 2)} м²")
            self.lbl_cabinets.setText(f"Кабинетов: {cabinets}")
            self.lbl_resolution.setText(f"Разрешение: {res_x} × {res_y} пикселей")
            self.lbl_pixels.setText(f"Пикселей всего: {int(total_pixels)}")
            self.lbl_cables.setText(f"Витая пара: {cables} шт.")
            # Сохраняем параметры для использования при accept()
            self._area = area
            self._cabinets = cabinets
            self._res_x = res_x
            self._res_y = res_y
            self._total_pixels = total_pixels
            self._cables = cables

        def _choose_from_catalog(self) -> None:
            """Позволяет выбрать позицию экрана из базы и подставить её параметры."""
            dlg = CatalogSelectDialog(self.page, parent=self)
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.selected_row
            if not data:
                return
            self.selected_catalog = data
            # Подставляем цену и параметры
            try:
                price = float(data.get("unit_price") or 0.0)
                self.spin_price.setValue(price)
            except Exception:
                pass
            vendor = normalize_case(data.get("vendor") or "")
            dept = normalize_case(data.get("department") or "")
            # Устанавливаем текст комбобоксов (добавляем если отсутствует)
            def set_combo(combo: QtWidgets.QComboBox, text: str) -> None:
                if not text:
                    return
                # Ищем существующий индекс без учёта регистра
                for i in range(combo.count()):
                    if combo.itemText(i).lower() == text.lower():
                        combo.setCurrentIndex(i)
                        return
                # Не найдено — добавляем
                combo.addItem(text)
                combo.setCurrentIndex(combo.count() - 1)
            set_combo(self.cmb_vendor, vendor)
            set_combo(self.cmb_department, dept)
            # Обновляем подписи
            try:
                self._update_labels()
            except Exception:
                pass

    # Перед созданием диалога вычисляем предложенный номер экрана
    default_id = 1
    try:
        import re as _re_id
        cur_id = page.db._conn.cursor()
        cur_id.execute(
            "SELECT name FROM items WHERE project_id=? AND name LIKE 'LED экран%'",
            (page.project_id,),
        )
        rows_id = cur_id.fetchall()
        max_id_found = 0
        for r in rows_id:
            # r может быть sqlite3.Row или tuple
            nm = r["name"] if isinstance(r, dict) else r[0]
            m_id = _re_id.match(r"LED экран\s*(\d+)", nm)
            if m_id:
                try:
                    val_id = int(m_id.group(1))
                    if val_id > max_id_found:
                        max_id_found = val_id
                except Exception:
                    continue
        default_id = max_id_found + 1 if max_id_found >= 1 else 1
    except Exception:
        default_id = 1
    # Создаём и отображаем диалог
    dlg = ScreenMasterDialog(page)
    # Устанавливаем предложенный номер экрана
    try:
        dlg.spin_id.setValue(int(default_id))
    except Exception:
        pass
    if dlg.exec() != QtWidgets.QDialog.Accepted:
        return
    # Извлекаем введённые данные
    width = float(dlg.ed_width.value())
    height = float(dlg.ed_height.value())
    area = width * height
    price_per_m2 = float(dlg.spin_price.value())
    vendor = dlg.cmb_vendor.currentText().strip() or ""
    department = dlg.cmb_department.currentText().strip() or ""
    area_qty = area  # используем площадь как количество
    # Подготовка основной позиции (экран)
    # Формируем номер экрана (ID), если указан
    try:
        id_val = int(dlg.spin_id.value())
    except Exception:
        id_val = 1
    screen_name = (
        f"LED экран {id_val} {fmt_num(width, 2)}×{fmt_num(height, 2)} м "
        f"({dlg._cabinets} кабинетов, {dlg._res_x}×{dlg._res_y} пикселей)"
    )
    screen_unit_price = price_per_m2
    screen_amount = area_qty * screen_unit_price
    # Формируем имя группы для экрана и связанных аксессуаров.  Используем
    # номер экрана, чтобы объединить сам экран, витые пары и
    # видеопроцессор в одну логическую группу. Это упрощает
    # восприятие сметы: все элементы, относящиеся к конкретному экрану,
    # отображаются вместе. Префикс «LED экран №» выбран по аналогии со
    # сценическим подиумом и не влияет на вычисления.
    group_name_screen = f"LED экран №{id_val}"
    # Составляем список записей для проекта и каталог
    items_for_db: List[Dict[str, Any]] = []
    catalog_entries: List[Dict[str, Any]] = []
    # Запись для экрана
    items_for_db.append({
        "project_id": page.project_id,
        "type": "equipment",
        # Используем единый group_name для экрана и его аксессуаров.  Это поле
        # позволяет группировать связанные позиции в смете (например,
        # экран, витую пару и видеопроцессор).  Здесь и далее применяем
        # переменную group_name_screen вместо универсальной
        # «Аренда оборудования».
        "group_name": group_name_screen,
        "name": screen_name,
        "qty": area_qty,
        "coeff": 1.0,
        "amount": screen_amount,
        "unit_price": screen_unit_price,
        "source_file": "SCREEN_MASTER",
        "vendor": vendor,
        "department": department,
        "zone": "",
        "power_watts": 0.0,
        "import_batch": f"screen-{datetime.utcnow().isoformat()}"
    })
    catalog_entries.append({
        "name": screen_name,
        "unit_price": screen_unit_price,
        "class": "equipment",
        "vendor": vendor,
        "power_watts": 0.0,
        "department": department,
    })
    # Дополнительные аксессуары: витая пара
    if dlg.chk_cable.isChecked():
        # Пытаемся найти товар по названию "витая пара" в каталоге
        cable_rows = []
        try:
            cable_rows = page.db.catalog_list({"name": "витая пара"})
        except Exception:
            cable_rows = []
        if cable_rows:
            # sqlite3.Row поддерживает доступ по ключу, но не имеет метода get
            row0 = cable_rows[0]
            try:
                cable_price = float(row0["unit_price"] or 0.0)
            except Exception:
                cable_price = 0.0
            try:
                cable_vendor = normalize_case(row0["vendor"] or vendor)
            except Exception:
                cable_vendor = vendor
            try:
                cable_department = normalize_case(row0["department"] or department)
            except Exception:
                cable_department = department
        else:
            cable_price = 0.0
            cable_vendor = vendor
            cable_department = department
        cable_qty = max(1, dlg._cables)
        cable_name = f"Витая пара для LED {id_val} {fmt_num(width, 2)}×{fmt_num(height, 2)} м"
        items_for_db.append({
            "project_id": page.project_id,
            "type": "equipment",
            # Все аксессуары для LED‑экрана помещаем в ту же группу,
            # что и сам экран.  Это объединяет витые пары с экраном в
            # интерфейсе сводной сметы.
            "group_name": group_name_screen,
            "name": cable_name,
            "qty": float(cable_qty),
            "coeff": 1.0,
            "amount": cable_price * cable_qty,
            "unit_price": cable_price,
            "source_file": "SCREEN_MASTER",
            "vendor": cable_vendor,
            "department": cable_department,
            "zone": "",
            "power_watts": 0.0,
            "import_batch": f"screen-{datetime.utcnow().isoformat()}"
        })
        catalog_entries.append({
            "name": cable_name,
            "unit_price": cable_price,
            "class": "equipment",
            "vendor": cable_vendor,
            "power_watts": 0.0,
            "department": cable_department,
        })
    # Видеопроцессор
    if dlg.chk_vp.isChecked():
        vp_rows = []
        try:
            # ищем позицию по ключевому слову "процессор"
            vp_rows = page.db.catalog_list({"name": "процессор"})
        except Exception:
            vp_rows = []
        if vp_rows:
            row0 = vp_rows[0]
            try:
                vp_price = float(row0["unit_price"] or 0.0)
            except Exception:
                vp_price = 0.0
            try:
                vp_vendor = normalize_case(row0["vendor"] or vendor)
            except Exception:
                vp_vendor = vendor
            try:
                vp_department = normalize_case(row0["department"] or department)
            except Exception:
                vp_department = department
        else:
            vp_price = 0.0
            vp_vendor = vendor
            vp_department = department
        vp_name = f"Видеопроцессор для LED {id_val} {fmt_num(width, 2)}×{fmt_num(height, 2)} м"
        items_for_db.append({
            "project_id": page.project_id,
            "type": "equipment",
            # Видеопроцессор относится к тому же экрану, поэтому используем
            # общее имя группы для всего комплекта
            "group_name": group_name_screen,
            "name": vp_name,
            "qty": 1.0,
            "coeff": 1.0,
            "amount": vp_price,
            "unit_price": vp_price,
            "source_file": "SCREEN_MASTER",
            "vendor": vp_vendor,
            "department": vp_department,
            "zone": "",
            "power_watts": 0.0,
            "import_batch": f"screen-{datetime.utcnow().isoformat()}"
        })
        catalog_entries.append({
            "name": vp_name,
            "unit_price": vp_price,
            "class": "equipment",
            "vendor": vp_vendor,
            "power_watts": 0.0,
            "department": vp_department,
        })
    # Конструктив для установки LED‑экрана (добавляется, если выбран соответствующий флаг)
    # Конструктив представляет собой раму или набор крепежей для монтажа экрана.
    if dlg.chk_structure.isChecked():
        try:
            struct_price = float(dlg.spin_structure_price.value() or 0.0)
        except Exception:
            struct_price = 0.0
        struct_name = f"Конструктив для LED {id_val} {fmt_num(width, 2)}×{fmt_num(height, 2)} м"
        items_for_db.append({
            "project_id": page.project_id,
            "type": "equipment",
            "group_name": group_name_screen,
            "name": struct_name,
            "qty": 1.0,
            "coeff": 1.0,
            "amount": struct_price,
            "unit_price": struct_price,
            "source_file": "SCREEN_MASTER",
            # Для конструктивов используем те же подрядчика и отдел, что и для экрана
            "vendor": vendor,
            "department": department,
            "zone": "",
            "power_watts": 0.0,
            "import_batch": f"screen-{datetime.utcnow().isoformat()}"
        })
        catalog_entries.append({
            "name": struct_name,
            "unit_price": struct_price,
            "class": "equipment",
            "vendor": vendor,
            "power_watts": 0.0,
            "department": department,
        })

    # Запись в базу
    try:
        page.db.add_items_bulk(items_for_db)
        # Обновляем/добавляем в каталог
        if hasattr(page.db, "catalog_add_or_ignore"):
            page.db.catalog_add_or_ignore(catalog_entries)
        page._log(f"Мастер экрана: добавлено позиций {len(items_for_db)} (площадь {fmt_num(area_qty,2)} м², цена {fmt_num(price_per_m2,2)}).")
        QtWidgets.QMessageBox.information(page, "Готово", f"Экран и аксессуары добавлены ({len(items_for_db)} позиций).")
    except Exception as ex:
        page._log(f"Мастер экрана: ошибка добавления: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось добавить экран: {ex}")
        return
    # Обновляем таблицы сметы
    try:
        page._reload_zone_tabs()
    except Exception:
        pass

# 22. Универсальный мастер добавления
def open_master_addition(page: Any) -> None:
    """
    Открывает универсальный мастер добавления различных элементов в смету.

    Мастер привязан к текущей выбранной зоне. Он содержит кнопки для
    добавления LED‑экрана, колонок (аудиосистемы), коммутации, сценического
    подиума и технического директора. Каждая кнопка вызывает соответствующий
    специализированный мастер или выполняет расчёт и запись в базу данных.

    При добавлении экрана или колонок мастер автоматически назначает зону
    для только что созданных позиций: новые записи из мастеров экрана
    (``open_screen_master``) и колонок (``open_column_master``) записываются
    в базу с пустым полем ``zone``. Этот мастер обновляет поле ``zone``
    у вновь созданных записей на название активной зоны. Благодаря этому
    добавление становится контекстным — элементы попадают именно в ту зону,
    на вкладке которой находится пользователь.

    :param page: объект ``ProjectPage`` со свойствами ``db``, ``project_id``
                 и ``zone_tabs``.
    """
    # 22.1 Проверяем, открыт ли проект
    if getattr(page, "project_id", None) is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return
    # 22.2 Определяем активную зону. Если вкладка не выбрана, используем пустую строку.
    try:
        idx = page.zone_tabs.currentIndex()
        zone_name = page.zone_tabs.tabText(idx).strip() if idx >= 0 else ""
    except Exception:
        zone_name = ""
    # 22.3 Логируем открытие мастера
    try:
        if hasattr(page, "_log"):
            page._log(f"Универсальный мастер добавления: открыт диалог для зоны '{zone_name}'.")
    except Exception:
        pass
    class MasterAddDialog(QtWidgets.QDialog):
        """
        Внутренний диалог универсального мастера добавления.

        Отображает название зоны и пять кнопок для добавления различных
        сущностей: экран, колонки, коммутация, подиум и технический директор.
        Каждая кнопка вызывает соответствующий метод, который выполняет
        расчёты и добавляет записи в БД. После успешного добавления
        происходит перезагрузка таблиц сметы.
        """
        def __init__(self, parent_page: Any, zone: str) -> None:  # type: ignore
            super().__init__(parent_page)
            self.page = parent_page
            self.zone_name = zone
            self.setWindowTitle("Мастер добавления")
            self.resize(400, 300)
            # Основная вертикальная компоновка для диалога
            v = QtWidgets.QVBoxLayout(self)
            v.setContentsMargins(8, 8, 8, 8)
            v.setSpacing(6)
            # Создаём вкладки: первая вкладка содержит основные кнопки, вторая — настройки
            tabs = QtWidgets.QTabWidget()
            # ---- Первая вкладка: основные действия ----
            tab_main = QtWidgets.QWidget(); layout_main = QtWidgets.QVBoxLayout(tab_main)
            layout_main.setContentsMargins(4, 4, 4, 4)
            layout_main.setSpacing(6)
            lbl_zone = QtWidgets.QLabel()
            if self.zone_name:
                lbl_zone.setText(f"Текущая зона: {self.zone_name}")
            else:
                lbl_zone.setText("Текущая зона не выбрана")
            layout_main.addWidget(lbl_zone)
            btn_screen = QtWidgets.QPushButton("Добавить экран")
            btn_column = QtWidgets.QPushButton("Добавить колонки")
            btn_commut = QtWidgets.QPushButton("Добавить коммутацию")
            btn_stage = QtWidgets.QPushButton("Добавить подиум")
            btn_director = QtWidgets.QPushButton("Добавить тех. директора")
            btn_screen.clicked.connect(self._add_screen)
            btn_column.clicked.connect(self._add_column)
            btn_commut.clicked.connect(self._add_commutation)
            btn_stage.clicked.connect(self._add_stage)
            btn_director.clicked.connect(self._add_director)
            for b in (btn_screen, btn_column, btn_commut, btn_stage, btn_director):
                layout_main.addWidget(b)
            layout_main.addStretch(1)
            # ---- Вторая вкладка: настройки ----
            tab_settings = QtWidgets.QWidget(); layout_settings = QtWidgets.QVBoxLayout(tab_settings)
            layout_settings.setContentsMargins(4, 4, 4, 4)
            layout_settings.setSpacing(6)
            # Пока вторая вкладка содержит заглушку. Здесь в будущем будут преднастройки кнопок мастера.
            placeholder = QtWidgets.QLabel(
                "Настройки предустановок для кнопок мастера будут реализованы здесь.\n"
                "Например, выбор витых пар и процессоров для LED‑экрана,\n"
                "галочка \"Добавить конструктив\" и другие параметры."
            )
            placeholder.setWordWrap(True)
            layout_settings.addWidget(placeholder)
            layout_settings.addStretch(1)
            # Добавляем вкладки в TabWidget
            tabs.addTab(tab_main, "Добавление")
            tabs.addTab(tab_settings, "Настройки")
            # Добавляем TabWidget на основной layout
            v.addWidget(tabs)
            # Кнопка закрытия находится под вкладками
            btn_close = QtWidgets.QPushButton("Закрыть")
            btn_close.clicked.connect(self.reject)
            v.addWidget(btn_close, alignment=QtCore.Qt.AlignRight)
        def _add_screen(self) -> None:
            # Сохраняем максимальный id перед вызовом мастера
            try:
                cur = self.page.db._conn.cursor()
                row = cur.execute("SELECT COALESCE(MAX(id),0) FROM items").fetchone()
                max_id_before = int(row[0]) if row and row[0] is not None else 0
            except Exception:
                max_id_before = 0
            open_screen_master(self.page)
            try:
                cur = self.page.db._conn.cursor()
                cur.execute(
                    "SELECT id FROM items WHERE id>? AND project_id=? AND COALESCE(source_file,'')='SCREEN_MASTER'",
                    (max_id_before, self.page.project_id),
                )
                rows = cur.fetchall()
                for r in rows:
                    try:
                        it_id = int(r[0])
                        self.page.db.update_item_field(it_id, "zone", self.zone_name or "")
                    except Exception:
                        continue
                if rows:
                    try:
                        if hasattr(self.page, "_log"):
                            self.page._log(f"Мастер добавления: экран добавлен в зону '{self.zone_name}'.")
                    except Exception:
                        pass
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка при назначении зоны экрана: {ex}", "error")
                except Exception:
                    pass
            try:
                self.page._reload_zone_tabs()
            except Exception:
                pass
            self.accept()
        def _add_column(self) -> None:
            # Перед добавлением колонок запрашиваем подрядчика. Если пользователь
            # отменил ввод, действие прерываем. Отдел для колонок всегда "звук".
            vendor_name, ok = QtWidgets.QInputDialog.getText(
                self, "Подрядчик колонок", "Введите название подрядчика для аудиосистемы:",
                text=""
            )
            if not ok:
                return
            vendor_name = normalize_case(vendor_name.strip()) if vendor_name else ""
            try:
                cur = self.page.db._conn.cursor()
                row = cur.execute("SELECT COALESCE(MAX(id),0) FROM items").fetchone()
                max_id_before = int(row[0]) if row and row[0] is not None else 0
            except Exception:
                max_id_before = 0
            open_column_master(self.page)
            try:
                cur = self.page.db._conn.cursor()
                cur.execute(
                    "SELECT id FROM items WHERE id>? AND project_id=? AND COALESCE(source_file,'')='COLUMN_MASTER'",
                    (max_id_before, self.page.project_id),
                )
                rows = cur.fetchall()
                for r in rows:
                    try:
                        it_id = int(r[0])
                        # Обновляем зону, подрядчика и отдел у каждой новой записи аудиосистемы
                        self.page.db.update_item_fields(it_id, {
                            "zone": self.zone_name or "",
                            "vendor": vendor_name,
                            "department": normalize_case("звук")
                        })
                    except Exception:
                        continue
                if rows:
                    try:
                        if hasattr(self.page, "_log"):
                            self.page._log(
                                f"Мастер добавления: колонки добавлены в зону '{self.zone_name}'"
                                f" с подрядчиком '{vendor_name or 'не указан'}' и отделом 'звук'."
                            )
                    except Exception:
                        pass
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка при назначении зоны колонок: {ex}", "error")
                except Exception:
                    pass
            try:
                self.page._reload_zone_tabs()
            except Exception:
                pass
            self.accept()
        def _add_commutation(self) -> None:
            # Запрашиваем подрядчика для коммутации. Если пользователь отменил ввод — выход.
            vendor_name, ok = QtWidgets.QInputDialog.getText(
                self, "Подрядчик коммутации", "Введите название подрядчика для коммутации:",
                text=""
            )
            if not ok:
                return
            vendor_name = normalize_case(vendor_name.strip()) if vendor_name else ""
            try:
                cur = self.page.db._conn.cursor()
                cur.execute(
                    "SELECT COALESCE(SUM(amount),0) FROM items WHERE project_id=? AND COALESCE(zone,'')=? AND type='equipment'",
                    (self.page.project_id, self.zone_name or ""),
                )
                row = cur.fetchone()
                total_equipment = float(row[0] or 0.0)
            except Exception as ex:
                total_equipment = 0.0
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка чтения суммы оборудования: {ex}", "error")
                except Exception:
                    pass
            comm_sum = total_equipment * 0.015
            if comm_sum <= 0:
                QtWidgets.QMessageBox.information(self, "Внимание", "В выбранной зоне нет оборудования класса 'оборудование' для расчёта коммутации.")
                return
            import datetime
            batch = f"commutation-{datetime.datetime.utcnow().isoformat()}"
            item = {
                "project_id": self.page.project_id,
                "type": "other",
                "group_name": "Коммутация",
                "name": "Коммутация",
                "qty": 1.0,
                "coeff": 1.0,
                "amount": comm_sum,
                "unit_price": comm_sum,
                "source_file": "COMMUTATION_MASTER",
                "vendor": vendor_name,
                "department": "",
                "zone": self.zone_name or "",
                "power_watts": 0.0,
                "import_batch": batch,
            }
            try:
                self.page.db.add_items_bulk([item])
                if hasattr(self.page.db, "catalog_add_or_ignore"):
                    self.page.db.catalog_add_or_ignore([
                        {
                            "name": item["name"],
                            "unit_price": item["unit_price"],
                            "class": "other",
                            "vendor": "",
                            "power_watts": 0.0,
                            "department": "",
                        }
                    ])
                if hasattr(self.page, "_log"):
                    self.page._log(
                        f"Мастер добавления: добавлена коммутация в зону '{self.zone_name}' на сумму {fmt_num(comm_sum,2)}."
                    )
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка при добавлении коммутации: {ex}", "error")
                except Exception:
                    pass
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось добавить коммутацию: {ex}")
                return
            try:
                self.page._reload_zone_tabs()
            except Exception:
                pass
            self.accept()
        def _add_stage(self) -> None:
            try:
                open_stage_master(self.page, self.zone_name or "")
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка при открытии мастера подиума: {ex}", "error")
                except Exception:
                    pass
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось открыть мастер подиума: {ex}")
                return
            self.accept()
        def _add_director(self) -> None:
            try:
                total = float(self.page.db.project_total(self.page.project_id))
            except Exception:
                total = 0.0
            if total <= 0.0:
                QtWidgets.QMessageBox.information(self, "Внимание", "Сумма проекта равна нулю, нечего начислять.")
                return
            director_amount = total * 0.10
            import datetime
            batch = f"techdir-{datetime.datetime.utcnow().isoformat()}"
            # При добавлении тех. директора создаём позицию в зоне "Техдирекция" с подрядчиком "техдиректор"
            default_zone = "Техдирекция"
            default_vendor = normalize_case("техдиректор")
            item = {
                "project_id": self.page.project_id,
                "type": "other",
                "group_name": "Технический директор",
                "name": "Технический директор",
                "qty": 1.0,
                "coeff": 1.0,
                "amount": director_amount,
                "unit_price": director_amount,
                "source_file": "TECHDIR_MASTER",
                "vendor": default_vendor,
                "department": "",
                "zone": default_zone,
                "power_watts": 0.0,
                "import_batch": batch,
            }
            try:
                self.page.db.add_items_bulk([item])
                if hasattr(self.page.db, "catalog_add_or_ignore"):
                    self.page.db.catalog_add_or_ignore([
                        {
                            "name": item["name"],
                            "unit_price": item["unit_price"],
                            "class": "other",
                            "vendor": "",
                            "power_watts": 0.0,
                            "department": "",
                        }
                    ])
                if hasattr(self.page, "_log"):
                    self.page._log(
                        f"Мастер добавления: добавлен технический директор на сумму {fmt_num(director_amount,2)} (10% от {fmt_num(total,2)})."
                    )
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер добавления: ошибка при добавлении технического директора: {ex}", "error")
                except Exception:
                    pass
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось добавить технического директора: {ex}")
                return
            try:
                self.page._reload_zone_tabs()
            except Exception:
                pass
            self.accept()
    dlg = MasterAddDialog(page, zone_name)
    dlg.exec()

# 23. Мастер добавления сценического подиума
def open_stage_master(page: Any, zone_name: str) -> None:
    """
    Открывает мастер добавления сценического подиума.

    Пользователь задаёт размеры подиума (ширина, глубина, высота), количество
    ступенек, цены за модули различных размеров, цену ступеньки и цену
    одной ножки, а также может включить режим «шип‑паз», который
    изменяет расчёт количества ножек. На основе введённых данных
    мастер определяет, сколько модулей каждого типа требуется для
    указанной сцены, сколько ступенек и ножек нужно, и добавляет
    соответствующие позиции в смету, группируя их по зоне.

    :param page: объект ProjectPage
    :param zone_name: имя зоны, в которую добавляется сцена
    """
    if getattr(page, "project_id", None) is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return
    try:
        if hasattr(page, "_log"):
            page._log(f"Мастер подиума: открыт диалог для зоны '{zone_name}'.")
    except Exception:
        pass
    class StageMasterDialog(QtWidgets.QDialog):
        def __init__(self, parent_page: Any, zone: str) -> None:  # type: ignore
            super().__init__(parent_page)
            self.page = parent_page
            self.zone_name = zone
            self.setWindowTitle("Добавление сценического подиума")
            self.resize(460, 360)
            layout = QtWidgets.QVBoxLayout(self)
            layout.setContentsMargins(8, 8, 8, 8)
            layout.setSpacing(6)
            form = QtWidgets.QFormLayout()
            form.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter)
            # Размеры подиума (ширина, глубина, высота)
            self.ed_width = QtWidgets.QDoubleSpinBox()
            self.ed_width.setDecimals(2)
            self.ed_width.setRange(0.5, 100.0)
            self.ed_width.setValue(2.0)
            self.ed_width.setSuffix(" м")
            self.ed_depth = QtWidgets.QDoubleSpinBox()
            self.ed_depth.setDecimals(2)
            self.ed_depth.setRange(0.5, 100.0)
            self.ed_depth.setValue(2.0)
            self.ed_depth.setSuffix(" м")
            self.ed_height = QtWidgets.QDoubleSpinBox()
            self.ed_height.setDecimals(0)
            self.ed_height.setRange(10.0, 300.0)
            self.ed_height.setValue(80.0)
            self.ed_height.setSuffix(" см")
            form.addRow("Ширина:", self.ed_width)
            form.addRow("Глубина:", self.ed_depth)
            form.addRow("Высота:", self.ed_height)

            # Номер подиума. Пользователь может указать любой номер,
            # по умолчанию предлагается следующий после существующих.
            self.spin_stage = QtWidgets.QSpinBox()
            self.spin_stage.setRange(1, 9999)
            # Определяем предложенный номер: ищем существующие подиумы и берём +1
            try:
                cur_id = self.page.db._conn.cursor()
                cur_id.execute(
                    "SELECT name FROM items WHERE project_id=? AND COALESCE(zone,'')=? AND name LIKE 'Сценический подиум №%'",
                    (self.page.project_id, self.zone_name or ""),
                )
                rows_id = cur_id.fetchall()
                import re as _re_stage
                max_stage = 0
                for r in rows_id:
                    nm = r[0] if not isinstance(r, dict) else r.get("name")
                    m = _re_stage.match(r"Сценический подиум №(\d+)", nm or "")
                    if m:
                        try:
                            val = int(m.group(1))
                            if val > max_stage:
                                max_stage = val
                        except Exception:
                            continue
                default_stage = max_stage + 1 if max_stage >= 1 else 1
            except Exception:
                default_stage = 1
            self.spin_stage.setValue(default_stage)
            form.addRow("Номер подиума:", self.spin_stage)

            # Параметры ковралина: чекбокс и цена за м²
            self.chk_carpet = QtWidgets.QCheckBox("Добавить ковралин")
            self.chk_carpet.setChecked(False)
            self.price_carpet = QtWidgets.QDoubleSpinBox()
            self.price_carpet.setDecimals(2)
            self.price_carpet.setRange(0.0, 1_000_000.0)
            self.price_carpet.setValue(0.0)
            self.price_carpet.setSuffix(" ₽/м²")
            carpet_layout = QtWidgets.QHBoxLayout()
            carpet_layout.addWidget(self.chk_carpet)
            carpet_layout.addWidget(self.price_carpet)
            form.addRow("Ковралин:", carpet_layout)

            # Параметры рауса: чекбокс и цена за погонный метр
            self.chk_raus = QtWidgets.QCheckBox("Добавить раус")
            self.chk_raus.setChecked(False)
            self.price_raus = QtWidgets.QDoubleSpinBox()
            self.price_raus.setDecimals(2)
            self.price_raus.setRange(0.0, 1_000_000.0)
            self.price_raus.setValue(0.0)
            self.price_raus.setSuffix(" ₽/м")
            raus_layout = QtWidgets.QHBoxLayout()
            raus_layout.addWidget(self.chk_raus)
            raus_layout.addWidget(self.price_raus)
            form.addRow("Раус:", raus_layout)
            self.spin_steps = QtWidgets.QSpinBox()
            self.spin_steps.setRange(0, 10)
            self.spin_steps.setValue(1)
            form.addRow("Ступенек:", self.spin_steps)
            self.price_2x1 = QtWidgets.QDoubleSpinBox()
            self.price_2x1.setDecimals(2)
            self.price_2x1.setRange(0.0, 1_000_000.0)
            self.price_2x1.setValue(0.0)
            self.price_2x1.setSuffix(" ₽")
            self.price_1x1 = QtWidgets.QDoubleSpinBox()
            self.price_1x1.setDecimals(2)
            self.price_1x1.setRange(0.0, 1_000_000.0)
            self.price_1x1.setValue(0.0)
            self.price_1x1.setSuffix(" ₽")
            self.price_1x0_5 = QtWidgets.QDoubleSpinBox()
            self.price_1x0_5.setDecimals(2)
            self.price_1x0_5.setRange(0.0, 1_000_000.0)
            self.price_1x0_5.setValue(0.0)
            self.price_1x0_5.setSuffix(" ₽")
            self.price_step = QtWidgets.QDoubleSpinBox()
            self.price_step.setDecimals(2)
            self.price_step.setRange(0.0, 1_000_000.0)
            self.price_step.setValue(0.0)
            self.price_step.setSuffix(" ₽")
            self.price_leg = QtWidgets.QDoubleSpinBox()
            self.price_leg.setDecimals(2)
            self.price_leg.setRange(0.0, 1_000_000.0)
            self.price_leg.setValue(0.0)
            self.price_leg.setSuffix(" ₽")
            form.addRow("Цена 2×1 м:", self.price_2x1)
            form.addRow("Цена 1×1 м:", self.price_1x1)
            form.addRow("Цена 1×0.5 м:", self.price_1x0_5)
            form.addRow("Цена ступеньки:", self.price_step)
            form.addRow("Цена одной ноги:", self.price_leg)
            self.chk_ship = QtWidgets.QCheckBox("Использовать шип‑паз (общие ноги)")
            self.chk_ship.setChecked(False)
            form.addRow("Режим ног:", self.chk_ship)

            # 23.a Выбор подрядчика для подиума
            # Создаём выпадающий список подрядчиков, чтобы пользователь мог указать,
            # от какого подрядчика заказывается сценический подиум. Список
            # формируется на основе уникальных подрядчиков в каталоге. Поле editable,
            # чтобы была возможность ввести нового подрядчика вручную.
            self.cmb_vendor = QtWidgets.QComboBox()
            self.cmb_vendor.setEditable(True)
            # Заполняем существующими подрядчиками из каталога
            vendors: list[str] = []
            try:
                if hasattr(self.page.db, "catalog_distinct_values"):
                    vendors = self.page.db.catalog_distinct_values("vendor") or []
            except Exception:
                vendors = []
            # Добавляем непустые имена подрядчиков
            for v in vendors:
                if v and v.strip():
                    if self.cmb_vendor.findText(v, QtCore.Qt.MatchFlag.MatchFixedString) < 0:
                        self.cmb_vendor.addItem(v)
            # По умолчанию подрядчик не выбран
            form.addRow("Подрядчик:", self.cmb_vendor)
            layout.addLayout(form)
            btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            layout.addWidget(btns)
            btns.accepted.connect(self.accept)
            btns.rejected.connect(self.reject)
            # accept method defined below at class level

        def accept(self) -> None:  # type: ignore
            """Собирает параметры подиума, рассчитывает необходимые элементы и добавляет их в смету."""
            # Сохраняем выбранного подрядчика. Если пользователь оставил поле пустым,
            # используем пустую строку. Это значение применяется ко всем позициям подиума.
            try:
                vendor_selected = self.cmb_vendor.currentText().strip()
            except Exception:
                vendor_selected = ""
            # Получаем основные размеры и параметры
            w = float(self.ed_width.value())
            d = float(self.ed_depth.value())
            steps = int(self.spin_steps.value())
            price_2x1 = float(self.price_2x1.value())
            price_1x1 = float(self.price_1x1.value())
            price_1x0_5 = float(self.price_1x0_5.value())
            price_step = float(self.price_step.value())
            price_leg = float(self.price_leg.value())
            use_ship = self.chk_ship.isChecked()
            # Номер сцены, указанный пользователем
            try:
                stage_id = int(self.spin_stage.value())
            except Exception:
                stage_id = 1
            # Сохраняем цены ковралина и рауса (если включены)
            carpet_enabled = self.chk_carpet.isChecked()
            try:
                price_carpet = float(self.price_carpet.value()) if carpet_enabled else 0.0
            except Exception:
                price_carpet = 0.0
            raus_enabled = self.chk_raus.isChecked()
            try:
                price_raus = float(self.price_raus.value()) if raus_enabled else 0.0
            except Exception:
                price_raus = 0.0
            # Разбиваем размеры на сегменты по 2, 1 и 0.5 метра
            segments_x: list[float] = []
            rem_w = w
            eps = 1e-6
            while rem_w > eps:
                if rem_w >= 2.0 - eps:
                    segments_x.append(2.0)
                    rem_w -= 2.0
                elif rem_w >= 1.0 - eps:
                    segments_x.append(1.0)
                    rem_w -= 1.0
                else:
                    segments_x.append(1.0)
                    rem_w = 0.0
            segments_y: list[float] = []
            rem_d = d
            while rem_d > eps:
                if rem_d >= 1.0 - eps:
                    segments_y.append(1.0)
                    rem_d -= 1.0
                elif rem_d >= 0.5 - eps:
                    segments_y.append(0.5)
                    rem_d -= 0.5
                else:
                    segments_y.append(0.5)
                    rem_d = 0.0
            # Считаем количество модулей каждого типа. Чтобы максимизировать количество модулей 2×1 м,
            # используем площадь сцены. Однотипные модули считают по правилу:
            # максимально заполняем площадь экрана модулями 2×1 (или 1×2), затем оставшаяся площадь
            # покрывается модулями 1×1, а остаток в 0.5 м² закрывается 1×0.5 м. Такой подход
            # позволяет, например, для сцены 5×3 м получить 7 модулей 2×1 и один модуль 1×1.
            total_area = w * d
            # округляем до ближайших 0.5 м², чтобы избежать накопления ошибок
            area_units = round(total_area * 2) / 2.0
            count_2x1 = int(area_units // 2.0)
            remaining_units = area_units - count_2x1 * 2.0
            count_1x1 = int(remaining_units // 1.0)
            remaining_units -= count_1x1 * 1.0
            # оставшуюся площадь переводим в количество модулей 1×0.5
            if remaining_units > 1e-6:
                count_1x0_5 = int(round(remaining_units / 0.5))
            else:
                count_1x0_5 = 0
            # Количество ножек: зависит от режима (шип‑паз) и наличия ступенек
            if use_ship:
                legs_count = (len(segments_x) + 1) * (len(segments_y) + 1) + steps * 4
            else:
                legs_count = 4 * (count_2x1 + count_1x1 + count_1x0_5 + steps)
            import datetime
            batch = f"stage-{datetime.datetime.utcnow().isoformat()}"
            items_for_db: list[Dict[str, Any]] = []
            catalog_entries: list[Dict[str, Any]] = []
            # Вспомогательная функция для добавления позиции в смету и каталог
            def add_item(name: str, qty: float, unit_price: float) -> None:
                amount = unit_price * qty
                items_for_db.append({
                    "project_id": self.page.project_id,
                    "type": "equipment",
                    "group_name": f"Сценический подиум №{stage_id}",
                    "name": name,
                    "qty": qty,
                    "coeff": 1.0,
                    "amount": amount,
                    "unit_price": unit_price,
                    "source_file": "STAGE_MASTER",
                    # Используем выбранного пользователем подрядчика
                    "vendor": vendor_selected,
                    "department": "",
                    "zone": self.zone_name or "",
                    "power_watts": 0.0,
                    "import_batch": batch,
                })
                catalog_entries.append({
                    "name": name,
                    "unit_price": unit_price,
                    "class": "equipment",
                    # Также добавляем подрядчика в глобальный каталог
                    "vendor": vendor_selected,
                    "power_watts": 0.0,
                    "department": "",
                })
            # Добавляем панели
            if count_2x1 > 0:
                add_item(f"Панель подиума 2×1 м №{stage_id}", count_2x1, price_2x1)
            if count_1x1 > 0:
                add_item(f"Панель подиума 1×1 м №{stage_id}", count_1x1, price_1x1)
            if count_1x0_5 > 0:
                add_item(f"Панель подиума 1×0.5 м №{stage_id}", count_1x0_5, price_1x0_5)
            # Ступеньки
            if steps > 0:
                add_item(f"Ступенька подиума №{stage_id}", steps, price_step)
            # Ножки подиума
            if legs_count > 0:
                try:
                    h_cm = int(float(self.ed_height.value()))
                except Exception:
                    h_cm = int(self.ed_height.value() or 0)
                add_item(f"Нога подиума {h_cm} см №{stage_id}", legs_count, price_leg)
            # Ковралин
            if carpet_enabled:
                area = w * d
                if area > 0:
                    add_item(f"Ковралин подиума №{stage_id}", area, price_carpet)
            # Раус: считаем переднюю ширину + две боковых глубины (задняя часть не учитывается)
            if raus_enabled:
                perimeter = w + 2.0 * d
                if perimeter > 0:
                    add_item(f"Раус подиума №{stage_id}", perimeter, price_raus)
            # Запись в базу и логирование
            try:
                if items_for_db:
                    self.page.db.add_items_bulk(items_for_db)
                    if hasattr(self.page.db, "catalog_add_or_ignore"):
                        self.page.db.catalog_add_or_ignore(catalog_entries)
                    if hasattr(self.page, "_log"):
                        self.page._log(
                            f"Мастер подиума: добавлено {len(items_for_db)} позиций (2×1={count_2x1}, 1×1={count_1x1}, 1×0.5={count_1x0_5}, ступенек={steps}, ног={legs_count}) в зону '{self.zone_name}'."
                        )
            except Exception as ex:
                try:
                    if hasattr(self.page, "_log"):
                        self.page._log(f"Мастер подиума: ошибка добавления позиций: {ex}", "error")
                except Exception:
                    pass
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось добавить подиум: {ex}")
                return
            # Обновляем вкладки зон
            try:
                self.page._reload_zone_tabs()
            except Exception:
                pass
            # Закрываем диалог
            super().accept()
    dlg = StageMasterDialog(page, zone_name)
    dlg.exec()


# 21. Мастер добавления аудиосистемы (колонок)
def open_column_master(page: Any) -> None:
    """
    Открывает мастер добавления комплекта звуковых колонок.

    Мастер позволяет выбрать тип колонок (Main PA, FrontFill, InFill, OutFill,
    SideFill, Delay, Custom), указать количество и стоимость топов и сабов
    (как из базы данных, так и вручную), выбрать тип системы (активная или
    пассивная) и режим усилителей (Fullrange или Biamp). При пассивной
    системе рассчитывается число усилителей и коммутаторов SpeakOn NL4
    (0,5 м и 15 м) по заданным правилам. Результатом работы мастера
    становится набор позиций, добавленных в сводную смету (проект) с
    соответствующим префиксом, количеством и стоимостью. При выборе
    позиций из каталога также пополняется глобальный каталог.
    """
    # 21.1 Проверяем, открыт ли проект
    if getattr(page, "project_id", None) is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return
    # 21.2 Логируем открытие мастера
    try:
        if hasattr(page, "_log"):
            page._log("Мастер колонок: открыт диалог.")
    except Exception:
        pass

    class ColumnMasterDialog(QtWidgets.QDialog):
        """
        Внутренний диалог мастера колонок.

        Содержит элементы управления для выбора типа колонок, топов, сабов,
        режима системы (активная/пассивная) и усилителей. Также рассчитывает
        количество усилителей и SpeakOn-коммутацию.
        """
        def __init__(self, parent_page: Any) -> None:  # type: ignore
            super().__init__(parent_page)
            self.page = parent_page
            self.setWindowTitle("Добавление колонок")
            self.resize(480, 520)
            v = QtWidgets.QVBoxLayout(self)
            v.setContentsMargins(8, 8, 8, 8)
            v.setSpacing(6)
            # 21.3 Тип колонок
            form = QtWidgets.QFormLayout()
            form.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter)
            self.cmb_position = QtWidgets.QComboBox()
            self.cmb_position.addItems([
                "Main PA", "FrontFill", "InFill", "OutFill", "SideFill", "Delay", "Custom"
            ])
            form.addRow("Тип колонок:", self.cmb_position)
            # 21.4 Раздел топов
            self.chk_top = QtWidgets.QCheckBox("Добавить топы")
            self.chk_top.setChecked(True)
            form.addRow(self.chk_top)
            # Имя топов и выбор из базы
            h_top = QtWidgets.QHBoxLayout()
            self.ed_top_name = QtWidgets.QLineEdit(); self.ed_top_name.setPlaceholderText("Выберите или введите топ")
            # Разрешаем ручной ввод названия топов
            self.ed_top_name.setReadOnly(False)
            self.btn_top_select = QtWidgets.QPushButton("Из базы…")
            h_top.addWidget(self.ed_top_name)
            h_top.addWidget(self.btn_top_select)
            form.addRow("Топы:", h_top)
            # Количество и цена топов
            self.sp_top_qty = SmartDoubleSpinBox(); self.sp_top_qty.setDecimals(2); self.sp_top_qty.setMinimum(0.0); self.sp_top_qty.setValue(0.0)
            self.sp_top_price = SmartDoubleSpinBox(); self.sp_top_price.setDecimals(2)
            # Разрешаем широкий диапазон значений цены, чтобы не ограничивать 99,99
            self.sp_top_price.setRange(0.0, 1_000_000_000.0)
            self.sp_top_price.setValue(0.0)
            # Скрываем стрелочки у поля цены топов, ввод только с клавиатуры
            self.sp_top_price.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            h_top_params = QtWidgets.QHBoxLayout(); h_top_params.addWidget(QtWidgets.QLabel("Кол-во:")); h_top_params.addWidget(self.sp_top_qty);
            h_top_params.addWidget(QtWidgets.QLabel("Цена/шт:")); h_top_params.addWidget(self.sp_top_price)
            form.addRow("Параметры топов:", h_top_params)
            # 21.5 Раздел сабов
            self.chk_sub = QtWidgets.QCheckBox("Добавить сабы")
            self.chk_sub.setChecked(False)
            form.addRow(self.chk_sub)
            h_sub = QtWidgets.QHBoxLayout()
            self.ed_sub_name = QtWidgets.QLineEdit(); self.ed_sub_name.setPlaceholderText("Выберите или введите саб")
            # Разрешаем ручной ввод названия сабов
            self.ed_sub_name.setReadOnly(False)
            self.btn_sub_select = QtWidgets.QPushButton("Из базы…")
            h_sub.addWidget(self.ed_sub_name)
            h_sub.addWidget(self.btn_sub_select)
            form.addRow("Сабы:", h_sub)
            # Количество и цена сабов
            self.sp_sub_qty = SmartDoubleSpinBox(); self.sp_sub_qty.setDecimals(2); self.sp_sub_qty.setMinimum(0.0); self.sp_sub_qty.setValue(0.0)
            self.sp_sub_price = SmartDoubleSpinBox(); self.sp_sub_price.setDecimals(2)
            # Разрешаем широкий диапазон для цены сабов
            self.sp_sub_price.setRange(0.0, 1_000_000_000.0)
            self.sp_sub_price.setValue(0.0)
            # Скрываем стрелочки у поля цены сабов
            self.sp_sub_price.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            h_sub_params = QtWidgets.QHBoxLayout(); h_sub_params.addWidget(QtWidgets.QLabel("Кол-во:")); h_sub_params.addWidget(self.sp_sub_qty);
            h_sub_params.addWidget(QtWidgets.QLabel("Цена/шт:")); h_sub_params.addWidget(self.sp_sub_price)
            form.addRow("Параметры сабов:", h_sub_params)
            # Чекбокс добавления коробочек для сабов
            self.chk_sub_boxes = QtWidgets.QCheckBox("Добавить коробочки для сабов (2 шт.)")
            self.chk_sub_boxes.setChecked(False)
            # Коробочки доступны только если сабы включены
            self.chk_sub_boxes.setEnabled(False)
            # Размещаем чекбокс с небольшим отступом справа
            h_sub_boxes = QtWidgets.QHBoxLayout()
            h_sub_boxes.addSpacing(20)
            h_sub_boxes.addWidget(self.chk_sub_boxes)
            form.addRow("", h_sub_boxes)
            # 21.6 Выбор системы: активная или пассивная
            self.grp_system = QtWidgets.QGroupBox("Тип системы")
            rb_layout = QtWidgets.QHBoxLayout(self.grp_system)
            self.rb_active = QtWidgets.QRadioButton("Активная")
            self.rb_passive = QtWidgets.QRadioButton("Пассивная")
            self.rb_active.setChecked(True)
            rb_layout.addWidget(self.rb_active); rb_layout.addWidget(self.rb_passive)
            form.addRow(self.grp_system)
            # 21.7 Группа настроек пассивной системы
            self.grp_passive = QtWidgets.QGroupBox("Параметры пассивной системы")
            self.grp_passive.setCheckable(False)
            self.grp_passive.setEnabled(False)
            passive_layout = QtWidgets.QFormLayout(self.grp_passive)
            # Режим усилителей: Fullrange / Biamp
            self.rb_fullrange = QtWidgets.QRadioButton("Fullrange")
            self.rb_biamp = QtWidgets.QRadioButton("Biamp")
            self.rb_fullrange.setChecked(True)
            mode_h = QtWidgets.QHBoxLayout(); mode_h.addWidget(self.rb_fullrange); mode_h.addWidget(self.rb_biamp)
            passive_layout.addRow("Режим усилителей:", mode_h)
            # Тип сабов: одиночные или сдвоенные
            self.chk_double_sub = QtWidgets.QCheckBox("Сдвоенные сабы (2x)")
            self.chk_double_sub.setChecked(False)
            passive_layout.addRow(self.chk_double_sub)
            # Усилитель: имя, выбор из базы, количество, цена
            amp_h_name = QtWidgets.QHBoxLayout()
            self.ed_amp_name = QtWidgets.QLineEdit(); self.ed_amp_name.setPlaceholderText("Выберите или введите усилитель")
            # Разрешаем ручной ввод названия усилителя
            self.ed_amp_name.setReadOnly(False)
            self.btn_amp_select = QtWidgets.QPushButton("Из базы…")
            amp_h_name.addWidget(self.ed_amp_name); amp_h_name.addWidget(self.btn_amp_select)
            passive_layout.addRow("Усилитель:", amp_h_name)
            amp_h_params = QtWidgets.QHBoxLayout()
            self.sp_amp_qty = SmartDoubleSpinBox(); self.sp_amp_qty.setDecimals(0); self.sp_amp_qty.setMinimum(0);
            self.sp_amp_qty.setValue(0)
            self.sp_amp_price = SmartDoubleSpinBox(); self.sp_amp_price.setDecimals(2)
            # Разрешаем широкий диапазон для цены усилителя
            self.sp_amp_price.setRange(0.0, 1_000_000_000.0)
            self.sp_amp_price.setValue(0.0)
            # Скрываем стрелочки у поля цены усилителя
            self.sp_amp_price.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            amp_h_params.addWidget(QtWidgets.QLabel("Кол-во:")); amp_h_params.addWidget(self.sp_amp_qty);
            amp_h_params.addWidget(QtWidgets.QLabel("Цена/шт:")); amp_h_params.addWidget(self.sp_amp_price)
            passive_layout.addRow("Параметры усилителя:", amp_h_params)
            # Информация о коммутаторах
            self.lbl_connectors = QtWidgets.QLabel()
            passive_layout.addRow("Коммутация:", self.lbl_connectors)
            # Добавляем группы на форму
            form.addRow(self.grp_passive)
            # 21.8 Кнопки OK / Cancel
            btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            # Собираем форму
            v.addLayout(form)
            v.addWidget(btns)
            # 21.9 Сигналы
            self.btn_top_select.clicked.connect(self._select_top)
            self.btn_sub_select.clicked.connect(self._select_sub)
            self.btn_amp_select.clicked.connect(self._select_amp)
            self.chk_top.toggled.connect(self._top_enable_changed)
            self.chk_sub.toggled.connect(self._sub_enable_changed)
            self.rb_active.toggled.connect(self._system_toggle)
            self.rb_passive.toggled.connect(self._system_toggle)
            # Режимы и параметры изменяют автоподсчёт
            self.sp_top_qty.valueChanged.connect(self._update_passive_calc)
            self.sp_sub_qty.valueChanged.connect(self._update_passive_calc)
            self.rb_fullrange.toggled.connect(self._update_passive_calc)
            self.rb_biamp.toggled.connect(self._update_passive_calc)
            self.chk_double_sub.toggled.connect(self._update_passive_calc)
            btns.accepted.connect(self.accept)
            btns.rejected.connect(self.reject)
            # Инициализируем состояние
            self._top_vendor = ""
            self._top_department = ""
            self._sub_vendor = ""
            self._sub_department = ""
            self._amp_vendor = ""
            self._amp_department = ""
            # Обновим коммутатор info
            self._update_passive_calc()

        # 21.9.1 Обработчики
        def _top_enable_changed(self, state: bool) -> None:
            # Включаем/отключаем элементы топов
            enabled = self.chk_top.isChecked()
            self.ed_top_name.setEnabled(enabled)
            self.btn_top_select.setEnabled(enabled)
            self.sp_top_qty.setEnabled(enabled)
            self.sp_top_price.setEnabled(enabled)
            # Если отключено — сбрасываем значения
            if not enabled:
                self.ed_top_name.clear(); self.sp_top_qty.setValue(0.0); self.sp_top_price.setValue(0.0)
            self._update_passive_calc()

        def _sub_enable_changed(self, state: bool) -> None:
            enabled = self.chk_sub.isChecked()
            self.ed_sub_name.setEnabled(enabled)
            self.btn_sub_select.setEnabled(enabled)
            self.sp_sub_qty.setEnabled(enabled)
            self.sp_sub_price.setEnabled(enabled)
            # Коробочки для сабов доступны только если сабы выбраны
            self.chk_sub_boxes.setEnabled(enabled)
            if not enabled:
                # Сброс состояния при отключении сабов
                self.chk_sub_boxes.setChecked(False)
            if not enabled:
                self.ed_sub_name.clear(); self.sp_sub_qty.setValue(0.0); self.sp_sub_price.setValue(0.0)
            self._update_passive_calc()

        def _system_toggle(self) -> None:
            # Переключение между активной и пассивной системой
            is_passive = self.rb_passive.isChecked()
            self.grp_passive.setEnabled(is_passive)
            self._update_passive_calc()

        def _update_passive_calc(self) -> None:
            """Пересчитывает количество усилителей и коммутаторов для пассивной системы."""
            # Не считаем, если активная система
            if not self.rb_passive.isChecked():
                self.sp_amp_qty.setValue(0)
                # Отображаем пустую строку
                self.lbl_connectors.setText("Н/Д для активной системы")
                return
            import math
            top_qty = float(self.sp_top_qty.value() or 0.0) if self.chk_top.isChecked() else 0.0
            sub_qty = float(self.sp_sub_qty.value() or 0.0) if self.chk_sub.isChecked() else 0.0
            # Определяем ёмкость усилителя для топов
            if self.rb_fullrange.isChecked():
                amp_top_cap = 8  # Fullrange: 4 плеча * 2 топа
            else:
                amp_top_cap = 4  # Biamp: 2 плеча * 2 топа
            amps_for_tops = math.ceil(top_qty / amp_top_cap) if top_qty > 0 else 0
            # Ёмкость усилителя для сабов
            if self.chk_double_sub.isChecked():
                amp_sub_cap = 4  # сдвоенные: 4 на усилитель
            else:
                amp_sub_cap = 8
            amps_for_subs = math.ceil(sub_qty / amp_sub_cap) if sub_qty > 0 else 0
            # Количество усилителей суммируется: топы и сабы обслуживаются отдельно
            amps_needed = amps_for_tops + amps_for_subs
            # Обновляем спин количества усилителей, если оно ещё не редактировалось пользователем (<=0)
            if self.sp_amp_qty.value() <= 0 or self.sp_amp_qty.value() < amps_needed:
                # Автоподстановка расчётного значения
                self.sp_amp_qty.setValue(float(amps_needed))
            # Расчёт коммутации: для топов и сабов используем распределение по двум сторонам.
            def calc_connectors(count: float) -> tuple[int, int]:
                """Возвращает (short, long) кабели для заданного количества элементов.

                Элементы распределяются на две стороны: side1 = ceil(n/2), side2 = floor(n/2).
                На каждой стороне пары формируются только внутри стороны (не перемешиваются).
                Каждый комплект пары использует 1 короткий (0,5 м) и 1 длинный (15 м). Оставшиеся
                одиночные элементы используют только 1 длинный кабель. Всегда требуется минимум
                2 длинных кабеля, если count > 0.
                """
                n = int(count)
                if n <= 0:
                    return (0, 0)
                # Распределяем на две стороны
                side1 = (n + 1) // 2  # ceil(n/2)
                side2 = n - side1  # floor(n/2)
                # Для каждой стороны считаем пары и одиночки
                pairs1 = side1 // 2
                singles1 = side1 % 2
                pairs2 = side2 // 2
                singles2 = side2 % 2
                short = pairs1 + pairs2
                long = pairs1 + singles1 + pairs2 + singles2
                # Минимум два длинных кабеля
                if long < 2:
                    long = 2
                return (short, long)

            # Топы
            if top_qty > 0:
                spk_top_05, spk_top_15 = calc_connectors(top_qty)
            else:
                spk_top_05 = spk_top_15 = 0
            # Сабы
            if self.chk_sub.isChecked() and sub_qty > 0:
                if self.chk_double_sub.isChecked():
                    # Для сдвоенных сабов: каждый саб требует 1 длинный кабель
                    spk_sub_05 = 0
                    spk_sub_15 = max(2, int(sub_qty))
                else:
                    spk_sub_05, spk_sub_15 = calc_connectors(sub_qty)
            else:
                spk_sub_05 = spk_sub_15 = 0
            # Формируем информационную строку
            total_short = spk_top_05 + spk_sub_05
            total_long = spk_top_15 + spk_sub_15
            # Если выбраны коробочки для сабов, добавляем по одному короткому кабелю на коробку
            try:
                if self.chk_sub.isChecked() and self.chk_sub_boxes.isChecked():
                    # Коробочек всегда две, по одному короткому кабелю каждая
                    total_short += 2
            except Exception:
                pass
            if total_short or total_long:
                self.lbl_connectors.setText(
                    f"Короткий спикон: {int(total_short)}, Длинный спикон: {int(total_long)}"
                )
            else:
                self.lbl_connectors.setText("Нет кабелей")

        # 21.9.2 Выбор топов из базы
        def _select_top(self) -> None:
            dlg = CatalogSelectDialog(self.page, parent=self)
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.selected_row
            if not data:
                return
            try:
                name = normalize_case(data.get("name", ""))
                price = float(data.get("unit_price", 0.0) or 0.0)
                self._top_vendor = normalize_case(data.get("vendor", "") or "")
                self._top_department = normalize_case(data.get("department", "") or "")
                self.ed_top_name.setText(name)
                self.sp_top_price.setValue(price)
                # Устанавливаем количество по умолчанию как 2 или 1 (если пусто)
                if self.sp_top_qty.value() <= 0.0:
                    self.sp_top_qty.setValue(2.0)
                self._update_passive_calc()
            except Exception:
                pass

        # 21.9.3 Выбор сабов из базы
        def _select_sub(self) -> None:
            dlg = CatalogSelectDialog(self.page, parent=self)
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.selected_row
            if not data:
                return
            try:
                name = normalize_case(data.get("name", ""))
                price = float(data.get("unit_price", 0.0) or 0.0)
                self._sub_vendor = normalize_case(data.get("vendor", "") or "")
                self._sub_department = normalize_case(data.get("department", "") or "")
                self.ed_sub_name.setText(name)
                self.sp_sub_price.setValue(price)
                if self.sp_sub_qty.value() <= 0.0:
                    self.sp_sub_qty.setValue(2.0)
                self._update_passive_calc()
            except Exception:
                pass

        # 21.9.4 Выбор усилителей из базы
        def _select_amp(self) -> None:
            dlg = CatalogSelectDialog(self.page, parent=self)
            if dlg.exec() != QtWidgets.QDialog.Accepted:
                return
            data = dlg.selected_row
            if not data:
                return
            try:
                name = normalize_case(data.get("name", ""))
                price = float(data.get("unit_price", 0.0) or 0.0)
                self._amp_vendor = normalize_case(data.get("vendor", "") or "")
                self._amp_department = normalize_case(data.get("department", "") or "")
                self.ed_amp_name.setText(name)
                self.sp_amp_price.setValue(price)
                if self.sp_amp_qty.value() <= 0.0:
                    self.sp_amp_qty.setValue(1.0)
            except Exception:
                pass

    # 21.10 Создаём и отображаем диалог
    dlg = ColumnMasterDialog(page)
    if dlg.exec() != QtWidgets.QDialog.Accepted:
        return
    # 21.11 Собираем данные из диалога
    prefix = dlg.cmb_position.currentText().strip()
    # Готовим списки для добавления
    items_for_db: List[Dict[str, Any]] = []
    catalog_entries: List[Dict[str, Any]] = []
    # 21.12 Обработка топов
    if dlg.chk_top.isChecked() and dlg.ed_top_name.text().strip():
        try:
            top_name = dlg.ed_top_name.text().strip()
            top_qty = float(dlg.sp_top_qty.value() or 0.0)
            top_price = float(dlg.sp_top_price.value() or 0.0)
            if top_qty > 0.0:
                full_name = f"{prefix} {top_name}"
                amount = top_qty * top_price
                vendor = dlg._top_vendor or ""
                dept = dlg._top_department or ""
                items_for_db.append({
                    "project_id": page.project_id,
                    "type": "equipment",
                    "group_name": "Аренда оборудования",
                    "name": full_name,
                    "qty": top_qty,
                    "coeff": 1.0,
                    "amount": amount,
                    "unit_price": top_price,
                    "source_file": "COLUMN_MASTER",
                    "vendor": vendor,
                    "department": dept,
                    "zone": "",
                    "power_watts": 0.0,
                    "import_batch": f"columns-{datetime.utcnow().isoformat()}"
                })
                catalog_entries.append({
                    "name": full_name,
                    "unit_price": top_price,
                    "class": "equipment",
                    "vendor": vendor,
                    "power_watts": 0.0,
                    "department": dept,
                })
        except Exception as ex:
            page._log(f"Мастер колонок: ошибка обработки топов: {ex}", "error")
    # 21.13 Обработка сабов
    if dlg.chk_sub.isChecked() and dlg.ed_sub_name.text().strip():
        try:
            sub_name = dlg.ed_sub_name.text().strip()
            sub_qty = float(dlg.sp_sub_qty.value() or 0.0)
            sub_price = float(dlg.sp_sub_price.value() or 0.0)
            if sub_qty > 0.0:
                full_name = f"{prefix} {sub_name}"
                amount = sub_qty * sub_price
                vendor = dlg._sub_vendor or ""
                dept = dlg._sub_department or ""
                items_for_db.append({
                    "project_id": page.project_id,
                    "type": "equipment",
                    "group_name": "Аренда оборудования",
                    "name": full_name,
                    "qty": sub_qty,
                    "coeff": 1.0,
                    "amount": amount,
                    "unit_price": sub_price,
                    "source_file": "COLUMN_MASTER",
                    "vendor": vendor,
                    "department": dept,
                    "zone": "",
                    "power_watts": 0.0,
                    "import_batch": f"columns-{datetime.utcnow().isoformat()}"
                })
                catalog_entries.append({
                    "name": full_name,
                    "unit_price": sub_price,
                    "class": "equipment",
                    "vendor": vendor,
                    "power_watts": 0.0,
                    "department": dept,
                })
        except Exception as ex:
            page._log(f"Мастер колонок: ошибка обработки сабов: {ex}", "error")
    # 21.13.1 Добавление коробочек для сабов
    if dlg.chk_sub.isChecked() and dlg.chk_sub_boxes.isChecked():
        try:
            # По условию всегда две коробки
            for _ in range(2):
                box_name = f"{prefix} Коробка сабов"
                vendor = dlg._sub_vendor or ""
                dept = dlg._sub_department or ""
                unit_price = 0.0  # цена коробочки по умолчанию
                items_for_db.append({
                    "project_id": page.project_id,
                    "type": "equipment",
                    "group_name": "Аренда оборудования",
                    "name": box_name,
                    "qty": 1.0,
                    "coeff": 1.0,
                    "amount": unit_price,
                    "unit_price": unit_price,
                    "source_file": "COLUMN_MASTER",
                    "vendor": vendor,
                    "department": dept,
                    "zone": "",
                    "power_watts": 0.0,
                    "import_batch": f"columns-{datetime.utcnow().isoformat()}"
                })
                catalog_entries.append({
                    "name": box_name,
                    "unit_price": unit_price,
                    "class": "equipment",
                    "vendor": vendor,
                    "power_watts": 0.0,
                    "department": dept,
                })
        except Exception as ex:
            page._log(f"Мастер колонок: ошибка добавления коробочек: {ex}", "error")
    # 21.14 Пассивная система: усилители и кабели
    if dlg.rb_passive.isChecked():
        import math
        top_qty = float(dlg.sp_top_qty.value() or 0.0) if dlg.chk_top.isChecked() else 0.0
        sub_qty = float(dlg.sp_sub_qty.value() or 0.0) if dlg.chk_sub.isChecked() else 0.0
        # Усилители
        amp_name = dlg.ed_amp_name.text().strip()
        amp_qty = int(dlg.sp_amp_qty.value() or 0)
        amp_price = float(dlg.sp_amp_price.value() or 0.0)
        if amp_name and amp_qty > 0:
            # Добавляем усилители
            for i in range(amp_qty):
                amount = amp_price
                items_for_db.append({
                    "project_id": page.project_id,
                    "type": "equipment",
                    "group_name": "Аренда оборудования",
                    "name": amp_name,
                    "qty": 1.0,
                    "coeff": 1.0,
                    "amount": amount,
                    "unit_price": amp_price,
                    "source_file": "COLUMN_MASTER",
                    "vendor": dlg._amp_vendor or "",
                    "department": dlg._amp_department or "",
                    "zone": "",
                    "power_watts": 0.0,
                    "import_batch": f"columns-{datetime.utcnow().isoformat()}"
                })
            catalog_entries.append({
                "name": amp_name,
                "unit_price": amp_price,
                "class": "equipment",
                "vendor": dlg._amp_vendor or "",
                "power_watts": 0.0,
                "department": dlg._amp_department or "",
            })
        # Кабели SpeakOn
        # Используем ту же логику, что и в методе _update_passive_calc
        def calc_connectors(count: float) -> tuple[int, int]:
            n = int(count)
            if n <= 0:
                return (0, 0)
            side1 = (n + 1) // 2
            side2 = n - side1
            pairs1 = side1 // 2
            singles1 = side1 % 2
            pairs2 = side2 // 2
            singles2 = side2 % 2
            short = pairs1 + pairs2
            long = pairs1 + singles1 + pairs2 + singles2
            if long < 2:
                long = 2
            return (short, long)
        # Топы
        if top_qty > 0:
            spk_top_05, spk_top_15 = calc_connectors(top_qty)
        else:
            spk_top_05 = spk_top_15 = 0
        # Сабы
        if sub_qty > 0:
            if dlg.chk_double_sub.isChecked():
                spk_sub_05 = 0
                spk_sub_15 = max(2, int(sub_qty))
            else:
                spk_sub_05, spk_sub_15 = calc_connectors(sub_qty)
        else:
            spk_sub_05 = spk_sub_15 = 0
        # Функция для добавления кабелей из базы или вручную. Используем имена по умолчанию.
        def add_cables(name_contains: str, length_label: str, qty: int) -> None:
            """
            Добавляет позицию кабелей в смету и каталог. Пытается найти
            первую позицию в каталоге, содержащую `name_contains` и
            возвращает её название и цену. Если позиция не найдена,
            используется `name_contains` вместе с `length_label` как
            название и нулевая цена. Количество `qty` суммируется в одну
            строку. Длина кабеля не влияет на поиск, но добавляет
            пояснение в название, если позиция не найдена.
            """
            if qty <= 0:
                return
            try:
                # Ищем в каталоге позицию по имени, без учёта регистра
                rows = page.db.catalog_list({"name": name_contains})
            except Exception:
                rows = []
            if rows:
                row0 = rows[0]
                unit_price = float(row0["unit_price"] or 0.0)
                vendor = normalize_case(row0["vendor"] or "")
                dept = normalize_case(row0["department"] or "")
                item_name = normalize_case(row0["name"] or name_contains)
            else:
                unit_price = 0.0
                vendor = ""
                dept = ""
                # Добавляем длину в название для различения коротких и длинных
                item_name = normalize_case(f"{name_contains} {length_label}")
            amount = unit_price * qty
            items_for_db.append({
                "project_id": page.project_id,
                "type": "equipment",
                "group_name": "Аренда оборудования",
                "name": item_name,
                "qty": float(qty),
                "coeff": 1.0,
                "amount": amount,
                "unit_price": unit_price,
                "source_file": "COLUMN_MASTER",
                "vendor": vendor,
                "department": dept,
                "zone": "",
                "power_watts": 0.0,
                "import_batch": f"columns-{datetime.utcnow().isoformat()}"
            })
            catalog_entries.append({
                "name": item_name,
                "unit_price": unit_price,
                "class": "equipment",
                "vendor": vendor,
                "power_watts": 0.0,
                "department": dept,
            })
        # Суммируем количество кабелей для топов и сабов
        total_short = spk_top_05 + spk_sub_05
        total_long = spk_top_15 + spk_sub_15
        # Если выбраны коробочки для сабов, добавляем по одному короткому кабелю на каждую коробку
        if dlg.chk_sub.isChecked() and dlg.chk_sub_boxes.isChecked():
            # Всегда две коробки по заданию
            total_short += 2
        # Добавляем короткие и длинные SpeakOn (одна запись на каждую длину)
        add_cables("speakon", "NL4 0,5м", total_short)
        add_cables("speakon", "NL4 15м", total_long)
    # 21.15 Проверка наличия позиций
    if not items_for_db:
        QtWidgets.QMessageBox.information(page, "Внимание", "Ничего не выбрано для добавления.")
        return
    # 21.16 Запись позиций в базу
    try:
        page.db.add_items_bulk(items_for_db)
        if hasattr(page.db, "catalog_add_or_ignore"):
            page.db.catalog_add_or_ignore(catalog_entries)
        page._log(f"Мастер колонок: добавлено позиций {len(items_for_db)}.")
        QtWidgets.QMessageBox.information(page, "Готово", f"Добавлено позиций: {len(items_for_db)}.")
    except Exception as ex:
        page._log(f"Мастер колонок: ошибка добавления: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось добавить колонки: {ex}")
        return
    # 21.17 Перезагружаем таблицы сметы
    try:
        page._reload_zone_tabs()
    except Exception:
        pass


def edit_selected_screen(page: Any) -> None:
    """Редактирует выбранный экран и связанные с ним позиции (витая пара, видеопроцессор).

    Пользователь должен предварительно выделить строку в таблице сводной сметы,
    содержащую экран (наименование начинается с «LED экран»). Функция извлекает
    исходные параметры экрана (ширина, высота, разрешение модуля, цена за м²,
    подрядчик, отдел) из имени и таблицы, затем открывает диалог мастера с
    этими значениями. После изменения параметры записываются обратно в
    соответствующие строки таблицы «items» проекта, а также корректируются
    связанные записи витой пары и видеопроцессора. При отключении
    соответствующих чекбоксов в диалоге такие позиции будут удалены.
    """
    # 21.1 Проверяем выбор проекта
    if getattr(page, "project_id", None) is None:
        QtWidgets.QMessageBox.information(page, "Внимание", "Сначала откройте проект.")
        return
    # 21.2 Находим выбранную строку в таблицах зон
    selected_info = None
    for zone_val, tbl in page.zone_tables.items():
        sel = tbl.selectionModel() if tbl else None
        if sel and sel.hasSelection():
            rows = sel.selectedRows()
            if rows:
                row_idx = rows[0].row()
                # Собираем данные: название, цена, vendor, dept, zone
                name = tbl.item(row_idx, 0).text() if tbl.item(row_idx, 0) else ""
                price_txt = tbl.item(row_idx, 3).text() if tbl.item(row_idx, 3) else ""
                vendor = tbl.item(row_idx, 5).text() if tbl.item(row_idx, 5) else ""
                department = tbl.item(row_idx, 6).text() if tbl.item(row_idx, 6) else ""
                # Определяем зону ("Без зоны" -> "")
                zone_name = tbl.item(row_idx, 7).text() if tbl.item(row_idx, 7) else ""
                zone = "" if not zone_name or zone_name.lower() in {"без зоны", "<пусто>"} else zone_name.strip()
                selected_info = (name, price_txt, vendor, department, zone)
                break
    if not selected_info:
        QtWidgets.QMessageBox.information(page, "Редактирование", "Выберите экран для редактирования.")
        return
    name, price_txt, vendor_old, dept_old, zone_old = selected_info
    # 21.3 Проверяем, является ли выбранный элемент экраном
    if not name.lower().startswith("led экран"):
        QtWidgets.QMessageBox.information(page, "Редактирование", "Выбранная позиция не является экраном.")
        return
    import re
    import math
    # 21.4 Извлекаем размеры и разрешение из имени
    pattern = r"LED экран ([0-9]+(?:\.[0-9]+)?)×([0-9]+(?:\.[0-9]+)?) м \((\d+) кабинетов, (\d+)×(\d+) пикселей\)"
    m = re.match(pattern, name)
    if not m:
        QtWidgets.QMessageBox.information(page, "Редактирование", "Не удалось распарсить параметры экрана.")
        return
    try:
        width = float(m.group(1))
        height = float(m.group(2))
        cabinets_old = int(m.group(3))
        res_x_old = int(m.group(4))
        res_y_old = int(m.group(5))
    except Exception:
        QtWidgets.QMessageBox.information(page, "Редактирование", "Ошибка разбора параметров экрана.")
        return
    # 21.5 Вычисляем разрешение модуля из старых параметров
    cab_w_old = max(1, math.ceil(width * 2))
    cab_h_old = max(1, math.ceil(height * 2))
    mod_px_w = int(res_x_old / cab_w_old)
    mod_px_h = int(res_y_old / cab_h_old)
    # 21.6 Цена за м² (unit_price) - преобразуем
    try:
        price_per_m2 = to_float(price_txt, 0.0)
    except Exception:
        price_per_m2 = 0.0
    # 21.7 Определяем наличие витой пары и видеопроцессора
    has_cable = False
    has_vp = False
    try:
        cur = page.db._conn.cursor()
        cur.execute(
            "SELECT COUNT(*) FROM items WHERE project_id=? AND name LIKE ? AND LOWER(COALESCE(vendor,''))=LOWER(?) AND LOWER(COALESCE(zone,''))=LOWER(?)",
            (page.project_id, "Витая пара для LED%", vendor_old, zone_old or ""),
        )
        has_cable = cur.fetchone()[0] > 0
        cur.execute(
            "SELECT COUNT(*) FROM items WHERE project_id=? AND name LIKE ? AND LOWER(COALESCE(vendor,''))=LOWER(?) AND LOWER(COALESCE(zone,''))=LOWER(?)",
            (page.project_id, "Видеопроцессор для LED%", vendor_old, zone_old or ""),
        )
        has_vp = cur.fetchone()[0] > 0
    except Exception:
        has_cable = False
        has_vp = False
    # 21.8 Запускаем диалог мастера с предзаполненными значениями
    dlg = open_screen_master.__globals__["ScreenMasterDialog"](page)  # type: ignore
    dlg.ed_width.setValue(width)
    dlg.ed_height.setValue(height)
    dlg.spin_mod_w.setValue(mod_px_w)
    dlg.spin_mod_h.setValue(mod_px_h)
    dlg.spin_price.setValue(price_per_m2)
    # Устанавливаем подрядчика и отдел
    def set_combo(combo: QtWidgets.QComboBox, text: str) -> None:
        if not text:
            return
        for i in range(combo.count()):
            if combo.itemText(i).lower() == text.lower():
                combo.setCurrentIndex(i)
                return
        combo.addItem(text)
        combo.setCurrentIndex(combo.count() - 1)
    set_combo(dlg.cmb_vendor, vendor_old)
    set_combo(dlg.cmb_department, dept_old)
    dlg.chk_cable.setChecked(has_cable)
    dlg.chk_vp.setChecked(has_vp)
    # Обновляем расчёты
    try:
        dlg._update_labels()
    except Exception:
        pass
    # Показываем диалог
    if dlg.exec() != QtWidgets.QDialog.Accepted:
        return
    # 21.9 Собираем новые параметры
    width_new = float(dlg.ed_width.value())
    height_new = float(dlg.ed_height.value())
    mod_w_new = int(dlg.spin_mod_w.value())
    mod_h_new = int(dlg.spin_mod_h.value())
    price_new = float(dlg.spin_price.value())
    vendor_new = dlg.cmb_vendor.currentText().strip()
    dept_new = dlg.cmb_department.currentText().strip()
    # Пересчитываем derived values
    cab_w_new = max(1, math.ceil(width_new * 2))
    cab_h_new = max(1, math.ceil(height_new * 2))
    cabinets_new = cab_w_new * cab_h_new
    res_x_new = cab_w_new * mod_w_new
    res_y_new = cab_h_new * mod_h_new
    area_new = width_new * height_new
    total_px = res_x_new * res_y_new
    cables_new = max(1, math.ceil(total_px / 650_000))
    # Формируем новые имена
    new_screen_name = (
        f"LED экран {fmt_num(width_new, 2)}×{fmt_num(height_new, 2)} м "
        f"({cabinets_new} кабинетов, {res_x_new}×{res_y_new} пикселей)"
    )
    new_cable_name = f"Витая пара для LED {fmt_num(width_new, 2)}×{fmt_num(height_new, 2)} м"
    new_vp_name = f"Видеопроцессор для LED {fmt_num(width_new, 2)}×{fmt_num(height_new, 2)} м"
    import_batch = None
    # 21.10 Обновляем записи в таблице items
    try:
        cur = page.db._conn.cursor()
        # Находим экранную позицию
        cur.execute(
            "SELECT id, import_batch FROM items WHERE project_id=? AND name=? COLLATE NOCASE AND LOWER(COALESCE(vendor,''))=LOWER(?) AND LOWER(COALESCE(zone,''))=LOWER(?) LIMIT 1",
            (page.project_id, name, vendor_old, zone_old or ""),
        )
        row = cur.fetchone()
        if row:
            item_id = row["id"] if isinstance(row, dict) else row[0]
            import_batch = row["import_batch"] if isinstance(row, dict) else row[1]
            # Обновляем экран
            page.db.update_item_fields(item_id, {
                "name": new_screen_name,
                "qty": area_new,
                "unit_price": price_new,
                "amount": area_new * price_new,
                "vendor": vendor_new,
                "department": dept_new,
            })
        # Обновляем/удаляем витую пару
        cur.execute(
            "SELECT id, unit_price FROM items WHERE project_id=? AND name LIKE ? AND LOWER(COALESCE(vendor,''))=LOWER(?) AND LOWER(COALESCE(zone,''))=LOWER(?)",
            (page.project_id, "Витая пара для LED%", vendor_old, zone_old or ""),
        )
        cable_rows = cur.fetchall()
        if dlg.chk_cable.isChecked():
            # Надо либо обновить существующие, либо добавить новую, если нет
            if cable_rows:
                for r in cable_rows:
                    cid = r["id"] if isinstance(r, dict) else r[0]
                    c_price = float(r["unit_price"] if isinstance(r, dict) else r[1] or 0.0)
                    page.db.update_item_fields(cid, {
                        "name": new_cable_name,
                        "qty": float(cables_new),
                        "amount": c_price * cables_new,
                        "vendor": vendor_new,
                        "department": dept_new,
                    })
            else:
                # Добавляем новую запись
                page.db.add_items_bulk([
                    {
                        "project_id": page.project_id,
                        "type": "equipment",
                        "group_name": "Аренда оборудования",
                        "name": new_cable_name,
                        "qty": float(cables_new),
                        "coeff": 1.0,
                        "amount": 0.0,
                        "unit_price": 0.0,
                        "source_file": "SCREEN_MASTER",
                        "vendor": vendor_new,
                        "department": dept_new,
                        "zone": zone_old or "",
                        "power_watts": 0.0,
                        "import_batch": import_batch or f"screen-edit-{datetime.utcnow().isoformat()}"
                    }
                ])
        else:
            # Пользователь снял галочку: удаляем существующие записи
            if cable_rows:
                ids = [ (r["id"] if isinstance(r, dict) else r[0]) for r in cable_rows ]
                page.db.delete_items(ids)
        # Обновляем/удаляем видеопроцессор
        cur.execute(
            "SELECT id, unit_price FROM items WHERE project_id=? AND name LIKE ? AND LOWER(COALESCE(vendor,''))=LOWER(?) AND LOWER(COALESCE(zone,''))=LOWER(?)",
            (page.project_id, "Видеопроцессор для LED%", vendor_old, zone_old or ""),
        )
        vp_rows = cur.fetchall()
        if dlg.chk_vp.isChecked():
            if vp_rows:
                for r in vp_rows:
                    vid = r["id"] if isinstance(r, dict) else r[0]
                    v_price = float(r["unit_price"] if isinstance(r, dict) else r[1] or 0.0)
                    page.db.update_item_fields(vid, {
                        "name": new_vp_name,
                        "vendor": vendor_new,
                        "department": dept_new,
                    })
            else:
                # Добавляем видеопроцессор
                page.db.add_items_bulk([
                    {
                        "project_id": page.project_id,
                        "type": "equipment",
                        "group_name": "Аренда оборудования",
                        "name": new_vp_name,
                        "qty": 1.0,
                        "coeff": 1.0,
                        "amount": 0.0,
                        "unit_price": 0.0,
                        "source_file": "SCREEN_MASTER",
                        "vendor": vendor_new,
                        "department": dept_new,
                        "zone": zone_old or "",
                        "power_watts": 0.0,
                        "import_batch": import_batch or f"screen-edit-{datetime.utcnow().isoformat()}"
                    }
                ])
        else:
            # Пользователь снял галочку: удаляем видеопроцессор
            if vp_rows:
                ids = [ (r["id"] if isinstance(r, dict) else r[0]) for r in vp_rows ]
                page.db.delete_items(ids)
        # Добавляем/обновляем записи в каталоге
        try:
            catalog_entries: List[Dict[str, Any]] = []
            catalog_entries.append({
                "name": new_screen_name,
                "unit_price": price_new,
                "class": "equipment",
                "vendor": vendor_new,
                "power_watts": 0.0,
                "department": dept_new,
            })
            if dlg.chk_cable.isChecked():
                catalog_entries.append({
                    "name": new_cable_name,
                    "unit_price": 0.0,
                    "class": "equipment",
                    "vendor": vendor_new,
                    "power_watts": 0.0,
                    "department": dept_new,
                })
            if dlg.chk_vp.isChecked():
                catalog_entries.append({
                    "name": new_vp_name,
                    "unit_price": 0.0,
                    "class": "equipment",
                    "vendor": vendor_new,
                    "power_watts": 0.0,
                    "department": dept_new,
                })
            if hasattr(page.db, "catalog_add_or_ignore"):
                page.db.catalog_add_or_ignore(catalog_entries)
        except Exception:
            pass
        # Логирование
        try:
            page._log(
                f"Редактирование экрана: обновлено. Размеры {width_new}×{height_new} м, рез. модуля {mod_w_new}×{mod_h_new}, кабелей {cables_new}.",
            )
        except Exception:
            pass
    except Exception as ex:
        # Ошибка обновления
        page._log(f"Ошибка редактирования экрана: {ex}", "error")
        QtWidgets.QMessageBox.critical(page, "Ошибка", f"Не удалось обновить экран: {ex}")
        return
    # 21.11 Обновляем интерфейс
    try:
        page._reload_zone_tabs()
    except Exception:
        pass
