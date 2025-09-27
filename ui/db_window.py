"""
Назначение:
    Окно «База данных (каталог)» — просмотр/поиск/редактирование глобального каталога.

Как работает:
    - Фильтры по наименованию/классу/подрядчику/отделу.
    - Импорт/экспорт, удаление дублей, выделение отклонений цен.
    - Редактирование класса/мощностей, commit и синхронизация с открытым проектом.

Стиль:
    - Пронумерованные секции и краткие комментарии.
"""

# 1. Импорт
from PySide6 import QtWidgets, QtGui, QtCore              # Qt
from typing import Optional, Callable                     # типы
from pathlib import Path
import csv                                                # импорт/экспорт CSV
from .common import (                                     # общие константы/утилиты
    CLASS_RU2EN, CLASS_EN2RU, WRAP_THRESHOLD,
    fmt_num, fmt_sign, to_float,
    apply_auto_col_resize, setup_auto_col_resize, setup_priority_name,
    # Импортируем функции для канонического поиска
    make_search_key, contains_search
)
from .delegates import ClassRuDelegate, WrapTextDelegate   # делегаты
from db import DB                                         # база данных

# 2. Класс DatabaseWindow
class DatabaseWindow(QtWidgets.QDialog):
    def __init__(self, db: DB, parent=None, log_fn=None,
                 project_id_for_sync: Optional[int] = None,
                 reload_summary_cb: Optional[Callable] = None):
        super().__init__(parent)
        self.db = db
        self.log_fn = log_fn
        self.project_id_for_sync = project_id_for_sync
        self.reload_summary_cb = reload_summary_cb
        self.setWindowTitle("База данных (каталог)")
        self.resize(1200, 780)

        # 7.1 Верхняя панель фильтров
        top = QtWidgets.QHBoxLayout()
        top.setContentsMargins(4, 4, 4, 4); top.setSpacing(8)
        self.edit_name = QtWidgets.QLineEdit(); self.edit_name.setPlaceholderText("Фильтр по наименованию...")
        self.combo_class = QtWidgets.QComboBox()
        self.combo_vendor = QtWidgets.QComboBox()
        self.combo_department = QtWidgets.QComboBox()
        self.check_deviation = QtWidgets.QCheckBox("Отклонение от средней цены")
        self.btn_import = QtWidgets.QPushButton("Импорт CSV")
        self.btn_export = QtWidgets.QPushButton("Экспорт CSV")
        for w in (self.edit_name, self.combo_class, self.combo_vendor, self.combo_department,
                  self.check_deviation, self.btn_import, self.btn_export):
            top.addWidget(w)
        top.addStretch(1)

        # 7.2 Панель действий
        actions = QtWidgets.QHBoxLayout()
        actions.setContentsMargins(4, 0, 4, 4); actions.setSpacing(8)
        self.btn_commit = QtWidgets.QPushButton("Сохранить изменения")
        self.btn_delete = QtWidgets.QPushButton("Удалить выбранные")
        self.combo_mass_class = QtWidgets.QComboBox(); self.combo_mass_class.addItems(list(CLASS_RU2EN.keys()))
        self.btn_mass_set = QtWidgets.QPushButton("Присвоить класс выбранным")
        self.btn_check_dups = QtWidgets.QPushButton("Проверить дубли")
        self.btn_remove_dups = QtWidgets.QPushButton("Удалить дубли")
        for w in (self.btn_commit, self.btn_delete, self.combo_mass_class, self.btn_mass_set, self.btn_check_dups, self.btn_remove_dups):
            actions.addWidget(w)
        actions.addStretch(1)

        # 7.3 Таблица каталога
        # Расширяем таблицу: добавляем столбец "Склад (шт)", который
        # отображает количество оборудования на складе у подрядчика. Теперь
        # всего 10 столбцов: ID, Наименование, Класс, Подрядчик, Цена,
        # Потребление, Отдел, Склад, Добавлено, Отклонение.
        self.table = QtWidgets.QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels([
            "ID",
            "Наименование",
            "Класс",
            "Подрядчик",
            "Цена",
            "Потребление (Вт)",
            "Отдел",
            "Склад (шт)",
            "Добавлено",
            "Отклонение",
        ])
        setup_priority_name(self.table, name_col=1)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.SelectedClicked)
        self.table.setWordWrap(False)
        self.table.setItemDelegateForColumn(1, WrapTextDelegate(self.table, wrap_threshold=WRAP_THRESHOLD))
        self.table.setItemDelegateForColumn(3, WrapTextDelegate(self.table, wrap_threshold=WRAP_THRESHOLD))
        self.table.setItemDelegateForColumn(6, WrapTextDelegate(self.table, wrap_threshold=WRAP_THRESHOLD))
        self.table.setItemDelegateForColumn(2, ClassRuDelegate(self.table))
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        # 7.4 Компоновка
        lay = QtWidgets.QVBoxLayout(self); lay.addLayout(top); lay.addLayout(actions); lay.addWidget(self.table, 1)

        # 7.5 Сигналы
        self.edit_name.textChanged.connect(self.reload)
        self.combo_class.currentTextChanged.connect(self.reload)
        self.combo_vendor.currentTextChanged.connect(self.reload)
        self.combo_department.currentTextChanged.connect(self.reload)
        self.check_deviation.toggled.connect(self.reload)
        self.btn_import.clicked.connect(self.on_import_csv)
        self.btn_export.clicked.connect(self.on_export_csv)
        self.btn_commit.clicked.connect(self.on_commit)
        self.btn_delete.clicked.connect(self.on_delete_selected)
        self.btn_mass_set.clicked.connect(self.on_mass_set_class)
        self.btn_check_dups.clicked.connect(self.on_check_dups)
        self.btn_remove_dups.clicked.connect(self.on_remove_dups)
        self.table.itemChanged.connect(self.on_item_changed)

        # 7.6 Первичная загрузка
        self.reload_filters(); self.reload()

    # 7.7 Лог-вспомогалка
    def _log(self, msg: str, level: str = "info"):
        if callable(self.log_fn): self.log_fn(msg, level)

    # 7.8 Загрузка списков фильтров
    def reload_filters(self):
        self.combo_class.blockSignals(True); self.combo_vendor.blockSignals(True); self.combo_department.blockSignals(True)
        self.combo_class.clear(); self.combo_class.addItem("<ВСЕ>")
        for ru in CLASS_RU2EN.keys(): self.combo_class.addItem(ru)
        self.combo_vendor.clear(); self.combo_vendor.addItem("<ALL>")
        for v in self.db.catalog_distinct_values("vendor"): self.combo_vendor.addItem(v)
        self.combo_department.clear(); self.combo_department.addItem("<ALL>")
        for d in self.db.catalog_distinct_values("department"): self.combo_department.addItem(d)
        self.combo_class.blockSignals(False); self.combo_vendor.blockSignals(False); self.combo_department.blockSignals(False)

    # 7.9 Перечитать таблицу каталога
    def reload(self):
        class_ru = self.combo_class.currentText()
        class_en = CLASS_RU2EN.get(class_ru, None)
        # Формируем фильтры без учёта поля name: поиск по названию выполняем в Python,
        # чтобы игнорировать регистр и кириллица/латиница-хомоглифы.
        filters = {
            "name": "",  # пустая строка для исключения поиска в SQL
            "class": class_en if class_ru not in ("", "<ВСЕ>") else "<ALL>",
            "vendor": self.combo_vendor.currentText(),
            "department": self.combo_department.currentText(),
        }
        rows = self.db.catalog_list(filters)
        # Фильтруем по введённому наименованию
        search_raw = self.edit_name.text() or ""
        if search_raw:
            filtered = []
            for r in rows:
                try:
                    nm = r["name"] if r["name"] is not None else ""
                    ven = r["vendor"] if r["vendor"] is not None else ""
                    dep = r["department"] if r["department"] is not None else ""
                    cls = CLASS_EN2RU.get((r["class"] or "equipment"), "Оборудование")
                    if (
                        contains_search(nm, search_raw)
                        or contains_search(ven, search_raw)
                        or contains_search(dep, search_raw)
                        or contains_search(cls, search_raw)
                    ):
                        filtered.append(r)
                except Exception:
                    continue
            rows = filtered
            try:
                self._log(f"Каталог: поиск '{search_raw}' отфильтровал {len(rows)} строк.")
            except Exception:
                pass
        self.table.blockSignals(True); self.table.setRowCount(0)
        show_dev = self.check_deviation.isChecked()

        for r in rows:
            row = self.table.rowCount(); self.table.insertRow(row)
            class_ru_show = CLASS_EN2RU.get((r["class"] or "equipment"), "Оборудование")
            # Вставляем значения столбцов. Учтите, что stock_qty может быть None.
            stock_val = 0.0
            try:
                stock_val = float(r["stock_qty"] or 0)
            except Exception:
                stock_val = 0.0
            cells = [
                str(r["id"]),
                r["name"],
                class_ru_show,
                r["vendor"] or "",
                fmt_num(float(r["unit_price"]), 2),
                fmt_num(float(r["power_watts"] or 0), 0),
                r["department"] or "",
                fmt_num(stock_val, 2),
                r["created_at"],
                "",
            ]
            for col, val in enumerate(cells):
                item = QtWidgets.QTableWidgetItem(val)
                if col == 0:
                    item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table.setItem(row, col, item)

            if show_dev:
                avg = self.db.catalog_avg_price_by_name(r["name"])
                diff = float(r["unit_price"]) - float(avg or 0)
                dev_item = self.table.item(row, 9)
                dev_item.setText(fmt_sign(diff, 2))
                if diff > 0:
                    dev_item.setForeground(QtGui.QBrush(QtGui.QColor(220, 80, 80)))
                elif diff < 0:
                    dev_item.setForeground(QtGui.QBrush(QtGui.QColor(70, 200, 120)))
            else:
                self.table.item(row, 9).setText("")

        # Столбец девиации находится в последней колонке (index=9)
        self.table.setColumnHidden(9, not show_dev)
        self.table.blockSignals(False)
        apply_auto_col_resize(self.table)
        self._log(f"Каталог: обновлена таблица ({len(rows)} строк), автоширина применена.")

    # 7.10 Импорт/экспорт/commit
    def on_import_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите CSV", "", "CSV (*.csv)")
        if not path: return
        try:
            added = self.db.catalog_import_csv(Path(path))
            QtWidgets.QMessageBox.information(self, "Готово", f"Импортировано строк (включая игнор дублей): {added}")
            self._log(f"Импорт CSV: {path} (+{added})")
            self.reload_filters(); self.reload()
        except Exception as ex:
            self._log(f"Ошибка импорта CSV: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось импортировать: {ex}")

    def on_export_csv(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить CSV", "catalog_export.csv", "CSV (*.csv)")
        if not path: return
        try:
            class_ru = self.combo_class.currentText()
            class_en = CLASS_RU2EN.get(class_ru, None)
            filters = {
                "name": self.edit_name.text(),
                "class": class_en if class_ru not in ("", "<ВСЕ>") else "<ALL>",
                "vendor": self.combo_vendor.currentText(),
                "department": self.combo_department.currentText(),
            }
            n = self.db.catalog_export_csv(Path(path), filters)
            self._log(f"Экспорт CSV: {path} ({n} строк)")
            QtWidgets.QMessageBox.information(self, "Готово", f"Экспортировано строк: {n}")
        except Exception as ex:
            self._log(f"Ошибка экспорта CSV: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать: {ex}")

    def on_commit(self):
        try:
            self.db.commit()
            self._log("Каталог: сохранение изменений (commit)")
            if self.project_id_for_sync is not None:
                try:
                    upd_class, upd_power = self.db.project_sync_from_catalog(self.project_id_for_sync)
                    self._log(f"Синхронизация сметы: обновлено классов={upd_class}, мощностей={upd_power}.")
                    if callable(self.reload_summary_cb):
                        self.reload_summary_cb()
                except Exception as sync_ex:
                    self._log(f"Ошибка синхронизации сметы: {sync_ex}", "error")
            QtWidgets.QMessageBox.information(self, "Сохранено", "Изменения сохранены.")
        except Exception as ex:
            self._log(f"Ошибка commit: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить: {ex}")

    # 7.11 Удаление/массовая смена/дубли
    def on_delete_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows: return
        ids = [int(self.table.item(r.row(), 0).text()) for r in rows]
        if QtWidgets.QMessageBox.question(self, "Подтверждение", f"Удалить {len(ids)} записей из базы?") != QtWidgets.QMessageBox.Yes:
            return
        try:
            n = self.db.catalog_delete_ids(ids)
            self._log(f"Каталог: удалено записей {n}")
            QtWidgets.QMessageBox.information(self, "Готово", f"Удалено: {n}")
            self.reload()
        except Exception as ex:
            self._log(f"Ошибка удаления: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось удалить: {ex}")

    def on_mass_set_class(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows: return
        ids = [int(self.table.item(r.row(), 0).text()) for r in rows]
        ru = self.combo_mass_class.currentText()
        en = CLASS_RU2EN.get(ru, "equipment")
        try:
            n = self.db.catalog_bulk_update_class(ids, en)
            self._log(f"Каталог: массовая смена класса -> {ru} ({n} шт.)")
            QtWidgets.QMessageBox.information(self, "Готово", f"Обновлено: {n}")
            self.reload()
        except Exception as ex:
            self._log(f"Ошибка массовой смены класса: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось сменить класс массово: {ex}")

    def on_check_dups(self):
        dups = self.db.catalog_find_duplicates()
        if not dups:
            QtWidgets.QMessageBox.information(self, "Результат", "Дубликаты не найдены.")
            self._log("Каталог: дубликаты не найдены")
            return
        dup_ids = {i for ids in dups.values() for i in ids}
        for r in range(self.table.rowCount()):
            rid = int(self.table.item(r, 0).text())
            if rid in dup_ids:
                for c in range(self.table.columnCount()):
                    self.table.item(r, c).setBackground(QtGui.QColor(255, 220, 220))
        self._log(f"Каталог: найдено групп дублей {len(dups)}")
        QtWidgets.QMessageBox.information(self, "Результат", f"Найдено групп дублей: {len(dups)}. Подсветил красным.")

    def on_remove_dups(self):
        try:
            deleted = self.db.catalog_delete_duplicates()
            if deleted == 0:
                self._log("Каталог: дублей для удаления нет")
                QtWidgets.QMessageBox.information(self, "Результат", "Дубликаты отсутствуют.")
            else:
                self._log(f"Каталог: удалено дублей {deleted}")
                QtWidgets.QMessageBox.information(self, "Готово", f"Удалено дублей: {deleted}")
                self.reload()
        except Exception as ex:
            self._log(f"Ошибка удаления дублей: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось удалить: {ex}")

    # 7.12 Редактирование полей (класс/мощность)
    def on_item_changed(self, item: QtWidgets.QTableWidgetItem):
        row = item.row(); col = item.column()
        rid = int(self.table.item(row, 0).text())
        if col == 2:
            ru = self.table.item(row, 2).text().strip() or "Оборудование"
            en = CLASS_RU2EN.get(ru, "equipment")
            try:
                self.db.catalog_update_field(rid, "class", en)
                self._log(f"Каталог: записан класс '{ru}' для id={rid}")
            except Exception as ex:
                self._log(f"Ошибка записи класса: {ex}", "error")
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось записать класс: {ex}")
        elif col == 5:
            txt = (self.table.item(row, 5).text() or "").strip()
            try:
                val = to_float(txt, 0.0)
                if val < 0: val = 0
                self.db.catalog_update_field(rid, "power_watts", val)
                self._log(f"Каталог: записана мощность {fmt_num(val,0)} Вт для id={rid}")
            except Exception as ex:
                self._log(f"Ошибка записи мощности: {ex}", "error")
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось записать мощность: {ex}")

# 8. Страница проекта (вкладки)
