"""
Назначение:
    Диалоги: перенос между зонами и подтверждение изменения мощностей.

Как работает:
    - MoveDialog — задаёт объёмы для переноса выбранных позиций в другую зону.
    - PowerMismatchDialog — показывает расхождения мощностей и предлагает обновить каталог.

Стиль:
    - Нумерованные секции и краткие комментарии.
"""

# 1. Импорт
from PySide6 import QtWidgets, QtGui, QtCore
from typing import List, Tuple, Dict
from .common import fmt_num
from .widgets import SmartDoubleSpinBox

# 2. Диалог переноса между зонами
class MoveDialog(QtWidgets.QDialog):
    def __init__(self, items_data: List[Tuple[int, str, float, int]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Перенос позиций в зону")
        self.resize(780, 420)
        self.result_moves: Dict[int, float] = {}
        v = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "Наименование", "Доступно", "Перенести"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        for item_id, name, avail, _row in items_data:
            r = self.table.rowCount(); self.table.insertRow(r)
            id_item = QtWidgets.QTableWidgetItem(str(item_id)); id_item.setFlags(id_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.table.setItem(r, 0, id_item)
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(name))
            avail_item = QtWidgets.QTableWidgetItem(fmt_num(avail,3)); avail_item.setFlags(avail_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.table.setItem(r, 2, avail_item)
            spin = SmartDoubleSpinBox(); spin.setDecimals(3); spin.setMinimum(0.000); spin.setMaximum(avail); spin.setValue(min(avail, 1.000))
            self.table.setCellWidget(r, 3, spin)
        v.addWidget(self.table, 1)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._accept); btns.rejected.connect(self.reject); v.addWidget(btns)
    def _accept(self):
        moves: Dict[int, float] = {}
        for r in range(self.table.rowCount()):
            item_id = int(self.table.item(r, 0).text())
            spin: SmartDoubleSpinBox = self.table.cellWidget(r, 3)
            mv = float(spin.value() or 0)
            if mv > 0: moves[item_id] = mv
        self.result_moves = moves; self.accept()

# 10. Диалог расхождений по мощностям

# 3. Диалог расхождений по мощностям
class PowerMismatchDialog(QtWidgets.QDialog):
    def __init__(self, diffs: List[Tuple[str, str, float, List[float]]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Обновление мощностей в каталоге")
        self.resize(900, 420)
        v = QtWidgets.QVBoxLayout(self)
        label = QtWidgets.QLabel("Обнаружены новые значения «Потребление (Вт)» для указанных позиций.\n"
                                 "Применить новые значения ко всем записям каталога с тем же (Наименование, Подрядчик)?")
        label.setWordWrap(True); v.addWidget(label)
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Наименование", "Подрядчик", "Старое(ые) значения, Вт", "Новое, Вт"])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        for name, vendor, new_pw, olds in diffs:
            r = self.table.rowCount(); self.table.insertRow(r)
            self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(name))
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(vendor))
            self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(", ".join(fmt_num(o,0) for o in olds)))
            self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(fmt_num(new_pw,0)))
        v.addWidget(self.table, 1)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.button(QtWidgets.QDialogButtonBox.Ok).setText("Применить новые значения")
        btns.button(QtWidgets.QDialogButtonBox.Cancel).setText("Оставить старые")
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)
        v.addWidget(btns)

# 11. Виджет Drag&Drop для файлов
