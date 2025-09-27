"""
Назначение:
    Делегаты отображения/редактирования для таблиц (перенос текста и выбор класса).

Как работает:
    - ClassRuDelegate — выпадающий список с русскими названиями классов.
    - WrapTextDelegate — умный перенос длинного текста в ячейках (с порогом WRAP_THRESHOLD).

Стиль:
    - Нумерованные секции и краткие комментарии.
"""

# 1. Импорт
from PySide6 import QtWidgets, QtGui, QtCore
from .common import CLASS_RU2EN, WRAP_THRESHOLD

# 2. Делегат класса (RU)
class ClassRuDelegate(QtWidgets.QStyledItemDelegate):
    RU_LIST = list(CLASS_RU2EN.keys())
    def createEditor(self, parent, option, index):
        combo = QtWidgets.QComboBox(parent); combo.addItems(self.RU_LIST); return combo
    def setEditorData(self, editor, index):
        txt = index.data() or "Оборудование"; i = max(0, editor.findText(txt)); editor.setCurrentIndex(i)
    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), QtCore.Qt.EditRole)

# 3. Делегат переноса текста
class WrapTextDelegate(QtWidgets.QStyledItemDelegate):
    """Перенос текста включается только при длине > WRAP_THRESHOLD."""
    def __init__(self, table: QtWidgets.QTableWidget, min_height: int = 28, wrap_threshold: int = WRAP_THRESHOLD, parent=None):
        super().__init__(parent)
        self._table = table
        self._min_h = min_height
        self._thr = max(0, int(wrap_threshold))
    def paint(self, painter: QtGui.QPainter, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex):
        opt = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        style = opt.widget.style() if opt.widget else QtWidgets.QApplication.style()
        text = str(index.data() or "")
        opt.text = ""
        style.drawControl(QtWidgets.QStyle.CE_ItemViewItem, opt, painter, opt.widget)
        if not text: return
        rect = opt.rect.adjusted(4, 2, -4, -2)
        if len(text) <= self._thr:
            painter.save()
            if opt.state & QtWidgets.QStyle.State_Selected:
                painter.setPen(opt.palette.color(QtGui.QPalette.Active, QtGui.QPalette.HighlightedText))
            else:
                painter.setPen(opt.palette.color(QtGui.QPalette.Active, QtGui.QPalette.Text))
            fm = opt.fontMetrics
            elided = fm.elidedText(text, QtCore.Qt.ElideRight, rect.width())
            painter.drawText(rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, elided)
            painter.restore()
        else:
            doc = QtGui.QTextDocument()
            doc.setDefaultFont(opt.font)
            doc.setTextWidth(rect.width())
            doc.setPlainText(text)
            painter.save()
            painter.translate(rect.topLeft())
            ctx = QtGui.QAbstractTextDocumentLayout.PaintContext()
            if opt.state & QtWidgets.QStyle.State_Selected:
                ctx.palette.setColor(QtGui.QPalette.Text, opt.palette.color(QtGui.QPalette.Active, QtGui.QPalette.HighlightedText))
            doc.documentLayout().draw(painter, ctx)
            painter.restore()
    def sizeHint(self, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex) -> QtCore.QSize:
        text = str(index.data() or "")
        width = max(30, self._table.columnWidth(index.column()) - 8)
        if len(text) <= self._thr:
            h = option.fontMetrics.height() + 6
            return QtCore.QSize(width, max(self._min_h, h))
        else:
            doc = QtGui.QTextDocument()
            doc.setDefaultFont(option.font)
            doc.setTextWidth(width)
            doc.setPlainText(text)
            h = int(doc.size().height())
            return QtCore.QSize(width, max(self._min_h, h + 6))

# 6. Главное окно
