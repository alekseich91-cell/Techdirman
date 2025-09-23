"""
Назначение:
    Набор виджетов интерфейса: умный спинбокс, виджет drop-картинки, док «Лог», виджет drop-файла.

Как работает:
    - SmartDoubleSpinBox — отображает значения без хвостовых нулей и принимает запятую.
    - ImageDropLabel — drag&drop изображения, хранит обложку проекта.
    - LogDock — док-панель лога с управлением высотой/сохранением.
    - FileDropLabel — упрощённый drop для файлов.

Стиль:
    - Нумерованные секции и краткие комментарии.
"""

# 1. Импорт
from PySide6 import QtWidgets, QtGui, QtCore
from pathlib import Path
import shutil
from .common import ASSETS_DIR, to_float

# 2. SmartDoubleSpinBox
class SmartDoubleSpinBox(QtWidgets.QDoubleSpinBox):
    def textFromValue(self, value: float) -> str:
        s = f"{float(value):.{self.decimals()}f}".replace(".", ",").rstrip("0").rstrip(",")
        return s or "0"
    def valueFromText(self, text: str) -> float:
        return to_float(text, 0.0)
    def validate(self, text: str, pos: int):
        # Принимаем запятую как десятичный разделитель
        try:
            _ = to_float(text, 0.0)
            return (QtGui.QValidator.Acceptable, text, pos)
        except Exception:
            return (QtGui.QValidator.Intermediate, text, pos)

# 3. Виджет предпросмотра изображения (обложка проекта)

# 3. ImageDropLabel
class ImageDropLabel(QtWidgets.QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setText("Перетащите сюда картинку (PNG/JPG)")
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setAcceptDrops(True)
        self.setMinimumHeight(200)
        self._project_id = None
        self._stored_path = None
        self.setWordWrap(True)
    def set_project_id(self, project_id: int):
        self._project_id = project_id
    def dragEnterEvent(self, e: QtGui.QDragEnterEvent):
        if e.mimeData().hasUrls():
            for u in e.mimeData().urls():
                p = u.toLocalFile().lower()
                if p.endswith((".png", ".jpg", ".jpeg")):
                    e.acceptProposedAction()
                    return
        e.ignore()
    def dropEvent(self, e: QtGui.QDropEvent):
        if self._project_id is None:
            QtWidgets.QMessageBox.information(self, "Внимание", "Сначала выберите проект.")
            e.ignore()
            return
        for u in e.mimeData().urls():
            src = Path(u.toLocalFile())
            if src.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
                continue
            dest_dir = ASSETS_DIR / f"project_{self._project_id}"
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / f"cover{src.suffix.lower()}"
            try:
                shutil.copy2(src, dest)
                self._stored_path = str(dest)
                self._load_pixmap(dest)
                e.acceptProposedAction()
                return
            except Exception as ex:
                QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось скопировать: {ex}")
                e.ignore()
                return
        e.ignore()
    def _load_pixmap(self, path: Path):
        pm = QtGui.QPixmap(str(path))
        if pm.isNull():
            self.setText("Не удалось загрузить изображение")
            return
        self.setPixmap(pm.scaled(self.size(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
    def resizeEvent(self, ev: QtGui.QResizeEvent):
        super().resizeEvent(ev)
        if self.pixmap() and self._stored_path:
            self._load_pixmap(Path(self._stored_path))

# 4. Док-панель «Лог»

# 4. LogDock
class LogDock(QtWidgets.QDockWidget):
    resized = QtCore.Signal(float)
    saveRequested = QtCore.Signal(float)
    COLLAPSED_RATIO = 0.15
    EXPANDED_RATIO = 0.50
    def __init__(self, parent=None):
        super().__init__("Лог", parent)
        self.setAllowedAreas(QtCore.Qt.BottomDockWidgetArea)
        self.view = QtWidgets.QTextEdit()
        self.view.setReadOnly(True)
        self.view.setStyleSheet("QTextEdit { background: #1e1e1e; color: #dddddd; }")
        self.setWidget(self.view)
        # Заголовок
        self._title_widget = QtWidgets.QWidget()
        h = QtWidgets.QHBoxLayout(self._title_widget); h.setContentsMargins(6, 2, 6, 2)
        title_lbl = QtWidgets.QLabel("Лог"); title_lbl.setStyleSheet("QLabel { font-weight: 600; }")
        self.chk_expand = QtWidgets.QCheckBox("Развернуть лог")
        self.chk_remember = QtWidgets.QCheckBox("Запомнить размер")
        h.addWidget(title_lbl); h.addStretch(1); h.addWidget(self.chk_expand); h.addWidget(self.chk_remember)
        self.setTitleBarWidget(self._title_widget)
        self.chk_expand.toggled.connect(self._on_expand_toggled)
        self.chk_remember.toggled.connect(self._on_remember_toggled)
        self._current_ratio = self.COLLAPSED_RATIO
    def apply_initial_state(self, ratio: float, expanded: bool):
        self.chk_expand.blockSignals(True)
        self.chk_expand.setChecked(expanded)
        self.chk_expand.blockSignals(False)
        self._current_ratio = max(0.08, min(0.8, float(ratio or self.COLLAPSED_RATIO)))
        self.resized.emit(self._desired_ratio())
    def _desired_ratio(self) -> float:
        return self.EXPANDED_RATIO if self.chk_expand.isChecked() else self._current_ratio
    def _on_expand_toggled(self, _checked: bool):
        self.resized.emit(self._desired_ratio())
    def _on_remember_toggled(self, checked: bool):
        if not checked:
            return
        mw = self.parent()
        if not isinstance(mw, QtWidgets.QMainWindow):
            return
        total_h = max(1, mw.size().height())
        dock_h = max(1, self.size().height())
        ratio = dock_h / total_h
        self._current_ratio = max(0.08, min(0.8, ratio))
        self.saveRequested.emit(self._current_ratio)

# 5. Делегаты таблиц (класс, перенос текста)

# 5. FileDropLabel
class FileDropLabel(QtWidgets.QLabel):
    def __init__(self, accept_exts: tuple[str, ...], on_file, parent=None):
        super().__init__(parent)
        self.setText("Перетащите сюда файл сметы (XLSX/CSV)")
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setAcceptDrops(True)
        self.accept_exts = tuple(e.lower() for e in accept_exts)
        self.on_file = on_file
    def dragEnterEvent(self, e: QtGui.QDragEnterEvent):
        if e.mimeData().hasUrls():
            for u in e.mimeData().urls():
                p = u.toLocalFile().lower()
                if any(p.endswith(ext) for ext in self.accept_exts):
                    e.acceptProposedAction(); return
        e.ignore()
    def dropEvent(self, e: QtGui.QDropEvent):
        for u in e.mimeData().urls():
            src = Path(u.toLocalFile())
            if src.suffix.lower() in self.accept_exts:
                self.setText(str(src))
                if callable(self.on_file): self.on_file(src)
                e.acceptProposedAction(); return
        e.ignore()
