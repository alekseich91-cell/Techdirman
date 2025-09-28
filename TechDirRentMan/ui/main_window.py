"""
Назначение:
    Главное окно приложения TechDirRentMan: список проектов, вкладки проекта и док-панель лога.

Как работает:
    - Создаёт/копирует/удаляет проекты, открывает окно базы данных.
    - Инициализирует ProjectPage и LogDock, управляет сохранением (commit).
    - Делегирует операции в DB.

Стиль:
    - Нумерованные секции + краткие комментарии.
"""

# 1. Импорт
# 1. Импорт
from PySide6 import QtWidgets, QtGui, QtCore  # Qt
from typing import Optional, Callable          # типы
from pathlib import Path                       # для путей при копировании файлов
import shutil, json                            # для копирования и обработки файлов

from .project_page import ProjectPage         # страница проекта
from .db_window import DatabaseWindow         # окно каталога
from .widgets import LogDock                  # док-панель лога
from db import DB                             # база данных
from .common import ASSETS_DIR, DATA_DIR      # директории ассетов и данных

# 2. Класс MainWindow
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, db: DB, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TechDirRentMan — базовый каркас")
        self.resize(1360, 860)
        self.db = db

        # 6.1 Верхняя панель — красная кнопка «Сохранить изменения»
        self.toolbar = QtWidgets.QToolBar()
        self.toolbar.setMovable(False)
        self.addToolBar(QtCore.Qt.TopToolBarArea, self.toolbar)
        spacer = QtWidgets.QWidget(); spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        self.toolbar.addWidget(spacer)
        self.btn_save_global = QtWidgets.QPushButton("Сохранить изменения")
        self.btn_save_global.setStyleSheet("QPushButton { background-color:#d9534f; color:white; padding:6px 12px; border-radius:4px; }"
                                           "QPushButton:hover { background-color:#c9302c; }")
        self.btn_save_global.clicked.connect(self._global_save)
        self.toolbar.addWidget(self.btn_save_global)

        # 6.2 Левый сайдбар проектов
        left = QtWidgets.QWidget(); left_layout = QtWidgets.QVBoxLayout(left)
        self.list_projects = QtWidgets.QListWidget()
        self.btn_add = QtWidgets.QPushButton("Создать проект")
        self.btn_copy = QtWidgets.QPushButton("Сделать копию проекта")
        # Кнопка переименования проекта (добавлена в версии расширения)
        self.btn_rename = QtWidgets.QPushButton("Переименовать проект")
        self.btn_db = QtWidgets.QPushButton("Перейти в базу данных")
        self.btn_del = QtWidgets.QPushButton("Удалить проект")
        left_layout.addWidget(QtWidgets.QLabel("Проекты"))
        left_layout.addWidget(self.list_projects, 1)
        # Добавляем кнопки управления проектами. Кнопка переименования
        # помещается между созданием/копированием и переходом в базу.
        left_layout.addWidget(self.btn_add)
        left_layout.addWidget(self.btn_copy)
        left_layout.addWidget(self.btn_rename)
        left_layout.addWidget(self.btn_db)
        left_layout.addWidget(self.btn_del)

        # 6.3 Правая часть — стек страниц
        self.stack = QtWidgets.QStackedWidget()
        self.page_empty = QtWidgets.QWidget()
        self.page_project = ProjectPage(db=self.db, log_fn=self.log)
        self.stack.addWidget(self.page_empty); self.stack.addWidget(self.page_project)

        splitter = QtWidgets.QSplitter()
        splitter.addWidget(left); splitter.addWidget(self.stack)
        splitter.setStretchFactor(0, 0); splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

        # 6.4 Док «Лог»
        self.log_dock = LogDock(self)
        self.addDockWidget(QtCore.Qt.BottomDockWidgetArea, self.log_dock)
        self.log_dock.resized.connect(self._apply_log_ratio)
        self.log_dock.saveRequested.connect(self._save_log_default)

        # 6.5 Настройки (размер лога)
        self.settings = QtCore.QSettings("TechDirRentMan", "TDRM")
        ratio = float(self.settings.value("log/height_ratio", LogDock.COLLAPSED_RATIO))
        expanded = bool(self.settings.value("log/expanded", False))
        remember = bool(self.settings.value("log/remember", False))
        QtCore.QTimer.singleShot(0, lambda: self.log_dock.apply_initial_state(ratio, expanded))
        self.log_dock.chk_remember.setChecked(remember)

        # 6.6 Сигналы сайдбара
        self.btn_add.clicked.connect(self.create_project)
        self.btn_copy.clicked.connect(self.copy_project)
        # Обработчик кнопки переименования проекта
        self.btn_rename.clicked.connect(self.rename_project)
        self.btn_db.clicked.connect(self.open_database_window)
        self.btn_del.clicked.connect(self.delete_project)
        self.list_projects.itemSelectionChanged.connect(self.open_selected_project)

        # 6.7 Загрузка проектов
        self.reload_projects()
        self.log("Приложение запущено.")

    # 6.8 Сохранение по красной кнопке
    def _global_save(self):
        try:
            self.db.commit()
            self.log("Сохранение: изменения записаны.")
            QtWidgets.QMessageBox.information(self, "Сохранено", "Изменения сохранены.")
        except Exception as ex:
            self.log(f"Ошибка сохранения: {ex}", "error")
            QtWidgets.QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить: {ex}")

    # 6.9 Применение доли высоты лога
    def _apply_log_ratio(self, ratio: float):
        ratio = max(0.08, min(0.8, float(ratio or LogDock.COLLAPSED_RATIO)))
        total_h = max(1, self.size().height())
        dock_h = max(60, int(total_h * ratio))
        try:
            self.resizeDocks([self.log_dock], [dock_h], QtCore.Qt.Vertical)
        except Exception:
            pass
        self.settings.setValue("log/expanded", self.log_dock.chk_expand.isChecked())

    # 6.10 Сохранить размер лога по умолчанию
    def _save_log_default(self, ratio: float):
        self.settings.setValue("log/height_ratio", float(ratio))
        self.settings.setValue("log/remember", True)
        self.settings.setValue("log/expanded", self.log_dock.chk_expand.isChecked())
        self.log("Размер лог-панели сохранён как основной.")

    # 6.11 Лог (в док + stdout)
    def log(self, msg: str, level: str = "info"):
        """
        Логирует сообщение в док‑панель и stdout.

        Если док‑панель ещё не создана (в начале инициализации),
        сообщение выводится только в stdout. После создания панели
        сообщения также добавляются в виджет лога.

        :param msg: текст сообщения
        :param level: уровень (info/error)
        """
        # Формируем HTML‑строку в зависимости от уровня
        if level == "error":
            html = f'<span style="color:#ff6b6b;">[ERROR]</span> {QtGui.QGuiApplication.translate("ui", msg)}'
        else:
            html = f'<span style="color:#6fbf73;">[INFO]</span> {QtGui.QGuiApplication.translate("ui", msg)}'
        # Если док‑панель уже создана, добавляем запись в неё
        log_dock = getattr(self, "log_dock", None)
        if log_dock is not None and hasattr(log_dock, "view"):
            try:
                log_dock.view.append(html)
                log_dock.view.moveCursor(QtGui.QTextCursor.End)
            except Exception:
                pass
        # Выводим в stdout независимо от наличия панели
        try:
            print(f"{level.upper()}: {msg}")
        except Exception:
            pass

    # 6.12 Операции с проектами
    def reload_projects(self):
        """
        Перезагружает список проектов. Помимо названия и даты создания
        окрашивает элементы списка в зависимости от статуса проекта:
            • «В работе»  → красный цвет
            • «Завершен» → зелёный цвет
            • «Резерв»   → синий цвет
            • «Тестовый»→ системный цвет (по умолчанию)
        """
        self.list_projects.clear()
        # Получаем все проекты; столбец 'status' добавлен в схему
        for r in self.db.list_projects():
            name = r["name"]
            created_at = r["created_at"]
            status = r["status"] if "status" in r.keys() else None
            it = QtWidgets.QListWidgetItem(f"{name}  ({created_at})")
            it.setData(QtCore.Qt.UserRole, r["id"])
            try:
                col = None
                if status == "В работе":
                    col = QtGui.QColor("red")
                elif status == "Завершен":
                    col = QtGui.QColor("green")
                elif status == "Резерв":
                    col = QtGui.QColor("blue")
                if col is not None:
                    it.setForeground(col)
            except Exception:
                pass
            self.list_projects.addItem(it)

    def create_project(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Новый проект", "Введите название:")
        if not ok or not name.strip():
            return
        pid = self.db.add_project(name.strip())
        self.reload_projects()
        for i in range(self.list_projects.count()):
            if self.list_projects.item(i).data(QtCore.Qt.UserRole) == pid:
                self.list_projects.setCurrentRow(i); break
        self.log(f"Создан проект '{name.strip()}'")

    def copy_project(self):
        it = self.list_projects.currentItem()
        if not it:
            QtWidgets.QMessageBox.information(self, "Внимание", "Выберите проект для копирования.")
            return
        src_pid = it.data(QtCore.Qt.UserRole)
        src_name = it.text().split("  (")[0]
        new_name, ok = QtWidgets.QInputDialog.getText(self, "Сделать копию проекта",
                                                      f"Введите название копии для «{src_name}»:",
                                                      text=f"{src_name} — копия")
        if not ok or not new_name.strip():
            return
        new_pid = self.db.add_project(new_name.strip())
        # 1. Копируем все позиции проекта в новую запись
        rows = self.db.list_items(src_pid)
        items = [{
            "project_id": new_pid,
            "type": r["type"],
            "group_name": r["group_name"],
            "name": r["name"],
            "qty": r["qty"],
            "coeff": r["coeff"],
            "amount": r["amount"],
            "unit_price": r["unit_price"],
            "source_file": r["source_file"],
            "vendor": r["vendor"] or "",
            "department": r["department"] or "",
            "zone": r["zone"] or "",
            "power_watts": r["power_watts"] or 0,
            "import_batch": None,
        } for r in rows]
        if items:
            self.db.add_items_bulk(items)
        # 2. Копируем связанные файлы и снимки
        try:
            # Директории ассетов для исходного и нового проекта
            src_assets = ASSETS_DIR / f"project_{src_pid}"
            dst_assets = ASSETS_DIR / f"project_{new_pid}"
            if src_assets.exists():
                # Рекурсивно копируем папку. В случае существования цели
                # разрешаем перезапись (dirs_exist_ok=True используется в python>=3.8)
                shutil.copytree(src_assets, dst_assets, dirs_exist_ok=True)
                # Обновляем файлы снимков: переименовываем project_id в имени и JSON
                snap_dir = dst_assets / "snapshots"
                if snap_dir.exists():
                    pattern = f"project_{src_pid}_"
                    for snap_file in list(snap_dir.glob(f"project_{src_pid}_*.json")):
                        try:
                            with open(snap_file, "r", encoding="utf-8") as fh:
                                snap_data = json.load(fh)
                            # обновляем идентификатор проекта в содержимом
                            snap_data["project_id"] = new_pid
                            # формируем новое имя файла с новым project_id
                            new_name_part = snap_file.name.replace(pattern, f"project_{new_pid}_")
                            new_path = snap_file.parent / new_name_part
                            with open(new_path, "w", encoding="utf-8") as fh:
                                json.dump(snap_data, fh, ensure_ascii=False, indent=2)
                            # удаляем старый файл, если имя изменилось
                            if new_path != snap_file:
                                snap_file.unlink()
                        except Exception as ex:
                            self.log(f"Ошибка обработки снимка {snap_file}: {ex}", "error")
            # Копируем файл с информацией о проекте (info_json).
            # Для сохранения информации используется функция info_json_path,
            # которая выбирает директорию в зависимости от того, где находится БД.
            try:
                from .info_tab import info_json_path as info_json_path_ext  # type: ignore
                # Создаём временные «страницы» с необходимыми атрибутами project_id и db
                class _Tmp:
                    pass
                src_page = _Tmp()
                src_page.project_id = src_pid
                src_page.db = self.db
                dst_page = _Tmp()
                dst_page.project_id = new_pid
                dst_page.db = self.db
                src_info = info_json_path_ext(src_page)
                dst_info = info_json_path_ext(dst_page)
            except Exception:
                # Если info_json_path недоступна, откатываемся к локальной директории данных
                src_info = DATA_DIR / f"project_{src_pid}_info.json"
                dst_info = DATA_DIR / f"project_{new_pid}_info.json"
            try:
                if src_info.exists():
                    dst_info.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src_info, dst_info)
            except Exception as ex:
                self.log(f"Ошибка копирования файла информации: {ex}", "error")
            # Копируем конфигурации тайминга и финансовых настроек
            try:
                timing_json = self.db.get_project_timing(src_pid)
                if timing_json:
                    self.db.set_project_timing(new_pid, timing_json)
                finance_json = self.db.get_project_finance(src_pid)
                if finance_json:
                    self.db.set_project_finance(new_pid, finance_json)
            except Exception as ex:
                self.log(f"Ошибка копирования настроек проекта: {ex}", "error")
        except Exception as ex:
            # Логируем общую ошибку копирования ассетов
            self.log(f"Ошибка копирования файлов проекта: {ex}", "error")
        # 3. Обновляем список проектов и выбираем только что созданный
        self.reload_projects()
        for i in range(self.list_projects.count()):
            if self.list_projects.item(i).data(QtCore.Qt.UserRole) == new_pid:
                self.list_projects.setCurrentRow(i)
                break
        # 4. Информируем пользователя
        self.log(f"Создана копия проекта '{src_name}' → '{new_name.strip()}'")

    def rename_project(self):
        """
        Переименовывает выбранный проект. Запрашивает новое название у
        пользователя, затем обновляет запись в базе данных и логирует
        результат. Если пользователь не выбрал проект или отменил ввод,
        операция не выполняется. При конфликте имён отображается
        предупреждающее сообщение.
        """
        it = self.list_projects.currentItem()
        if not it:
            QtWidgets.QMessageBox.information(self, "Внимание", "Выберите проект для переименования.")
            return
        pid = it.data(QtCore.Qt.UserRole)
        old_name = it.text().split("  (")[0]
        new_name, ok = QtWidgets.QInputDialog.getText(
            self,
            "Переименовать проект",
            f"Введите новое название для «{old_name}»:",
            text=old_name,
        )
        if not ok or not new_name.strip() or new_name.strip() == old_name:
            return
        try:
            self.db.rename_project(pid, new_name.strip())
            # Обновляем список и выделяем переименованный проект
            self.reload_projects()
            for i in range(self.list_projects.count()):
                if self.list_projects.item(i).data(QtCore.Qt.UserRole) == pid:
                    self.list_projects.setCurrentRow(i)
                    break
            # Если открыт этот проект — обновляем заголовок страницы
            if isinstance(self.stack.currentWidget(), ProjectPage) and self.page_project.project_id == pid:
                self.page_project.project_name = new_name.strip()
            self.log(f"Проект '{old_name}' переименован в '{new_name.strip()}'")
        except Exception as ex:
            # Обрабатываем возможный конфликт имён или другие ошибки
            self.log(f"Ошибка переименования проекта: {ex}", "error")
            QtWidgets.QMessageBox.critical(
                self,
                "Ошибка",
                f"Не удалось переименовать проект: {ex}",
            )

    def open_database_window(self):
        pid = self.page_project.project_id if isinstance(self.stack.currentWidget(), ProjectPage) else None
        reload_cb: Optional[Callable] = None
        if pid is not None:
            reload_cb = self.page_project._reload_zone_tabs
        dlg = DatabaseWindow(self.db, parent=self, log_fn=self.log,
                             project_id_for_sync=pid, reload_summary_cb=reload_cb)
        dlg.exec()

    def delete_project(self):
        it = self.list_projects.currentItem()
        if not it: return
        if QtWidgets.QMessageBox.question(self, "Подтверждение", "Удалить выбранный проект?") != QtWidgets.QMessageBox.Yes:
            return
        self.db.delete_project(it.data(QtCore.Qt.UserRole))
        self.reload_projects()
        self.stack.setCurrentWidget(self.page_empty)
        self.log("Проект удалён")

    def open_selected_project(self):
        it = self.list_projects.currentItem()
        if not it:
            self.stack.setCurrentWidget(self.page_empty)
            return
        pid = it.data(QtCore.Qt.UserRole); name = it.text().split("  (")[0]
        self.page_project.load_project(pid, name)
        self.stack.setCurrentWidget(self.page_project)
        self.log(f"Открыт проект '{name}'")

# 7. Окно «База данных (каталог)»
