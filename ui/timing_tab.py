"""
Модуль `timing_tab` реализует современную и наглядную вкладку «Тайминг» на основе
графического таймлайна.  В отличие от прошлых реализаций, временные блоки
представлены абсолютными минутами от начала проекта и отображаются на
`QGraphicsView`‑холсте.  Блоки автоматически размещаются без пересечений
(рефлоу) и могут быть добавлены хоть в каждую минуту.  Тайминг надёжно
привязывается к текущему проекту: при открытии проекта данные загружаются
автоматически, а при изменении — сохраняются в БД.

Основные принципы:

1. **Минутная модель.** Блоки хранятся в абсолютных минутах (`start_min`) и
   длительности (`duration_min`), а не в «строках» таблицы.  Строки сетки
   служат лишь ориентиром.
2. **Автоматическое размещение.** При вставке или редактировании блок
   «вклеивается» в колонку и при необходимости сдвигается вперёд, чтобы
   избежать наложений.  Нет отказов из‑за пересечений.
3. **Наглядные колонки.** Слева отображается список колонок, которые можно
   добавлять, переименовывать и удалять.  Каждая колонка — дорожка на
   таймлайне.
4. **Контекстное меню.** Правая кнопка по таймлайну открывает меню для
   добавления, редактирования, удаления блока и добавления следующего блока
   сразу после выбранного.  Возможные действия зависят от того, попали ли вы
   в существующий блок.
5. **Привязка к проекту.** Данные тайминга хранятся в поле
   `projects.timing_json` для каждого `project_id`.  Вкладка автоматически
   загружает тайминг при открытии проекта и сохраняет изменения без лишних
   нажатий кнопок.
6. **Синхронизация с экспортом.** Вкладка предоставляет методы
   `timing_get_json()` и `timing_export_image(path, width_px)` для экспорта
   тайминга (например, в PDF‑экспорт), а также `timing_reload_for_current_project()`
   для внешнего обновления.

Код разбит на пронумерованные секции и снабжён краткими комментариями.
Все ключевые действия (загрузка, сохранение, добавление/редактирование
блоков, операции с колонками) записываются в лог через `page._log`, если он
определён.
"""

from __future__ import annotations

# 1. Импорт стандартных и Qt‑модулей
from dataclasses import dataclass, asdict
from typing import List, Optional, Dict, Any
import json

from PySide6 import QtWidgets, QtCore, QtGui

from .timeline_canvas import TimelineView, EventModel, HEADER_HEIGHT


# 2. Структура временного блока для экспорта в старом формате
@dataclass
class TimingBlock:
    """Структура для совместимости со старым экспортом (по строкам).

    Хотя вкладка работает с минутами, некоторые модули (например, экспорт в PDF)
    ожидают старый формат с атрибутами `start_min`, `duration_min`, `col`,
    `title` и `color`.  Поэтому при изменении данных мы обновляем
    `page.timing_blocks` списком таких объектов.
    """
    start_min: int
    duration_min: int
    col: int
    title: str
    color: str = "#6fbf73"


# 3. Диалог создания/редактирования блоков
class BlockDialog(QtWidgets.QDialog):
    """Диалог для создания или редактирования события на таймлайне.

    Пользователь выбирает день, колонку, время начала, длительность, заголовок
    и цвет.  Внизу отображается автоматически вычисляемое время окончания.
    """

    def __init__(
        self,
        parent: Optional[QtWidgets.QWidget],
        day_labels: List[str],
        columns: List[str],
        step_minutes: int,
        existing: Optional[EventModel] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Событие тайминга")
        self._step = max(1, int(step_minutes))
        self._duration_default = self._step
        # 3.1 Виджеты формы
        layout = QtWidgets.QFormLayout(self)
        self.cbo_day = QtWidgets.QComboBox(); self.cbo_day.addItems(day_labels)
        self.cbo_col = QtWidgets.QComboBox(); self.cbo_col.addItems(columns)
        self.time_start = QtWidgets.QTimeEdit(); self.time_start.setDisplayFormat("HH:mm")
        self.time_start.setMinimumTime(QtCore.QTime(0, 0)); self.time_start.setMaximumTime(QtCore.QTime(23, 59))
        self.spin_dur = QtWidgets.QSpinBox(); self.spin_dur.setRange(1, 7 * 24 * 60)
        self.spin_dur.setSingleStep(self._step); self.spin_dur.setValue(self._duration_default)
        self.edit_title = QtWidgets.QLineEdit()
        self.btn_color = QtWidgets.QPushButton("Цвет")
        self.lbl_end = QtWidgets.QLabel("")
        self._color = "#6fbf73"
        # 3.2 Выбор цвета
        def pick_color() -> None:
            c = QtWidgets.QColorDialog.getColor(QtGui.QColor(self._color), self, "Цвет события")
            if c.isValid():
                self._color = c.name()
        self.btn_color.clicked.connect(pick_color)
        # 3.3 Обновление времени окончания
        def update_end() -> None:
            start_min = self.time_start.time().hour() * 60 + self.time_start.time().minute()
            dur = self.spin_dur.value()
            end_min = start_min + dur
            h = (end_min // 60) % 24; m = end_min % 60
            self.lbl_end.setText(f"Конец: {h:02d}:{m:02d}")
        self.time_start.timeChanged.connect(lambda *_: update_end())
        self.spin_dur.valueChanged.connect(lambda *_: update_end())
        # 3.4 Сборка формы
        layout.addRow("День:", self.cbo_day)
        layout.addRow("Колонка:", self.cbo_col)
        layout.addRow("Начало:", self.time_start)
        layout.addRow("Длительность (мин):", self.spin_dur)
        layout.addRow("Название:", self.edit_title)
        layout.addRow("Цвет:", self.btn_color)
        layout.addRow(self.lbl_end)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addRow(btn_box)
        btn_box.accepted.connect(self.accept); btn_box.rejected.connect(self.reject)
        # 3.5 Инициализация существующего события
        if existing:
            mins = existing.start_min
            day_idx = mins // (24 * 60)
            min_in_day = mins % (24 * 60)
            self.cbo_day.setCurrentIndex(day_idx if 0 <= day_idx < len(day_labels) else 0)
            self.time_start.setTime(QtCore.QTime((min_in_day // 60) % 24, min_in_day % 60))
            self.spin_dur.setValue(int(existing.duration_min))
            self.cbo_col.setCurrentIndex(existing.col)
            self.edit_title.setText(existing.title)
            self._color = existing.color or self._color
        update_end()

    def get_event(self) -> EventModel:
        """Возвращает модель события из введённых данных."""
        start_time = self.time_start.time()
        start_min = self.cbo_day.currentIndex() * 24 * 60 + start_time.hour() * 60 + start_time.minute()
        duration_min = max(1, int(self.spin_dur.value()))
        return EventModel(
            col=self.cbo_col.currentIndex(),
            start_min=start_min,
            duration_min=duration_min,
            title=self.edit_title.text().strip(),
            color=self._color or "#6fbf73",
        )


# 4. Основная функция: построение вкладки тайминга
def build_timing_tab(page: Any, tab: QtWidgets.QWidget) -> None:
    """Формирует вкладку «Тайминг» в интерфейсе проекта.

    Создаёт графический таймлайн, список колонок и панель параметров.  Загружает
    данные из БД при открытии проекта и сохраняет их при изменении.  Вкладка
    предоставляет функции `timing_get_json()`, `timing_export_image()` и
    `timing_reload_for_current_project()` для использования во внешних
    модулях (например, экспорт PDF).
    """
    # 4.1 Инициализация параметров проекта
    page.timing_days = int(getattr(page, "timing_days", 3))
    page.timing_step_minutes = int(getattr(page, "timing_step_minutes", 60))
    page.timing_start_date = getattr(page, "timing_start_date", QtCore.QDate.currentDate())
    page.timing_column_names: List[str] = list(getattr(page, "timing_column_names", ["Общий"]))
    # список событий текущего проекта; будем хранить dict для быстрого сохранения
    page.timing_events: List[Dict[str, Any]] = []
    # устаревший формат для экспортных модулей
    page.timing_blocks: List[TimingBlock] = []
    # 4.2 Виджет таймлайна
    timeline = TimelineView(tab)
    timeline.set_columns(page.timing_column_names)
    timeline.set_step(page.timing_step_minutes)
    timeline.set_days(page.timing_days)

    # 4.3 Список колонок и кнопки управления колонками
    col_list = QtWidgets.QListWidget()
    col_list.addItems(page.timing_column_names)
    col_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
    # фиксированная ширина списка колонок для компактного интерфейса
    col_list.setFixedWidth(160)
    btn_add_col = QtWidgets.QPushButton("+ Колонка")
    btn_ren_col = QtWidgets.QPushButton("Переименовать")
    btn_del_col = QtWidgets.QPushButton("Удалить")

    # 4.4 Панель параметров
    spin_days = QtWidgets.QSpinBox(); spin_days.setRange(1, 60); spin_days.setValue(page.timing_days)
    combo_step = QtWidgets.QComboBox(); combo_step.addItems(["1", "5", "10", "15", "30", "60", "90", "120"])
    # установим значение шага, если имеется
    combo_step.setCurrentText(str(page.timing_step_minutes)) if str(page.timing_step_minutes) in ["1","5","10","15","30","60","90","120"] else combo_step.setCurrentText("15")
    date_start = QtWidgets.QDateEdit(); date_start.setCalendarPopup(True); date_start.setDate(page.timing_start_date)
    btn_add_event = QtWidgets.QPushButton("Добавить событие")
    btn_save = QtWidgets.QPushButton("Сохранить")

    # 4.5 Сборка интерфейса
    left_panel = QtWidgets.QVBoxLayout()
    left_panel.addWidget(QtWidgets.QLabel("Колонки"))
    left_panel.addWidget(col_list, 1)
    btn_col_bar = QtWidgets.QHBoxLayout()
    btn_col_bar.addWidget(btn_add_col); btn_col_bar.addWidget(btn_ren_col); btn_col_bar.addWidget(btn_del_col)
    left_panel.addLayout(btn_col_bar)
    left_panel.addStretch(1)

    top_bar = QtWidgets.QHBoxLayout()
    top_bar.addWidget(QtWidgets.QLabel("Дней:")); top_bar.addWidget(spin_days)
    top_bar.addWidget(QtWidgets.QLabel("Шаг (мин):")); top_bar.addWidget(combo_step)
    top_bar.addWidget(QtWidgets.QLabel("Начало:")); top_bar.addWidget(date_start)
    top_bar.addStretch(1)
    top_bar.addWidget(btn_add_event)
    top_bar.addWidget(btn_save)

    right_panel = QtWidgets.QVBoxLayout()
    right_panel.addLayout(top_bar)
    right_panel.addWidget(timeline, 1)

    main_layout = QtWidgets.QHBoxLayout(tab)
    main_layout.addLayout(left_panel, 0)
    main_layout.addLayout(right_panel, 1)

    # 4.6 Вспомогательные функции
    def _to_blocks(events: List[Dict[str, Any]]) -> List[TimingBlock]:
        """Преобразует список событий в список TimingBlock для экспорта."""
        blocks: List[TimingBlock] = []
        for ev in events:
            blocks.append(TimingBlock(ev["start_min"], ev["duration_min"], ev["col"], ev.get("title", ""), ev.get("color", "#6fbf73")))
        return blocks

    def _sync_events_from_timeline() -> None:
        """Синхронизирует `page.timing_events` и `page.timing_blocks` с данными таймлайна."""
        data = timeline.export_events()
        page.timing_events = data
        page.timing_blocks = _to_blocks(data)

    def _save() -> None:
        """Сохраняет текущие настройки и события в БД."""
        pid = getattr(page, "project_id", None)
        db = getattr(page, "db", None)
        if not pid or not db:
            return
        try:
            _sync_events_from_timeline()
            payload = {
                "start_date": date_start.date().toString("yyyy-MM-dd"),
                "days": int(spin_days.value()),
                "step_minutes": int(combo_step.currentText()),
                "columns": page.timing_column_names,
                "events": page.timing_events,
            }
            db.set_project_timing(pid, json.dumps(payload, ensure_ascii=False))
            if hasattr(page, "_log"):
                page._log("Тайминг сохранён.")
        except Exception as ex:
            if hasattr(page, "_log"):
                page._log(f"Ошибка сохранения тайминга: {ex}", "error")

    def _load() -> None:
        """Загружает тайминг для текущего проекта."""
        pid = getattr(page, "project_id", None)
        db = getattr(page, "db", None)
        # очистка
        page.timing_events = []
        page.timing_blocks = []
        if not pid or not db:
            timeline.clear_events()
            if hasattr(page, "_log"):
                page._log("Проект не выбран или БД недоступна.")
            return
        try:
            raw = db.get_project_timing(pid)
            if not raw:
                # нет сохранённых данных для проекта
                timeline.clear_events()
                if hasattr(page, "_log"):
                    page._log("Тайминг для проекта не найден, начато с пустого.")
                return
            data = json.loads(raw)
            # присваиваем параметры
            page.timing_start_date = QtCore.QDate.fromString(data.get("start_date", ""), "yyyy-MM-dd") or QtCore.QDate.currentDate()
            page.timing_days = int(data.get("days", spin_days.value()))
            page.timing_step_minutes = int(data.get("step_minutes", combo_step.currentText()))
            page.timing_column_names = list(data.get("columns", ["Общий"])) or ["Общий"]
            # обновляем виджеты
            spin_days.setValue(page.timing_days)
            if str(page.timing_step_minutes) in [combo_step.itemText(i) for i in range(combo_step.count())]:
                combo_step.setCurrentText(str(page.timing_step_minutes))
            else:
                combo_step.setCurrentText(str(page.timing_step_minutes))
            date_start.setDate(page.timing_start_date)
            # обновляем список колонок
            col_list.blockSignals(True)
            col_list.clear(); col_list.addItems(page.timing_column_names)
            col_list.blockSignals(False)
            # обновляем таймлайн
            timeline.set_columns(page.timing_column_names)
            timeline.set_step(page.timing_step_minutes)
            timeline.set_days(page.timing_days)
            events_data = data.get("events")
            # поддержка старого формата `blocks`
            if events_data is None and "blocks" in data:
                events_data = data.get("blocks")
            if not isinstance(events_data, list):
                events_data = []
            timeline.load_events(events_data)
            _sync_events_from_timeline()
            if hasattr(page, "_log"):
                page._log("Тайминг загружен.")
        except Exception as ex:
            if hasattr(page, "_log"):
                page._log(f"Ошибка загрузки тайминга: {ex}", "error")
        # обновляем высоту после загрузки
        timeline.set_days(spin_days.value())

    # 4.7 Обработчики изменений виджетов
    def _apply_settings() -> None:
        """Применяет изменённые дни, шаг, дату начала и обновляет таймлайн."""
        page.timing_days = spin_days.value()
        page.timing_step_minutes = int(combo_step.currentText())
        page.timing_start_date = date_start.date()
        timeline.set_days(page.timing_days)
        timeline.set_step(page.timing_step_minutes)
        # корректируем существующие события под новый шаг: округляем время и длительность
        events = timeline.export_events()
        for ev in events:
            # округлим время к ближайшему значению, кратному новому шагу
            step = page.timing_step_minutes
            ev["start_min"] = int(round(ev["start_min"] / step)) * step
            ev["duration_min"] = max(step, int(round(ev["duration_min"] / step)) * step)
        timeline.load_events(events)
        _sync_events_from_timeline()
        _save()

    spin_days.valueChanged.connect(_apply_settings)
    combo_step.currentIndexChanged.connect(_apply_settings)
    date_start.dateChanged.connect(_apply_settings)

    # 4.8 Колонки: добавление, переименование, удаление
    def _add_column() -> None:
        text, ok = QtWidgets.QInputDialog.getText(tab, "Новая колонка", "Название колонки:")
        if ok and text.strip():
            page.timing_column_names.append(text.strip())
            col_list.addItem(text.strip())
            timeline.set_columns(page.timing_column_names)
            _sync_events_from_timeline(); _save()
            if hasattr(page, "_log"):
                page._log(f"Добавлена колонка '{text.strip()}'.")
    btn_add_col.clicked.connect(_add_column)

    def _rename_column() -> None:
        row = col_list.currentRow()
        if row < 0:
            return
        old_name = page.timing_column_names[row]
        new_name, ok = QtWidgets.QInputDialog.getText(tab, "Переименование колонки", "Новое название:", text=old_name)
        if ok and new_name.strip():
            page.timing_column_names[row] = new_name.strip()
            col_list.item(row).setText(new_name.strip())
            timeline.set_columns(page.timing_column_names)
            _save()
            if hasattr(page, "_log"):
                page._log(f"Колонка '{old_name}' переименована в '{new_name.strip()}'.")
    btn_ren_col.clicked.connect(_rename_column)

    def _delete_column() -> None:
        row = col_list.currentRow()
        if row < 0 or len(page.timing_column_names) <= 1:
            QtWidgets.QMessageBox.warning(tab, "Удаление", "Нельзя удалить последнюю колонку.")
            return
        name = page.timing_column_names[row]
        # подтверждение
        if QtWidgets.QMessageBox.question(tab, "Удаление", f"Удалить колонку '{name}'?") != QtWidgets.QMessageBox.Yes:
            return
        # удаляем колонку и сдвигаем события
        page.timing_column_names.pop(row)
        # корректируем события: удаляем из удалённой колонки и смещаем индексы
        events = []
        for ev in timeline.export_events():
            if ev["col"] == row:
                continue  # drop events from deleted column
            if ev["col"] > row:
                ev["col"] -= 1
            events.append(ev)
        timeline.set_columns(page.timing_column_names)
        col_list.takeItem(row)
        timeline.load_events(events)
        _sync_events_from_timeline(); _save()
        if hasattr(page, "_log"):
            page._log(f"Колонка '{name}' удалена.")
    btn_del_col.clicked.connect(_delete_column)

    # 4.9 Добавление события через кнопку
    def _add_event_button() -> None:
        days_labels = [page.timing_start_date.addDays(i).toString("dd MMM") for i in range(page.timing_days)]
        dlg = BlockDialog(tab, days_labels, page.timing_column_names, page.timing_step_minutes, None)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        ev = dlg.get_event()
        timeline.add_event(ev.col, ev.start_min, ev.duration_min, ev.title, ev.color)
        _sync_events_from_timeline(); _save()
        if hasattr(page, "_log"):
            page._log(f"Добавлено событие '{ev.title}'")
    btn_add_event.clicked.connect(_add_event_button)

    # 4.10 Таймлайн: контекстное меню (ПКМ)
    def _show_context_menu(pos: QtCore.QPoint) -> None:
        # переводим позицию в координаты виджета таймлайна
        event = timeline.get_event_at(pos)
        menu = QtWidgets.QMenu(timeline)
        if event:
            act_edit = menu.addAction("Редактировать событие")
            act_del = menu.addAction("Удалить событие")
            act_next = menu.addAction("Добавить следующий")
            selected = menu.exec(timeline.mapToGlobal(pos))
            if selected == act_edit:
                # редактирование
                _edit_event(event)
            elif selected == act_del:
                if QtWidgets.QMessageBox.question(tab, "Удаление", "Удалить событие?") == QtWidgets.QMessageBox.Yes:
                    _delete_event(event)
            elif selected == act_next:
                _add_next_event(event)
        else:
            act_add = menu.addAction("Добавить событие")
            if menu.exec(timeline.mapToGlobal(pos)) == act_add:
                # позиция для нового события: вычисляем день и время
                scene_pos = timeline.mapToScene(pos)
                # Учтём левый отступ: вычисляем индекс колонки после него
                col_index = int((scene_pos.x() - timeline._scene.left_margin) // timeline._scene.col_width)
                # корректируем координату Y на высоту заголовка
                y_val = scene_pos.y() - HEADER_HEIGHT
                # переводим в минуты от начала проекта
                y_min = int(round(y_val / timeline._scene.minute_px))
                if y_min < 0:
                    y_min = 0
                day_index = y_min // (24 * 60)
                start_in_day = y_min % (24 * 60)
                if day_index >= page.timing_days:
                    day_index = page.timing_days - 1
                days_labels = [page.timing_start_date.addDays(i).toString("dd MMM") for i in range(page.timing_days)]
                dlg = BlockDialog(tab, days_labels, page.timing_column_names, page.timing_step_minutes, None)
                dlg.cbo_day.setCurrentIndex(day_index)
                dlg.cbo_col.setCurrentIndex(col_index if 0 <= col_index < len(page.timing_column_names) else 0)
                dlg.time_start.setTime(QtCore.QTime((start_in_day // 60) % 24, start_in_day % 60))
                if dlg.exec() == QtWidgets.QDialog.Accepted:
                    ev = dlg.get_event()
                    timeline.add_event(ev.col, ev.start_min, ev.duration_min, ev.title, ev.color)
                    _sync_events_from_timeline(); _save()
                    if hasattr(page, "_log"):
                        page._log(f"Добавлено событие '{ev.title}'.")
    timeline.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
    timeline.customContextMenuRequested.connect(_show_context_menu)

    # 4.11 Обработчики редактирования, удаления и добавления следующего события
    def _edit_event(item: Any) -> None:
        # item может быть EventItem или EventModel; извлечём модель
        if isinstance(item, EventModel):
            model = item
        elif hasattr(item, "model"):
            model = item.model
        else:
            return
        days_labels = [page.timing_start_date.addDays(i).toString("dd MMM") for i in range(page.timing_days)]
        dlg = BlockDialog(tab, days_labels, page.timing_column_names, page.timing_step_minutes, model)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        new_ev = dlg.get_event()
        # сохраняем идентификатор события, чтобы обновить именно его
        # найдём его в списке событий
        events = timeline.export_events()
        found = None
        for ev in events:
            if ev["start_min"] == model.start_min and ev["duration_min"] == model.duration_min and ev["col"] == model.col and ev.get("title", "") == model.title:
                found = ev
                break
        if found:
            found.update(asdict(new_ev))
            timeline.load_events(events)
            _sync_events_from_timeline(); _save()
            if hasattr(page, "_log"):
                page._log(f"Событие '{model.title}' отредактировано.")

    def _delete_event(item: Any) -> None:
        # удаляем событие из модели
        if hasattr(item, "model"):
            model = item.model
        elif isinstance(item, EventModel):
            model = item
        else:
            return
        events = [ev for ev in timeline.export_events() if not (ev["col"] == model.col and ev["start_min"] == model.start_min and ev["duration_min"] == model.duration_min and ev.get("title", "") == model.title)]
        timeline.load_events(events)
        _sync_events_from_timeline(); _save()
        if hasattr(page, "_log"):
            page._log(f"Событие '{model.title}' удалено.")

    def _add_next_event(item: Any) -> None:
        # создаём новый блок сразу после выбранного
        if hasattr(item, "model"):
            model = item.model
        elif isinstance(item, EventModel):
            model = item
        else:
            return
        next_start = model.start_min + model.duration_min
        day_index = next_start // (24 * 60)
        if day_index >= page.timing_days:
            day_index = page.timing_days - 1
        start_in_day = next_start % (24 * 60)
        days_labels = [page.timing_start_date.addDays(i).toString("dd MMM") for i in range(page.timing_days)]
        dlg = BlockDialog(tab, days_labels, page.timing_column_names, page.timing_step_minutes, None)
        dlg.cbo_day.setCurrentIndex(day_index)
        dlg.cbo_col.setCurrentIndex(model.col)
        dlg.spin_dur.setValue(model.duration_min)
        dlg.time_start.setTime(QtCore.QTime((start_in_day // 60) % 24, start_in_day % 60))
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            new_ev = dlg.get_event()
            new_ev.start_min = next_start
            timeline.add_event(new_ev.col, new_ev.start_min, new_ev.duration_min, new_ev.title, new_ev.color)
            _sync_events_from_timeline(); _save()
            if hasattr(page, "_log"):
                page._log(f"Добавлено следующее событие '{new_ev.title}'.")

    # 4.12 Сигнал обновления таймлайна
    def _timeline_changed() -> None:
        _sync_events_from_timeline()
        _save()
    timeline.changed.connect(_timeline_changed)

    # 4.13 Кнопка сохранения
    btn_save.clicked.connect(_save)

    # 4.14 Публичные методы страницы для интеграции
    def timing_get_json() -> Dict[str, Any]:
        """Возвращает тайминг текущего проекта в виде словаря."""
        _sync_events_from_timeline()
        return {
            "start_date": date_start.date().toString("yyyy-MM-dd"),
            "days": spin_days.value(),
            "step_minutes": int(combo_step.currentText()),
            "columns": page.timing_column_names,
            "events": page.timing_events,
        }
    page.timing_get_json = timing_get_json  # type: ignore

    def timing_export_image(path: str, width_px: int = 1400) -> None:
        """Сохраняет изображение таймлайна в указанном файле."""
        try:
            img = timeline.render_to_image(width_px)
            img.save(path)
            if hasattr(page, "_log"):
                page._log(f"Изображение таймлайна сохранено в {path}.")
        except Exception as ex:
            if hasattr(page, "_log"):
                page._log(f"Ошибка экспорта изображения таймлайна: {ex}", "error")
    page.timing_export_image = timing_export_image  # type: ignore

    def timing_reload_for_current_project() -> None:
        """Принудительно перезагружает тайминг из БД для текущего проекта."""
        _load()
    page.timing_reload_for_current_project = timing_reload_for_current_project  # type: ignore

    # Совместимость: функция загрузки для ProjectPage (загружается автоматически при смене проекта).
    page.load_timing_data = _load  # type: ignore

    # 4.15 Начальная загрузка тайминга
    _load()