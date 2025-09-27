"""
Модуль ``timeline_canvas`` реализует базовый графический таймлайн без
пересечений на базе ``QGraphicsView``.  Блоки хранятся в абсолютных
минутах от даты начала проекта, а визуальная сетка служит для
ориентирования.  При добавлении и перемещении блоков происходит
автоматический сдвиг (рефлоу), чтобы избежать наложений.

Основные компоненты:

* ``EventModel`` — простая структура данных (``dataclass``) с
  атрибутами ``start_min``, ``duration_min``, ``col``, ``title`` и
  ``color``.  Время хранится в минутах от начала проекта.
* ``EventItem`` — графический элемент, наследующий
  ``QGraphicsRectItem``.  Позволяет перетаскивать блоки по вертикали и
  изменять их длительность перетягиванием нижней границы.
* ``TimelineScene`` — сцена, на которую помещаются все блоки.  Содержит
  методы ``place_and_reflow`` для автосдвига блока вперёд, ``export_events``
  для выгрузки данных, ``import_events`` для загрузки и ``reflow_column``
  для коррекции наложений в колонке.
* ``TimelineView`` — обёртка над сценой.  Предоставляет методы
  ``set_columns``, ``set_step``, ``add_event``, ``load_events``,
  ``export_events``, ``clear_events``, ``reflow_all`` и ``render_to_image``.

Код разделён на пронумерованные секции и снабжён краткими комментариями
для удобства понимания.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional
from PySide6 import QtWidgets, QtCore, QtGui

# Дополнительная высота заголовка (в пикселях) для отображения названия колонок.
HEADER_HEIGHT: int = 24

# Отступ слева для часовой шкалы (в пикселях).  На этой области
# отображаются подписи времени (например, 00:00, 01:00, …).  Все
# элементы (колонки и события) рисуются справа от этого отступа.
LEFT_MARGIN: int = 60


# 1. Константы для оформления
GRID_BG = QtGui.QColor("#1f232a")
GRID_LINE = QtGui.QColor("#3a3f46")
TEXT_COLOR = QtGui.QColor("#e9edf5")
BORDER_COLOR = QtGui.QColor("#0f1216")


# 2. Структура данных события
@dataclass
class EventModel:
    col: int
    start_min: int
    duration_min: int
    title: str
    color: str = "#6fbf73"

    @property
    def end_min(self) -> int:
        return self.start_min + max(0, self.duration_min)


# 3. Элемент события
class EventItem(QtWidgets.QGraphicsRectItem):
    """Графический элемент события.  Поддерживает перетаскивание и растяжение."""

    def __init__(self, model: EventModel, minute_px: float, col_width: int, left_margin: int = LEFT_MARGIN) -> None:
        super().__init__()
        self.setFlags(
            QtWidgets.QGraphicsItem.ItemIsMovable
            | QtWidgets.QGraphicsItem.ItemIsSelectable
        )
        self.setAcceptHoverEvents(True)
        self.model = model
        self.minute_px = minute_px
        self.col_width = col_width
        # Левый отступ для позиционирования (см. LEFT_MARGIN)
        self.left_margin = left_margin
        self._drag_resize = False
        self._drag_origin_y = 0.0
        self._orig_start = model.start_min
        self._orig_dur = model.duration_min
        # Подпись внутри блока
        # Используем QGraphicsTextItem для переноса строк
        self._label = QtWidgets.QGraphicsTextItem(self)
        self._label.setDefaultTextColor(TEXT_COLOR)
        self._label.setZValue(1)
        self._update_rect()

    def _update_rect(self) -> None:
        """Обновляет положение и текст блока."""
        # Координаты X/Y с учётом заголовка (HEADER_HEIGHT)
        # по горизонтали учитываем левый отступ
        x = self.left_margin + self.model.col * self.col_width
        # по вертикали смещаем на высоту заголовка
        y = HEADER_HEIGHT + self.model.start_min * self.minute_px
        # Содержимое подписи: время начала–конца и название
        t1 = f"{(self.model.start_min // 60) % 24:02d}:{self.model.start_min % 60:02d}"
        t2 = f"{(self.model.end_min // 60) % 24:02d}:{self.model.end_min % 60:02d}"
        text = f"{t1}–{t2}"
        if self.model.title.strip():
            text += f"\n{self.model.title.strip()}"
        # Устанавливаем ширину текста для переноса строк (на несколько пикселей меньше ширины колонки)
        self._label.setTextWidth(self.col_width - 12)
        # Обновляем текст подписи
        self._label.setHtml("<div align='left'>" + text.replace("\n", "<br>") + "</div>")
        # Подсчитаем высоту текста, чтобы блок не был меньше содержимого
        text_height = self._label.boundingRect().height() + 8  # небольшой отступ
        # Высота блока в пикселях на основе длительности (масштаб minute_px)
        dur_height = max(1, int(self.model.duration_min * self.minute_px))
        h = max(int(dur_height), int(text_height))
        # Устанавливаем прямоугольник с небольшими отступами
        self.setRect(QtCore.QRectF(x + 1, y + 1, self.col_width - 2, h - 2))
        # Позиция подписи внутри блока
        self._label.setPos(self.rect().x() + 6, self.rect().y() + 4)
        # Цвет заливки и рамки
        self.setBrush(QtGui.QBrush(QtGui.QColor(self.model.color)))
        self.setPen(QtGui.QPen(BORDER_COLOR, 1))

    # 3.1 Наведение мыши — определяем область растяжения
    def hoverMoveEvent(self, ev: QtWidgets.QGraphicsSceneHoverEvent) -> None:
        if abs(ev.pos().y() - self.rect().bottom()) < 6:
            self.setCursor(QtCore.Qt.SizeVerCursor)
            self._drag_resize = True
        else:
            self.setCursor(QtCore.Qt.OpenHandCursor)
            self._drag_resize = False
        super().hoverMoveEvent(ev)

    # 3.2 Нажатие мыши — запоминаем исходное состояние
    def mousePressEvent(self, ev: QtWidgets.QGraphicsSceneMouseEvent) -> None:
        self._drag_origin_y = ev.scenePos().y()
        self._orig_start = self.model.start_min
        self._orig_dur = self.model.duration_min
        # Запоминаем исходную колонку для последующей корректировки других колонок
        self._orig_col = self.model.col
        self.setCursor(QtCore.Qt.ClosedHandCursor)
        super().mousePressEvent(ev)

    # 3.3 Перемещение мыши — изменяем начало или длительность
    def mouseMoveEvent(self, ev: QtWidgets.QGraphicsSceneMouseEvent) -> None:
        dy = ev.scenePos().y() - self._drag_origin_y
        dmin = int(round(dy / self.minute_px))
        if self._drag_resize:
            self.model.duration_min = max(1, self._orig_dur + dmin)
        else:
            self.model.start_min = max(0, self._orig_start + dmin)
        self._update_rect()
        super().mouseMoveEvent(ev)

    # 3.4 Отпускание мыши — завершаем перетаскивание
    def mouseReleaseEvent(self, ev: QtWidgets.QGraphicsSceneMouseEvent) -> None:
        self.setCursor(QtCore.Qt.OpenHandCursor)
        super().mouseReleaseEvent(ev)
        # После завершения перетаскивания обновим модель по горизонтали и вертикали.
        # Вычисляем новую колонку на основе фактической позиции X и лев.
        try:
            scene_rect = self.sceneBoundingRect()
            # X координата левого края блока
            x = scene_rect.x()
            # Учтём левый отступ: определяем новую колонку
            new_col = int((x - self.left_margin) // self.col_width) if self.col_width > 0 else self.model.col
            if new_col < 0:
                new_col = 0
            # Y координата верхней границы блока
            y = scene_rect.y() - HEADER_HEIGHT
            # Новое время начала (в минутах). Приводим к целому, но не ограничиваем до шага,
            # так как время в модели хранится в минутах. Отрицательные значения ставим в ноль.
            new_start = int(round(y / self.minute_px)) if self.minute_px > 0 else self.model.start_min
            if new_start < 0:
                new_start = 0
            changed = False
            old_col = self.model.col
            # Обновляем колонку, если изменилась
            if new_col != self.model.col:
                self.model.col = new_col
                changed = True
            # Обновляем время начала
            if new_start != self.model.start_min:
                self.model.start_min = new_start
                changed = True
            # Если что‑то изменилось, просим сцену разместить элемент заново без пересечений и сжатием
            if changed:
                # Обновляем прямоугольник под новые координаты
                self._update_rect()
                scene = self.scene()
                if hasattr(scene, "place_and_reflow"):
                    scene.place_and_reflow(self)
                # После перемещения/растяжения убираем зазоры как в исходной колонке, так и в новой
                try:
                    cols_to_fix = {old_col, self.model.col}
                    if hasattr(scene, "compress_columns"):
                        scene.compress_columns(cols_to_fix)
                except Exception:
                    pass
        except Exception:
            pass


# 4. Сцена таймлайна
class TimelineScene(QtWidgets.QGraphicsScene):
    """Сцена для событий.  Содержит алгоритм без пересечений."""

    updated = QtCore.Signal()

    def __init__(self, minute_px: float = 0.6, col_width: int = 240, step_min: int = 15, days: int = 3, left_margin: int = LEFT_MARGIN) -> None:
        super().__init__()
        self.minute_px = minute_px
        self.col_width = col_width
        self.step_min = step_min
        # Левый отступ, на котором будут отображаться часовые отметки
        self.left_margin = left_margin
        self.columns: List[str] = ["Общий"]
        self.days = max(1, days)
        self.items_map: List[EventItem] = []
        self.setBackgroundBrush(QtGui.QBrush(GRID_BG))

    # 4.1 Рисование сетки
    def drawBackground(self, painter: QtGui.QPainter, rect: QtCore.QRectF) -> None:
        """Рисует фон: заголовки колонок, вертикальные линии и горизонтальную сетку."""
        # Заливка фона
        painter.fillRect(rect, GRID_BG)
        # Ширина всей сцены
        total_width = self.left_margin + len(self.columns) * self.col_width
        header_bg = QtGui.QColor("#2f343d")
        # Рисуем заголовки колонок единожды у верхней границы сцены
        painter.setPen(QtGui.QPen(TEXT_COLOR))
        # Левый заголовок (пустой) для оси времени
        painter.fillRect(QtCore.QRectF(0, 0, self.left_margin, HEADER_HEIGHT), header_bg)
        # Заголовки для каждой колонки
        for i, name in enumerate(self.columns):
            x = self.left_margin + i * self.col_width
            painter.fillRect(QtCore.QRectF(x, 0, self.col_width, HEADER_HEIGHT), header_bg)
            painter.drawText(QtCore.QRectF(x, 0, self.col_width, HEADER_HEIGHT), QtCore.Qt.AlignCenter, name)
        # Вертикальные линии
        painter.setPen(QtGui.QPen(GRID_LINE, 1))
        # Линия после левого отступа
        painter.drawLine(self.left_margin, rect.top(), self.left_margin, rect.bottom())
        for i in range(len(self.columns) + 1):
            x = self.left_margin + i * self.col_width
            painter.drawLine(x, rect.top(), x, rect.bottom())
        # Горизонтальная сетка по шагу (без заголовка)
        top_m = max(0, int((rect.top() - HEADER_HEIGHT) / self.minute_px) - 1)
        bot_m = int((rect.bottom() - HEADER_HEIGHT) / self.minute_px) + 1
        s = max(1, self.step_min)
        for m in range((top_m // s) * s, bot_m, s):
            y = HEADER_HEIGHT + m * self.minute_px
            painter.drawLine(0, y, total_width, y)
            # Подписываем каждый час на левой оси
            if m % 60 == 0:
                h = (m // 60) % 24
                painter.setPen(QtGui.QPen(TEXT_COLOR))
                painter.drawText(QtCore.QRectF(0, y - self.minute_px * s, self.left_margin - 4, self.minute_px * s), QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, f"{h:02d}:00")
                painter.setPen(QtGui.QPen(GRID_LINE, 1))

    # 4.2 Установка колонок и параметров
    def set_columns(self, names: List[str]) -> None:
        self.columns = names[:]
        self.update()

    def set_params(self, step_min: int) -> None:
        self.step_min = step_min
        self.update()

    # 4.3 Работа с событиями
    def add_event(self, model: EventModel) -> EventItem:
        # При создании элемента передаём левый отступ, чтобы событие знало, где начинать
        item = EventItem(model, self.minute_px, self.col_width, self.left_margin)
        self.addItem(item)
        self.items_map.append(item)
        return item

    def _events_in_col(self, col: int) -> List[EventItem]:
        return sorted([it for it in self.items_map if it.model.col == col], key=lambda it: it.model.start_min)

    def _collide(self, a: EventItem, b: EventItem) -> bool:
        if a.model.col != b.model.col:
            return False
        return not (a.model.end_min <= b.model.start_min or b.model.end_min <= a.model.start_min)

    def place_and_reflow(self, item: EventItem) -> None:
        """Вставляет событие и сдвигает его вперёд при коллизии."""
        arr = self._events_in_col(item.model.col)
        arr = [x for x in arr if x is not item]
        idx = 0
        while idx < len(arr) and arr[idx].model.start_min <= item.model.start_min:
            idx += 1
        arr.insert(idx, item)
        # Сдвигаем текущий и последующие блоки вперёд, если есть пересечения.
        for i in range(1, len(arr)):
            prev, cur = arr[i - 1], arr[i]
            if self._collide(prev, cur):
                # Если блоки пересекаются, начинаем текущий сразу после предыдущего
                cur.model.start_min = prev.model.end_min
                cur._update_rect()
        # После устранения пересечений избавляемся от зазоров между событиями.
        for i in range(1, len(arr)):
            prev, cur = arr[i - 1], arr[i]
            expected_start = prev.model.end_min
            if cur.model.start_min != expected_start:
                cur.model.start_min = expected_start
                cur._update_rect()
        self.updated.emit()

    def compress_column(self, col: int) -> None:
        """Сдвигает все события в колонке вплотную друг к другу без пропусков.

        Элементы сортируются по времени начала; далее каждый следующий
        элемент располагается сразу после предыдущего, изменяя свой
        ``start_min``. Это позволяет убрать «пустые» промежутки после
        перемещения блоков между колонками.
        """
        arr = self._events_in_col(col)
        for i in range(1, len(arr)):
            prev, cur = arr[i - 1], arr[i]
            expected_start = prev.model.end_min
            if cur.model.start_min != expected_start:
                cur.model.start_min = expected_start
                cur._update_rect()
        # Оповещаем наблюдателей об изменениях
        self.updated.emit()

    def compress_columns(self, cols: list[int] | set[int]) -> None:
        """Применяет ``compress_column`` для нескольких колонок.

        Используйте этот метод, когда события перетаскиваются между
        колонками, чтобы убрать пустые промежутки и прилипать блоки друг к
        другу. Принимает набор индексов колонок.
        """
        for col in cols:
            self.compress_column(col)

    def export_events(self) -> List[Dict[str, Any]]:
        return [asdict(it.model) for it in self.items_map]

    def import_events(self, data: List[Dict[str, Any]]) -> None:
        for it in self.items_map:
            self.removeItem(it)
        self.items_map.clear()
        for d in data:
            self.add_event(EventModel(**d))
        # Ре-флоу по каждой колонке
        for c in range(len(self.columns)):
            arr = self._events_in_col(c)
            for i in range(1, len(arr)):
                if self._collide(arr[i - 1], arr[i]):
                    arr[i].model.start_min = arr[i - 1].model.end_min
                    arr[i]._update_rect()
        self.updated.emit()

    # 4.4 Поиск элемента по сцене
    def event_at(self, pos: QtCore.QPointF) -> EventItem | None:
        """Возвращает элемент события в указанной позиции сцены или None."""
        item = self.itemAt(pos, QtGui.QTransform())
        if isinstance(item, EventItem):
            return item
        # Если клик попал на подпись или рамку, ищем родителя
        if isinstance(item, QtWidgets.QGraphicsSimpleTextItem) and isinstance(item.parentItem(), EventItem):
            return item.parentItem()
        return None


# 5. Виджет таймлайна
class TimelineView(QtWidgets.QGraphicsView):
    """Обёртка над сценой. Предоставляет API для страницы тайминга."""

    changed = QtCore.Signal()

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setRenderHints(
            QtGui.QPainter.Antialiasing | QtGui.QPainter.TextAntialiasing
        )
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self._scene = TimelineScene()
        self.setScene(self._scene)
        self._scene.updated.connect(lambda: self.changed.emit())

    # 5.1 Колонки и параметры
    def set_columns(self, names: List[str]) -> None:
        self._scene.set_columns(names)
        self._fit_scene()

    def set_step(self, step_min: int) -> None:
        self._scene.set_params(step_min)
        self._fit_scene()

    def set_days(self, days: int) -> None:
        """Устанавливает количество дней, отображаемых по вертикали.

        Количество дней влияет на высоту сцены: каждые 24 часа
        помещаются один под другим. После обновления сцена
        пересчитывается.
        """
        self._scene.days = max(1, int(days))
        self._fit_scene()

    def _fit_scene(self) -> None:
        """Подгоняет размеры сцены под количество дней и колонок."""
        # Высота сцены учитывает высоту заголовка
        h = int(self._scene.days * 24 * 60 * self._scene.minute_px) + HEADER_HEIGHT + 200
        # Ширина сцены учитывает левый отступ и колонки
        w = int(self._scene.left_margin + len(self._scene.columns) * self._scene.col_width) + 200
        self._scene.setSceneRect(0, 0, w, h)

    # 5.2 Работа с событиями
    def clear_events(self) -> None:
        self._scene.import_events([])
        self.changed.emit()

    def load_events(self, data: List[Dict[str, Any]]) -> None:
        self._scene.import_events(data)
        self.changed.emit()

    def export_events(self) -> List[Dict[str, Any]]:
        return self._scene.export_events()

    def add_event(self, col: int, start_min: int, duration_min: int, title: str, color: str = "#6fbf73") -> None:
        item = self._scene.add_event(EventModel(col, start_min, duration_min, title, color))
        self._scene.place_and_reflow(item)
        self.changed.emit()

    def reflow_all(self) -> None:
        for c in range(len(self._scene.columns)):
            arr = self._scene._events_in_col(c)
            for i in range(1, len(arr)):
                if self._scene._collide(arr[i - 1], arr[i]):
                    arr[i].model.start_min = arr[i - 1].model.end_min
                    arr[i]._update_rect()
        self.changed.emit()

    # 5.3 Рендер в изображение для PDF
    def render_to_image(self, width: int) -> QtGui.QImage:
        rect = self._scene.sceneRect()
        scale = width / max(1, rect.width())
        img = QtGui.QImage(int(rect.width() * scale), int(rect.height() * scale), QtGui.QImage.Format_ARGB32)
        img.fill(QtCore.Qt.white)
        painter = QtGui.QPainter(img)
        painter.scale(scale, scale)
        self._scene.render(painter)
        painter.end()
        return img

    # 5.4 Получить событие по координатам виджета
    def get_event_at(self, view_pos: QtCore.QPoint) -> EventItem | None:
        """Возвращает элемент события под позицией в координатах виджета.

        Вспомогательная функция для контекстного меню. Преобразует позицию
        виджета в координаты сцены и запрашивает объект у сцены.
        """
        if self._scene is None:
            return None
        scene_pos = self.mapToScene(view_pos)
        return self._scene.event_at(scene_pos)