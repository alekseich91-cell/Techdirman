"""
Назначение:
    Пакет графического интерфейса (Qt Widgets) приложения TechDirRentMan.

Принцип работы:
    - Декомпозирован из прежнего monolith ui.py без потери строк кода.
    - Экспортирует ключевые классы для совместимости: MainWindow, DatabaseWindow, ProjectPage,
      а также некоторые виджеты и делегаты.

Стиль:
    - Нумерованные секции и краткие комментарии.
"""

# 1. Экспорт основных классов для совместимости
from .main_window import MainWindow  # noqa: F401
from .db_window import DatabaseWindow  # noqa: F401
from .project_page import ProjectPage  # noqa: F401

# 2. Экспорт вспомогательных классов/констант при необходимости
from .dialogs import MoveDialog, PowerMismatchDialog  # noqa: F401
from .widgets import ImageDropLabel, LogDock, SmartDoubleSpinBox, FileDropLabel  # noqa: F401
from .delegates import ClassRuDelegate, WrapTextDelegate  # noqa: F401
from .common import CLASS_RU2EN, CLASS_EN2RU, WRAP_THRESHOLD, fmt_num, fmt_sign, to_float  # noqa: F401
