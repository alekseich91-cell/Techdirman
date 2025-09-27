"""
Назначение: Точка входа в приложение.
Как работает:
- Готовит логи и БД.
- Запускает главное окно (PySide6).
"""
# 1. Импорт библиотек
import sys
from pathlib import Path
from PySide6 import QtWidgets

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from utils import init_logging, ensure_folders, load_config  # noqa: E402
from db import DB  # noqa: E402
from ui import MainWindow  # noqa: E402

# 2. Главная функция
def main():
    config = load_config()
    ensure_folders()
    init_logging(config["app"]["log_path"])
    db = DB(Path(config["app"]["db_path"]).resolve())
    db.init_schema()
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow(db=db)
    w.show()
    sys.exit(app.exec())

# 3. Точка входа
if __name__ == "__main__":
    main()
