"""
Назначение: Утилиты (конфиг, папки, логирование).
"""
# 1. Импорт
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
try:
    import tomllib  # 3.11+
except Exception:
    import tomli as tomllib  # 3.10

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config.toml"
DATA_DIR = ROOT / "data"
LOGS_DIR = ROOT / "logs"
ASSETS_DIR = ROOT / "assets"

# 2. Папки
def ensure_folders():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
# 3. Конфиг
def load_config() -> dict:
    with open(CONFIG_PATH, "rb") as f:
        return tomllib.load(f)

# 4. Логи
def init_logging(log_path: str):
    handler = RotatingFileHandler(log_path, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
    handler.setFormatter(fmt)
    lg = logging.getLogger()
    lg.setLevel(logging.INFO)
    lg.addHandler(handler)
    lg.info("Логирование инициализировано")
