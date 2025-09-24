# -*- coding: utf-8 -*-
"""
Назначение:
    Реализация вкладки «Бухгалтерия» для приложения TechDirRentMan. Эта вкладка
    отвечает за расчёт финансовых показателей проекта по двум направлениям:

    1. **Общая**: ввод коэффициентов и скидок/комиссий/налога по каждому
       подрядчику (только класс «оборудование»), расчёт итогов по
       клиентскому потоку. Порядок вычислений: базовая сумма → вычитаем
       скидку → вычитаем комиссию → прибавляем налог. Глобальный
       коэффициент подрядчика заменяет индивидуальные coeff для всех
       позиций класса «оборудование» данного подрядчика при сохранении.
       Пользователь может временно отключить применение глобального
       коэффициента для подрядчика через флажок в таблице, чтобы
       сравнить результаты с коэффициентом 1.0. При сохранении
       используется исходное значение коэффициента, а не состояние
       предпросмотра.

    2. **Внутренняя**: учёт внутренних скидок (наша скидка), доходов
       из сметы (привязанных к конкретным подрядчикам) и ручных расходов.
       Итоговая прибыль = внутренние скидки + доходы из сметы – расходы.
       Сумма, которую должны выплатить подрядчику, равна итогу по клиентскому
       потоку (с налогом) минус внутренняя скидка и минус доходы, привязанные
       к этому подрядчику. Начиная с версии A62L, комиссия подрядчика
       уменьшает нашу внутреннюю скидку так же, как клиентская скидка:
       из нашей скидки вычитаются как скидка, предоставленная клиенту,
       так и комиссия, уплаченная по итогам клиентского потока.

    3. **Оплаты**: фиксация выплат подрядчикам. В этой вкладке
       отображается список подрядчиков с рассчитанной суммой задолженности
       (учитывая скидки, комиссию, налог, внутренние скидки и доходы из
       сметы) и вводится сумма, уже выплаченная каждому подрядчику. После
       ввода оплата вычитается из задолженности, показывая остаток к
       выплате. Значения выплат сохраняются при нажатии кнопки «Сохранить».

    Начиная с версии A11-fixed, в таблице «Наша скидка по подрядчикам» добавлен
    столбец «Сумма скидки ₽», который показывает абсолютное значение нашей
    скидки (для справки).

Принцип работы:
    • Вкладка состоит из двух внутренних вкладок: «Общая» и «Внутренняя».
    • Все изменения пользователь делает в режиме предпросмотра: значения
      коэффициентов, скидок, комиссий и налогов хранятся во временных
      структурах. Таблицы обновляются в режиме реального времени.
    • Нажатие на красную кнопку «Сохранить изменения» переносит значения
      предпросмотра в реальные настройки, сохраняет данные через провайдер
      (в JSON или БД) и применяет глобальные коэффициенты к позициям.
    • Все важные действия и возникающие ошибки журналируются в файл
      ``logs/finance_tab.log``.

Стиль кода:
    • Файл разбит на пронумерованные секции с короткими заголовками.
    • Внутри секций используются краткие комментарии для объяснения логики.
    • Для расчётов и обновлений используются отдельные функции (например,
      ``aggregate_by_vendor`` или ``compute_client_flow``) для упрощения
      тестирования и повторного использования.
"""

# 1. Импорт библиотек и настройка логирования
from __future__ import annotations
import json
import logging
import os
import traceback
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Any

from PySide6 import QtCore, QtGui, QtWidgets
# Импортируем SmartDoubleSpinBox для более удобного ввода чисел без стрелок
from .widgets import SmartDoubleSpinBox  # type: ignore
# Импортируем отображение классов (EN→RU)
from .common import CLASS_EN2RU

# Настройка логов в файл
LOG_DIR = os.path.join(os.path.dirname(__file__), "..", "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "finance_tab.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("finance_tab")


# 2. Простые структуры данных
@dataclass
class Item:
    """Элемент сметы (строка проекта)."""
    id: str
    vendor: str
    cls: str              # класс: 'equipment' и т.д.
    department: str
    zone: str
    name: str
    price: float
    qty: float
    coeff: float = 1.0
    # Служебное поле для восстановления при отмене глоб. коэффициента
    original_coeff: Optional[float] = None

    def amount(self, effective_coeff: Optional[float] = None) -> float:
        # Считаем сумму по строке с учётом эффективного коэффициента
        c = effective_coeff if effective_coeff is not None else self.coeff
        return float(self.price) * float(self.qty) * float(c)


@dataclass
class VendorSettings:
    """Настройки подрядчика для вкладки «Общая», «Внутренняя» и «Оплаты».

    Помимо стандартных параметров (коэффициент, скидки, комиссии и налог),
    в класс добавлено поле ``paid``, которое отражает сумму, уже
    выплаченную подрядчику. Это поле используется на новой подвкладке
    «Оплаты» для вычисления остатка задолженности.
    """
    coeff: float = 1.0                         # Глобальный коэффициент (только для equipment)
    discount_pct: float = 0.0                  # Клиентская скидка % (только на equipment)
    commission_pct: float = 0.0                # Комиссия % (только на equipment)
    tax_pct: float = 0.0                       # Налог % (на сумму после скидки и комиссии)
    # Внутренняя (наша) скидка: либо % либо сумма. Одновременно не используются.
    our_discount_pct: Optional[float] = None   # Если задан %, приоритет над суммой
    our_discount_sum: Optional[float] = None   # Если задана сумма и % не задан
    # Сохраняет, был ли включён глобальный коэффициент для подрядчика в предыдущем сеансе.
    coeff_enabled: bool = False
    # Сумма, уже выплаченная подрядчику. Используется в подвкладке «Оплаты».
    paid: float = 0.0


@dataclass
class ProfitItem:
    """Позиция дохода из сметы, привязанная к подрядчику."""
    vendor: str
    description: str
    amount: float


@dataclass
class ExpenseItem:
    """Ручной расход."""
    name: str
    qty: float
    price: float

    def total(self) -> float:
        return float(self.qty) * float(self.price)


# 3. Провайдер данных по умолчанию (файловый JSON)
class FileDataProvider:
    """Файловое хранение состояний вкладки «Бухгалтерия» + загрузка позиций проекта.

    Ожидается структура проекта:
        project_root/
            data/
                project_<id>_items.json               — список позиций сметы
                project_<id>_finance.json             — настройки «Бухгалтерии»
    В реальном приложении вы можете заменить провайдер на БД-провайдер с тем же интерфейсом.
    """

    def __init__(self, project_root: str, project_id: str) -> None:
        self.project_root = project_root
        self.project_id = project_id
        self.data_dir = os.path.join(project_root, "data")
        os.makedirs(self.data_dir, exist_ok=True)
        self.items_path = os.path.join(self.data_dir, f"project_{project_id}_items.json")
        self.finance_path = os.path.join(self.data_dir, f"project_{project_id}_finance.json")

    def load_items(self) -> List[Item]:
        # Загружаем позиции проекта
        if not os.path.exists(self.items_path):
            logger.info("Файл позиций не найден, возвращаю пустой список: %s", self.items_path)
            return []
        with open(self.items_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        items: List[Item] = []
        for r in raw:
            # Получаем коэффициент элемента. Если original_coeff отсутствует в сохранённом файле,
            # считаем его исходным значением coeff (то есть коэффициентом из импортированной сметы).
            coeff_val = float(r.get("coeff", 1.0))
            orig = r.get("original_coeff", None)
            if orig is None:
                orig = coeff_val
            items.append(Item(
                id=str(r.get("id", "")),
                vendor=str(r.get("vendor", "")),
                cls=str(r.get("class", r.get("cls", ""))),
                department=str(r.get("department", "")),
                zone=str(r.get("zone", "")),
                name=str(r.get("name", "")),
                price=float(r.get("price", 0.0)),
                qty=float(r.get("qty", 0.0)),
                coeff=coeff_val,
                original_coeff=orig,
            ))
        return items

    def save_items(self, items: List[Item]) -> None:
        # Сохраняем позиции проекта
        raw = []
        for it in items:
            raw.append({
                "id": it.id,
                "vendor": it.vendor,
                "class": it.cls,
                "department": it.department,
                "zone": it.zone,
                "name": it.name,
                "price": it.price,
                "qty": it.qty,
                "coeff": it.coeff,
                "original_coeff": it.original_coeff,
            })
        with open(self.items_path, "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)
        logger.info("Позиции проекта сохранены: %s", self.items_path)

    def load_finance(self) -> Tuple[Dict[str, VendorSettings], List[ProfitItem], List[ExpenseItem]]:
        # Загружаем настройки «Бухгалтерии»
        if not os.path.exists(self.finance_path):
            logger.info("Файл finance.json не найден, возвращаю пустые настройки: %s", self.finance_path)
            return {}, [], []
        with open(self.finance_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        vendors: Dict[str, VendorSettings] = {}
        for name, v in raw.get("vendors", {}).items():
            vendors[name] = VendorSettings(
                coeff=float(v.get("coeff", 1.0)),
                discount_pct=float(v.get("discount_pct", 0.0)),
                commission_pct=float(v.get("commission_pct", 0.0)),
                tax_pct=float(v.get("tax_pct", 0.0)),
                our_discount_pct=(None if v.get("our_discount_pct") is None else float(v.get("our_discount_pct"))),
                our_discount_sum=(None if v.get("our_discount_sum") is None else float(v.get("our_discount_sum"))),
                # Загрузка состояния включения глобального коэффициента (по умолчанию False)
                coeff_enabled=bool(v.get("coeff_enabled", False)),
                # Загрузка выплаченной суммы (по умолчанию 0)
                paid=float(v.get("paid", 0.0)),
            )
        profits: List[ProfitItem] = []
        for p in raw.get("profit_items", []):
            profits.append(ProfitItem(
                vendor=str(p.get("vendor", "")),
                description=str(p.get("description", "")),
                amount=float(p.get("amount", 0.0)),
            ))
        expenses: List[ExpenseItem] = []
        for e in raw.get("expenses", []):
            expenses.append(ExpenseItem(
                name=str(e.get("name", "")),
                qty=float(e.get("qty", 0.0)),
                price=float(e.get("price", 0.0)),
            ))
        return vendors, profits, expenses

    def save_finance(self, vendors: Dict[str, VendorSettings], profits: List[ProfitItem], expenses: List[ExpenseItem]) -> None:
        # Сохраняем настройки «Бухгалтерии»
        raw_v: Dict[str, Dict[str, Any]] = {}
        for name, v in vendors.items():
            # Сохраняем все параметры, включая состояние включения глобального коэффициента.
            raw_v[name] = {
                "coeff": v.coeff,
                "discount_pct": v.discount_pct,
                "commission_pct": v.commission_pct,
                "tax_pct": v.tax_pct,
                "our_discount_pct": v.our_discount_pct,
                "our_discount_sum": v.our_discount_sum,
                "coeff_enabled": v.coeff_enabled,
                # Сохраняем сумму, уже выплаченную подрядчику
                "paid": v.paid,
            }
        raw_p = [{"vendor": p.vendor, "description": p.description, "amount": p.amount} for p in profits]
        raw_e = [{"name": e.name, "qty": e.qty, "price": e.price} for e in expenses]
        with open(self.finance_path, "w", encoding="utf-8") as f:
            json.dump({"vendors": raw_v, "profit_items": raw_p, "expenses": raw_e}, f, ensure_ascii=False, indent=2)
        logger.info("Настройки «Бухгалтерии» сохранены: %s", self.finance_path)


# 3a. Провайдер данных, использующий базу данных проекта
class DBDataProvider:
    """Провайдер данных, работающий напрямую с классом DB из ProjectPage.

    Этот провайдер загружает позиции проекта из таблицы items, а финансовые
    настройки из поля finance_json таблицы projects. При сохранении он
    обновляет поля coeff и amount для позиций в БД и записывает
    конфигурацию вкладки «Бухгалтерия» обратно в projects.finance_json.

    Для работы требуется объект страницы ProjectPage с атрибутами db и
    project_id. В случае отсутствия этих атрибутов провайдер
    возвращает пустые данные.
    """

    def __init__(self, page: object) -> None:
        self.page = page

    def load_items(self) -> List[Item]:
        """
        Загружает позиции проекта из базы данных.

        В таблице items поля называются id, vendor, type, department,
        zone, name, unit_price, qty, coeff. Объект sqlite3.Row не
        поддерживает метод get(), поэтому доступ к колонкам производится
        по индексам/именам через оператор [] с обработкой отсутствующих
        значений. При любой ошибке возвращается пустой список.
        """
        items: List[Item] = []
        try:
            proj_id = getattr(self.page, "project_id", None)
            db = getattr(self.page, "db", None)
            if proj_id is None or db is None:
                logger.info("DBDataProvider: нет project_id или db, возвращаю пустой список позиций")
                return []
            # Получаем строки из базы. Каждая строка представляет собой
            # sqlite3.Row, который можно индексировать по имени колонки.
            rows = db.list_items(proj_id)
            for row in rows:
                try:
                    # sqlite3.Row поддерживает доступ по ключу, но не метод get().
                    # При отсутствии поля возвращаем разумные значения.
                    item_id = str(row["id"]) if "id" in row.keys() else ""
                    vendor = ""
                    if "vendor" in row.keys() and row["vendor"] is not None:
                        vendor = str(row["vendor"]).strip()
                    vendor = vendor or "(без подрядчика)"
                    cls = "equipment"
                    if "type" in row.keys() and row["type"]:
                        cls = row["type"]
                    department = ""
                    if "department" in row.keys() and row["department"]:
                        department = row["department"]
                    zone = ""
                    if "zone" in row.keys() and row["zone"]:
                        zone = row["zone"]
                    name = ""
                    if "name" in row.keys() and row["name"]:
                        name = row["name"]
                    unit_price = 0.0
                    if "unit_price" in row.keys() and row["unit_price"] is not None:
                        unit_price = float(row["unit_price"])
                    qty = 0.0
                    if "qty" in row.keys() and row["qty"] is not None:
                        qty = float(row["qty"])
                    coeff = 1.0
                    if "coeff" in row.keys() and row["coeff"] is not None:
                        coeff = float(row["coeff"])
                    # original_coeff сохраняет исходный коэффициент позиции (из БД)
                    orig_coeff = coeff
                    items.append(Item(
                        id=item_id,
                        vendor=vendor,
                        cls=cls,
                        department=department,
                        zone=zone,
                        name=name,
                        price=unit_price,
                        qty=qty,
                        coeff=coeff,
                        original_coeff=orig_coeff,
                    ))
                except Exception:
                    # Логируем ошибку для конкретной строки и продолжаем
                    logger.error("DBDataProvider.load_items: ошибка обработки строки: %s", traceback.format_exc())
                    continue
            return items
        except Exception:
            logger.error("DBDataProvider.load_items: %s", traceback.format_exc())
            return []

    def save_items(self, items: List[Item]) -> None:
        """Сохраняет изменённые коэффициенты позиций обратно в базу данных.

        Для каждой позиции класса equipment обновляется поле coeff и
        вычисляется новое значение amount = unit_price * qty * coeff.
        """
        try:
            proj_id = getattr(self.page, "project_id", None)
            db = getattr(self.page, "db", None)
            if proj_id is None or db is None:
                logger.info("DBDataProvider.save_items: нет project_id или db, ничего не сохраняю")
                return
            for it in items:
                try:
                    item_id = int(it.id)
                    # Получаем текущие данные строки для расчёта суммы
                    row = db.get_item_by_id(item_id)
                    if row is None:
                        continue
                    unit_price = float(row["unit_price"]) if row["unit_price"] is not None else float(it.price)
                    qty = float(row["qty"]) if row["qty"] is not None else float(it.qty)
                    new_amount = unit_price * qty * float(it.coeff)
                    # Обновляем поля coeff и amount
                    db.update_item_fields(item_id, {"coeff": float(it.coeff), "amount": new_amount})
                except Exception:
                    logger.error("DBDataProvider.save_items (item %s): %s", it.id, traceback.format_exc())
        except Exception:
            logger.error("DBDataProvider.save_items: %s", traceback.format_exc())

    def load_finance(self) -> Tuple[Dict[str, VendorSettings], List[ProfitItem], List[ExpenseItem]]:
        """Загружает финансовые настройки из поля projects.finance_json.

        Возвращает словарь настроек подрядчиков, список позиций доходов и
        список расходов. Если поле finance_json пустое, возвращает пустые
        структуры.
        """
        vendors: Dict[str, VendorSettings] = {}
        profits: List[ProfitItem] = []
        expenses: List[ExpenseItem] = []
        try:
            proj_id = getattr(self.page, "project_id", None)
            db = getattr(self.page, "db", None)
            if proj_id is None or db is None:
                return vendors, profits, expenses
            raw_json = db.get_project_finance(proj_id)
            if not raw_json:
                return vendors, profits, expenses
            raw = json.loads(raw_json)
            for name, v in raw.get("vendors", {}).items():
                vendors[name] = VendorSettings(
                    coeff=float(v.get("coeff", 1.0)),
                    discount_pct=float(v.get("discount_pct", 0.0)),
                    commission_pct=float(v.get("commission_pct", 0.0)),
                    tax_pct=float(v.get("tax_pct", 0.0)),
                    our_discount_pct=(None if v.get("our_discount_pct") is None else float(v.get("our_discount_pct"))),
                    our_discount_sum=(None if v.get("our_discount_sum") is None else float(v.get("our_discount_sum"))),
                    coeff_enabled=bool(v.get("coeff_enabled", False)),
                    paid=float(v.get("paid", 0.0)),
                )
            for p in raw.get("profit_items", []):
                profits.append(ProfitItem(
                    vendor=str(p.get("vendor", "")),
                    description=str(p.get("description", "")),
                    amount=float(p.get("amount", 0.0)),
                ))
            for e in raw.get("expenses", []):
                expenses.append(ExpenseItem(
                    name=str(e.get("name", "")),
                    qty=float(e.get("qty", 0.0)),
                    price=float(e.get("price", 0.0)),
                ))
        except Exception:
            logger.error("DBDataProvider.load_finance: %s", traceback.format_exc())
        return vendors, profits, expenses

    def save_finance(self, vendors: Dict[str, VendorSettings], profits: List[ProfitItem], expenses: List[ExpenseItem]) -> None:
        """Сохраняет финансовые настройки в поле projects.finance_json.

        Параметры аналогичны FileDataProvider.save_finance.
        """
        try:
            proj_id = getattr(self.page, "project_id", None)
            db = getattr(self.page, "db", None)
            if proj_id is None or db is None:
                logger.info("DBDataProvider.save_finance: нет project_id или db, ничего не сохраняю")
                return
            raw_v = {}
            for name, v in vendors.items():
                raw_v[name] = {
                    "coeff": v.coeff,
                    "discount_pct": v.discount_pct,
                    "commission_pct": v.commission_pct,
                    "tax_pct": v.tax_pct,
                    "our_discount_pct": v.our_discount_pct,
                    "our_discount_sum": v.our_discount_sum,
                    "coeff_enabled": v.coeff_enabled,
                    # Сохраняем сумму, уже выплаченную подрядчику
                    "paid": v.paid,
                }
            raw_p = [{"vendor": p.vendor, "description": p.description, "amount": p.amount} for p in profits]
            raw_e = [{"name": e.name, "qty": e.qty, "price": e.price} for e in expenses]
            json_str = json.dumps({"vendors": raw_v, "profit_items": raw_p, "expenses": raw_e}, ensure_ascii=False)
            db.set_project_finance(proj_id, json_str)
            logger.info("DBDataProvider.save_finance: настройки сохранены в projects.finance_json")
        except Exception:
            logger.error("DBDataProvider.save_finance: %s", traceback.format_exc())


# 4. Расчётные функции
def aggregate_by_vendor(items: List[Item], vendor_coeffs_preview: Dict[str, float]) -> Dict[str, Dict[str, float]]:
    """
    Агрегируем суммы по подрядчикам и классам с учётом предпросмотра
    коэффициентов. Если подрядчик отсутствует в ``vendor_coeffs_preview``
    (то есть глобальный коэффициент отключён в предпросмотре), для
    позиций класса ``equipment`` используется значение
    ``Item.original_coeff``, если оно задано; иначе применяется
    существующий коэффициент ``Item.coeff``. Это позволяет корректно
    игнорировать глобальный коэффициент в предпросмотре даже после
    сохранения, когда ``Item.coeff`` был заменён глобальным
    коэффициентом.

    :param items: список позиций сметы
    :param vendor_coeffs_preview: словарь vendor→coeff, содержащий
        только тех подрядчиков, для которых глобальный коэффициент
        включён. Отсутствие ключа означает, что глобальный коэффициент
        отключён и необходимо использовать оригинальный коэффициент
        позиции
    :return: словарь агрегированных сумм по подрядчикам
    """
    result: Dict[str, Dict[str, float]] = {}
    for it in items:
        v = it.vendor or "(без подрядчика)"
        # Подготовка словаря для подрядчика
        if v not in result:
            result[v] = {
                "equip_sum": 0.0,
                "other_sum": 0.0,
                "total_sum": 0.0,
            }
        # Определяем эффективный коэффициент для позиции класса equipment
        eff_coeff: Optional[float] = None
        if it.cls == "equipment":
            if v in vendor_coeffs_preview:
                # Глобальный коэффициент активен — применяем его
                eff_coeff = vendor_coeffs_preview[v]
            else:
                # Глобальный коэффициент отключён — используем original_coeff, если он известен
                if it.original_coeff is not None:
                    eff_coeff = it.original_coeff
                else:
                    # original_coeff отсутствует: используем значение coeff без изменений
                    eff_coeff = None
        # Рассчитываем сумму по позиции с учётом выбранного коэффициента
        amt = it.amount(effective_coeff=eff_coeff)
        # Суммируем по классам
        if it.cls == "equipment":
            result[v]["equip_sum"] += amt
        else:
            result[v]["other_sum"] += amt
        result[v]["total_sum"] += amt
    return result


def compute_client_flow(equip_sum: float, other_sum: float, discount_pct: float, commission_pct: float, tax_pct: float) -> Tuple[float, float, float, float, float]:
    """Порядок: Сумма → минус скидка → минус комиссия → плюс налог.
    Скидка и комиссия — только на equip_sum, налог — на общий промежуточный итог."""
    discount_amount = equip_sum * (discount_pct / 100.0)
    after_discount_equip = equip_sum - discount_amount
    commission_amount = after_discount_equip * (commission_pct / 100.0)
    after_commission_equip = after_discount_equip - commission_amount
    subtotal_before_tax = after_commission_equip + other_sum
    tax_amount = subtotal_before_tax * (tax_pct / 100.0)
    total_with_tax = subtotal_before_tax + tax_amount
    return discount_amount, commission_amount, tax_amount, subtotal_before_tax, total_with_tax


def compute_internal_discount(
    equip_sum: float,
    client_discount_amount: float,
    commission_amount: float,
    our_discount_pct: Optional[float],
    our_discount_sum: Optional[float],
) -> float:
    """
    Рассчитывает внутреннюю (нашу) скидку подрядчика.

    Порядок расчёта аналогичен вычислению итоговой суммы для клиента: из
    суммарного объёма работ класса «equipment» вычисляется реальная
    скидка подрядчика (либо процентом, либо фиксированной суммой). Затем
    из неё вычитаются две величины — скидка, предоставленная клиенту,
    и комиссия, уплаченная по итогам клиентского потока. Результат не
    может быть отрицательным.

    :param equip_sum: базовая сумма работ класса equipment (до скидок и комиссий)
    :param client_discount_amount: абсолютное значение скидки, предоставленной клиенту
    :param commission_amount: абсолютное значение комиссии, рассчитанной после скидки
    :param our_discount_pct: процент нашей скидки (если задан, имеет приоритет над our_discount_sum)
    :param our_discount_sum: фиксированная сумма нашей скидки (используется если процент не задан)
    :return: величина внутренней скидки (неотрицательное число)
    """
    try:
        # Определяем реальную скидку подрядчика: процент имеет приоритет над фиксированной суммой.
        real = 0.0
        if our_discount_pct is not None:
            real = equip_sum * (our_discount_pct / 100.0)
        elif our_discount_sum is not None:
            real = our_discount_sum
        # Вычитаем клиентскую скидку и комиссию, чтобы определить остаток нашей скидки.
        internal = real - client_discount_amount - commission_amount
        if internal < 0.0:
            internal = 0.0
        # Записываем информационную запись в лог о расчёте внутренней скидки
        logger.info(
            "compute_internal_discount: equip_sum=%.2f, real=%.2f, client_discount=%.2f, commission=%.2f, internal=%.2f",
            equip_sum,
            real,
            client_discount_amount,
            commission_amount,
            internal,
        )
        return internal
    except Exception:
        # В случае ошибки возвращаем ноль и пишем подробности в лог
        logger.error("Ошибка compute_internal_discount: %s", traceback.format_exc())
        return 0.0


def round2(x: float) -> float:
    """Округление до 2 знаков для отображения."""
    return float(f"{x:.2f}")


# 5. Виджет вкладки «Бухгалтерия»
class FinanceTab(QtWidgets.QWidget):
    """Главный виджет вкладки «Бухгалтерия» (две подкладки: Общая, Внутренняя)."""

    # Сигнал для внешней системы (если нужно отловить факт сохранения)
    saved = QtCore.Signal()

    def __init__(self, project_root: Optional[str] = None, project_id: Optional[str] = None, data_provider: Optional[FileDataProvider] = None, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        # 5.1 Инициализация провайдера данных
        try:
            if data_provider is not None:
                self.provider = data_provider
            else:
                # По умолчанию — файловый провайдер в текущей папке проекта
                if project_root is None:
                    # Если не передали путь проекта, используем папку модуля (упростим подключение)
                    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
                if project_id is None:
                    project_id = "default"
                self.provider = FileDataProvider(project_root=project_root, project_id=project_id)
        except Exception:
            logger.error("Ошибка инициализации провайдера данных: %s", traceback.format_exc())
            raise

        # 5.2 Загрузка исходных данных
        try:
            self.items: List[Item] = self.provider.load_items()
            self.vendors_settings: Dict[str, VendorSettings]
            loaded_vendors, loaded_profits, loaded_expenses = self.provider.load_finance()
            # Инициализируем настройки для всех подрядчиков, встречающихся в позициях
            self.vendors_settings = {}
            vendors_in_items = sorted({it.vendor or "(без подрядчика)" for it in self.items})
            for v in vendors_in_items:
                self.vendors_settings[v] = loaded_vendors.get(v, VendorSettings())
            # Доходы и расходы
            self.profit_items: List[ProfitItem] = loaded_profits
            self.expense_items: List[ExpenseItem] = loaded_expenses
        except Exception:
            logger.error("Ошибка загрузки данных: %s", traceback.format_exc())
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось загрузить данные для вкладки Бухгалтерия. Подробности в логах.")
            # Продолжаем с пустыми данными
            self.items = []
            self.vendors_settings = {}
            self.profit_items = []
            self.expense_items = []

        # 5.3 Состояние предпросмотра (применяются до сохранения)
        self.preview_vendor_coeffs: Dict[str, float] = {v: s.coeff for v, s in self.vendors_settings.items()}
        self.preview_discount_pct: Dict[str, float] = {v: s.discount_pct for v, s in self.vendors_settings.items()}
        self.preview_commission_pct: Dict[str, float] = {v: s.commission_pct for v, s in self.vendors_settings.items()}
        self.preview_tax_pct: Dict[str, float] = {v: s.tax_pct for v, s in self.vendors_settings.items()}
        self.preview_our_discount_pct: Dict[str, Optional[float]] = {v: s.our_discount_pct for v, s in self.vendors_settings.items()}
        self.preview_our_discount_sum: Dict[str, Optional[float]] = {v: s.our_discount_sum for v, s in self.vendors_settings.items()}

        # 5.3b Состояние предпросмотра сумм оплат. preview_paid[v] содержит сумму,
        # которую пользователь ввёл как уже выплаченную подрядчику. Это значение
        # отображается на подвкладке «Оплаты» и не записывается в реальные
        # настройки до нажатия кнопки «Сохранить».
        self.preview_paid: Dict[str, float] = {v: float(s.paid) for v, s in self.vendors_settings.items()}

        # 5.3a Состояние активности глобального коэффициента и отображаемые значения
        # preview_coeff_enabled[v] указывает, применяется ли глобальный коэффициент для vendor в предпросмотре.
        # _coeff_user_values[v] хранит фактическое значение коэффициента, введённое пользователем.
        # Настраиваем состояние включения коэффициента по загруженным данным.
        # Если coeff_enabled в настройках подрядчика установлен, берем его, иначе выключаем.
        self.preview_coeff_enabled: Dict[str, bool] = {v: s.coeff_enabled for v, s in self.vendors_settings.items()}
        self._coeff_user_values: Dict[str, float] = {v: s.coeff for v, s in self.vendors_settings.items()}

        # 5.4 Построение интерфейса
        self._build_ui()
        # 5.5 Первый пересчёт
        self.recalculate_all()

    # 5.a. Установка списка позиций из внешнего источника (например, базы данных)
    def set_items(self, items: List[Item]) -> None:
        """
        Обновляет внутренний список позиций и пересчитывает суммы.

        Этот метод может использоваться внешним кодом (например, build_finance_tab)
        для передачи списка позиций, загруженных из базы данных проекта.

        :param items: список объектов Item, представляющих позиции сметы
        """
        # Запоминаем новые позиции
        self.items = items
        # Обновляем список подрядчиков и их настройки
        vendors_in_items = sorted({it.vendor or "(без подрядчика)" for it in self.items})
        # Добавляем новые подрядчики с настройками по умолчанию
        for v in vendors_in_items:
            if v not in self.vendors_settings:
                self.vendors_settings[v] = VendorSettings()
        # Удаляем подрядчиков, отсутствующих в списке позиций
        for v in list(self.vendors_settings.keys()):
            if v not in vendors_in_items:
                self.vendors_settings.pop(v)
        # Синхронизируем предпросмотрные коэффициенты и параметры
        self.preview_vendor_coeffs = {v: s.coeff for v, s in self.vendors_settings.items()}
        self.preview_discount_pct = {v: s.discount_pct for v, s in self.vendors_settings.items()}
        self.preview_commission_pct = {v: s.commission_pct for v, s in self.vendors_settings.items()}
        self.preview_tax_pct = {v: s.tax_pct for v, s in self.vendors_settings.items()}
        self.preview_our_discount_pct = {v: s.our_discount_pct for v, s in self.vendors_settings.items()}
        self.preview_our_discount_sum = {v: s.our_discount_sum for v, s in self.vendors_settings.items()}
        # Синхронизируем предварительные суммы оплат. Для новых подрядчиков — 0.
        self.preview_paid = {v: float(self.vendors_settings[v].paid) for v in self.vendors_settings.keys()}
        # Для каждой вновь добавленной строки сбрасываем активность коэффициента и пользовательское значение
        # Загружаем состояние включения/выключения из настроек подрядчиков
        self.preview_coeff_enabled = {v: self.vendors_settings[v].coeff_enabled for v in self.vendors_settings.keys()}
        self._coeff_user_values = {v: s.coeff for v, s in self.vendors_settings.items()}
        # Пересчёт с новыми данными
        self.recalculate_all()

    # 5.b. Замена провайдера и перезагрузка данных
    def set_provider(self, provider: object) -> None:
        """Меняет провайдер данных для вкладки «Бухгалтерия» и перезагружает состояние.

        Используется при смене текущего проекта: при открытии нового проекта
        «Бухгалтерия» должна читать данные из базы, а не из файлового
        провайдера. Этот метод выполняет полную перезагрузку позиций,
        настроек подрядчиков, доходов и расходов, а затем пересчитывает
        таблицы.

        :param provider: новый провайдер данных (обычно DBDataProvider)
        """
        try:
            # Обновляем ссылку на провайдер
            self.provider = provider
            # Загружаем позиции проекта
            items = self.provider.load_items() or []
            # Загружаем финансовые настройки
            vendors, profits, expenses = self.provider.load_finance()
            # Инициализируем настройки подрядчиков на основе новых данных
            self.vendors_settings = {}
            vendors_in_items = sorted({it.vendor or "(без подрядчика)" for it in items})
            for v in vendors_in_items:
                self.vendors_settings[v] = vendors.get(v, VendorSettings())
            # Обновляем списки доходов/расходов
            self.profit_items = profits
            self.expense_items = expenses
            # Сохраняем новые позиции
            self.items = items
            # Синхронизируем предпросмотр
            self.preview_vendor_coeffs = {v: s.coeff for v, s in self.vendors_settings.items()}
            self.preview_discount_pct = {v: s.discount_pct for v, s in self.vendors_settings.items()}
            self.preview_commission_pct = {v: s.commission_pct for v, s in self.vendors_settings.items()}
            self.preview_tax_pct = {v: s.tax_pct for v, s in self.vendors_settings.items()}
            self.preview_our_discount_pct = {v: s.our_discount_pct for v, s in self.vendors_settings.items()}
            self.preview_our_discount_sum = {v: s.our_discount_sum for v, s in self.vendors_settings.items()}
            # Синхронизируем предварительные суммы оплат
            self.preview_paid = {v: float(s.paid) for v, s in self.vendors_settings.items()}
            # При смене провайдера сбрасываем активность коэффициентов и запоминаем введённые значения
            # Загружаем состояние включения/выключения из настроек подрядчиков
            self.preview_coeff_enabled = {v: self.vendors_settings[v].coeff_enabled for v in self.vendors_settings.keys()}
            self._coeff_user_values = {v: s.coeff for v, s in self.vendors_settings.items()}
            # Пересчитываем UI с новыми данными
            self.recalculate_all()
            logger.info("Провайдер данных обновлён и данные перезагружены")
        except Exception:
            logger.error("Ошибка смены провайдера: %s", traceback.format_exc())

    # 5.c. Слот обработки сигнала изменения сводной сметы
    def on_summary_changed(self) -> None:
        """Обработчик сигнала summary_changed.

        Этот метод вызывается, когда в сводной смете происходят изменения (добавление,
        редактирование, удаление позиций). Он запрашивает актуальные данные
        напрямую из базы данных через объект ProjectPage, конструирует список
        объектов Item и передаёт его в метод set_items(), тем самым обновляя
        таблицы и пересчитывая суммы.
        """
        try:
            # Получаем ссылку на ProjectPage, установленную в build_finance_tab
            page = getattr(self, "_page", None)
            if page is None:
                return
            # Проверяем наличие базы и project_id
            db = getattr(page, "db", None)
            pid = getattr(page, "project_id", None)
            if db is None or pid is None:
                return
            # Получаем строки из базы
            rows = db.list_items(pid)
            items: List[Item] = []
            for row in rows:
                try:
                    item_id = str(row["id"]) if "id" in row.keys() else ""
                    vendor = ""
                    if "vendor" in row.keys() and row["vendor"] is not None:
                        vendor = str(row["vendor"]).strip()
                    vendor = vendor or "(без подрядчика)"
                    cls = "equipment"
                    if "type" in row.keys() and row["type"]:
                        cls = row["type"]
                    department = ""
                    if "department" in row.keys() and row["department"]:
                        department = row["department"]
                    zone = ""
                    if "zone" in row.keys() and row["zone"]:
                        zone = row["zone"]
                    name = ""
                    if "name" in row.keys() and row["name"]:
                        name = row["name"]
                    unit_price = 0.0
                    if "unit_price" in row.keys() and row["unit_price"] is not None:
                        unit_price = float(row["unit_price"])
                    qty = 0.0
                    if "qty" in row.keys() and row["qty"] is not None:
                        qty = float(row["qty"])
                    coeff = 1.0
                    if "coeff" in row.keys() and row["coeff"] is not None:
                        coeff = float(row["coeff"])
                    items.append(Item(
                        id=item_id,
                        vendor=vendor,
                        cls=cls,
                        department=department,
                        zone=zone,
                        name=name,
                        price=unit_price,
                        qty=qty,
                        coeff=coeff,
                    ))
                except Exception:
                    logger.error("on_summary_changed: ошибка обработки строки: %s", traceback.format_exc())
                    continue
            # Передаём список позиций в метод set_items()
            self.set_items(items)
            logger.info("Финансы: обновлено %s позиций после изменения сметы", len(items))
        except Exception:
            logger.error("Ошибка обработки сигнала summary_changed: %s", traceback.format_exc())

    # 6. Построение интерфейса
    def _build_ui(self) -> None:
        # Главный лэйаут вкладки
        main_layout = QtWidgets.QVBoxLayout(self)

        # Подвкладки
        self.tabs = QtWidgets.QTabWidget(self)
        main_layout.addWidget(self.tabs, 1)

        # Общая
        self.tab_general = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_general, "Общая")
        self._build_tab_general(self.tab_general)

        # Внутренняя
        self.tab_internal = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_internal, "Внутренняя")
        self._build_tab_internal(self.tab_internal)

        # 8.c Новая вкладка «Оплаты» для учёта уже выплаченных сумм подрядчикам
        self.tab_payments = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_payments, "Оплаты")
        self._build_tab_payments(self.tab_payments)

        # Большая красная кнопка «Сохранить изменения»
        self.save_btn = QtWidgets.QPushButton("Сохранить изменения")
        self.save_btn.setStyleSheet("QPushButton { background: #cc0000; color: white; font-weight: 700; padding: 12px; border-radius: 6px; }\nQPushButton:hover { background: #e60000; }")
        self.save_btn.clicked.connect(self.on_save_clicked)
        main_layout.addWidget(self.save_btn, 0)

    # 7. Построение подвкладки «Общая»
    def _build_tab_general(self, parent: QtWidgets.QWidget) -> None:
        layout = QtWidgets.QVBoxLayout(parent)

        # 7.1 Таблица подрядчиков
        # Расширяем таблицу до 14 колонок, добавляя флажок включения коэффициента
        self.vendors_table = QtWidgets.QTableWidget(0, 14, parent)
        self.vendors_table.setHorizontalHeaderLabels([
            "Подрядчик",                # 0
            "Сумма equipment",          # 1
            "Глоб. коэф. вкл",          # 2 (checkbox для включения/отключения)
            "Коэф.",                    # 3 (editable)
            "Скидка %",                 # 4 (editable)
            "Скидка ₽",                 # 5
            "Комиссия %",               # 6 (editable)
            "Комиссия ₽",               # 7
            "Налог %",                  # 8 (editable)
            "Налог ₽",                  # 9
            "Итого с налогом",          # 10
            "Внутр. скидка ₽",          # 11 (из Внутренней)
            "Доходы из сметы ₽",        # 12 (привязанные к подрядчику)
            "Должны подрядчику ₽",      # 13 (итог для выплат)
        ])
        self.vendors_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.vendors_table.verticalHeader().setVisible(False)
        self.vendors_table.setAlternatingRowColors(True)
        layout.addWidget(self.vendors_table, 1)

        # 7.2 Сводные блоки
        summary_box = QtWidgets.QGroupBox("Сводные суммы")
        gl = QtWidgets.QGridLayout(summary_box)

        self.lbl_total_project = QtWidgets.QLabel("Итого проект: 0.00 ₽")  # с налогом
        self.lbl_by_departments = QtWidgets.QLabel("По отделам: —")        # текст сводки
        self.lbl_by_classes = QtWidgets.QLabel("По классам: —")            # текст сводки
        self.lbl_by_zones = QtWidgets.QLabel("По зонам: —")                # текст сводки

        gl.addWidget(self.lbl_total_project, 0, 0, 1, 2)
        gl.addWidget(self.lbl_by_departments, 1, 0, 1, 2)
        gl.addWidget(self.lbl_by_classes, 2, 0, 1, 2)
        gl.addWidget(self.lbl_by_zones, 3, 0, 1, 2)

        layout.addWidget(summary_box, 0)

        # Кнопка сброса коэффициентов до исходного значения (из импортированной сметы)
        btn_reset = QtWidgets.QPushButton("Сбросить коэффициенты")
        btn_reset.setToolTip("Вернуть глобальные коэффициенты подрядчиков к значениям, импортированным из сметы, и отключить их применение")
        btn_reset.clicked.connect(self._on_reset_coefficients_clicked)
        layout.addWidget(btn_reset, 0)

    # 8. Построение подвкладки «Внутренняя»
    def _build_tab_internal(self, parent: QtWidgets.QWidget) -> None:
        layout = QtWidgets.QHBoxLayout(parent)

        # Левая колонка: Доходы (внутр. скидки + из сметы)
        left = QtWidgets.QVBoxLayout()
        layout.addLayout(left, 1)

        # 8.1 Таблица «Наша скидка по подрядчикам»
        grp_our = QtWidgets.QGroupBox("Наша скидка по подрядчикам (заполняется: % ИЛИ сумма) — учитывается только на equipment")
        left.addWidget(grp_our, 1)
        vbox_our = QtWidgets.QVBoxLayout(grp_our)

        # Таблица «Наша скидка…» теперь имеет 6 колонок: последняя отображает,
        # сколько нам остаётся после учёта клиентской скидки и комиссии
        self.tbl_our_discount = QtWidgets.QTableWidget(0, 6, grp_our)
        self.tbl_our_discount.setHorizontalHeaderLabels([
            "Подрядчик",
            "Наша скидка %",
            "Наша скидка ₽",
            "Клиентская скидка % (для справки)",
            "Сумма скидки ₽",
            "Сколько остаётся ₽",
        ])
        self.tbl_our_discount.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tbl_our_discount.verticalHeader().setVisible(False)
        self.tbl_our_discount.setAlternatingRowColors(True)
        vbox_our.addWidget(self.tbl_our_discount, 1)

        # 8.2 Таблица «Доходы из сметы»
        grp_profit = QtWidgets.QGroupBox("Доходы из сметы (привязка к подрядчику)")
        left.addWidget(grp_profit, 1)
        vbox_profit = QtWidgets.QVBoxLayout(grp_profit)

        self.tbl_profit = QtWidgets.QTableWidget(0, 3, grp_profit)
        self.tbl_profit.setHorizontalHeaderLabels(["Подрядчик", "Описание", "Сумма ₽"])
        self.tbl_profit.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tbl_profit.verticalHeader().setVisible(False)
        self.tbl_profit.setAlternatingRowColors(True)
        vbox_profit.addWidget(self.tbl_profit, 1)

        btns_profit = QtWidgets.QHBoxLayout()
        # Кнопка добавления доходов, сформированных из сметы
        self.btn_add_profit = QtWidgets.QPushButton("Добавить из сметы…")
        # Новая кнопка для ручного добавления дохода (чаевые, случайный доход и т.п.)
        self.btn_add_manual_profit = QtWidgets.QPushButton("Добавить доход")
        # Кнопка для удаления выбранных доходов
        self.btn_remove_profit = QtWidgets.QPushButton("Удалить выбранное")
        btns_profit.addWidget(self.btn_add_profit, 0)
        btns_profit.addWidget(self.btn_add_manual_profit, 0)
        btns_profit.addWidget(self.btn_remove_profit, 0)
        vbox_profit.addLayout(btns_profit)

        # Соединяем сигналы с обработчиками
        self.btn_add_profit.clicked.connect(self.on_add_profit_clicked)
        self.btn_add_manual_profit.clicked.connect(self.on_add_manual_profit_clicked)
        self.btn_remove_profit.clicked.connect(self.on_remove_profit_clicked)

        # 8.3 Итого доходов
        self.lbl_income_total = QtWidgets.QLabel("Итого доходы: 0.00 ₽ (внутр. скидки + из сметы)")
        left.addWidget(self.lbl_income_total, 0)

        # Правая колонка: Расходы
        right = QtWidgets.QVBoxLayout()
        layout.addLayout(right, 1)

        grp_exp = QtWidgets.QGroupBox("Расходы (ручной ввод)")
        right.addWidget(grp_exp, 1)
        vbox_exp = QtWidgets.QVBoxLayout(grp_exp)

        self.tbl_expense = QtWidgets.QTableWidget(0, 4, grp_exp)
        self.tbl_expense.setHorizontalHeaderLabels(["Расход", "Кол-во", "Цена ₽", "Сумма ₽"])
        self.tbl_expense.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tbl_expense.verticalHeader().setVisible(False)
        self.tbl_expense.setAlternatingRowColors(True)
        vbox_exp.addWidget(self.tbl_expense, 1)

        btns_exp = QtWidgets.QHBoxLayout()
        self.btn_add_exp = QtWidgets.QPushButton("Добавить расход")
        self.btn_remove_exp = QtWidgets.QPushButton("Удалить выбранное")
        btns_exp.addWidget(self.btn_add_exp, 0)
        btns_exp.addWidget(self.btn_remove_exp, 0)
        vbox_exp.addLayout(btns_exp)

        self.btn_add_exp.clicked.connect(self.on_add_exp_clicked)
        self.btn_remove_exp.clicked.connect(self.on_remove_exp_clicked)

        # Подключаем обработчик изменения ячеек таблицы расходов. Это необходимо, чтобы
        # сохранять текст в столбике «Расход» (название расхода). Без этого текст,
        # введённый пользователем в таблицу, не попадал в модель и не сохранялся.
        try:
            self.tbl_expense.itemChanged.connect(self._on_expense_name_changed)
        except Exception:
            logger.error(
                "FinanceTab: не удалось подключить сигнал itemChanged для таблицы расходов", exc_info=True
            )

        # 8.4 Итого расходов и чистая прибыль
        self.lbl_expense_total = QtWidgets.QLabel("Итого расходы: 0.00 ₽")
        self.lbl_net_total = QtWidgets.QLabel("Итого после вычета расходов: 0.00 ₽")
        right.addWidget(self.lbl_expense_total, 0)
        right.addWidget(self.lbl_net_total, 0)

    # 8.c. Построение подвкладки «Оплаты»
    def _build_tab_payments(self, parent: QtWidgets.QWidget) -> None:
        """
        Создаёт интерфейс для подвкладки «Оплаты». Здесь выводится
        таблица подрядчиков с расчётной суммой задолженности, полем ввода
        выплаченной суммы и остатком. Изменение поля «Оплачено» влияет
        только на предпросмотр и сохраняется в конфигурацию при
        нажатии кнопки «Сохранить изменения».
        """
        layout = QtWidgets.QVBoxLayout(parent)

        # Таблица оплат: Подрядчик | Должны ₽ | Оплачено ₽ | Остаток ₽
        self.tbl_payments = QtWidgets.QTableWidget(0, 4, parent)
        self.tbl_payments.setHorizontalHeaderLabels([
            "Подрядчик",
            "Должны ₽",
            "Оплачено ₽",
            "Остаток ₽",
        ])
        self.tbl_payments.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.tbl_payments.verticalHeader().setVisible(False)
        self.tbl_payments.setAlternatingRowColors(True)
        layout.addWidget(self.tbl_payments, 1)

        # Примечание
        note = QtWidgets.QLabel(
            "Введите сумму, уже выплаченную подрядчику. Остаток рассчитывается автоматически."
        )
        note.setWordWrap(True)
        layout.addWidget(note, 0)


    # 9. Пересчёт таблиц (предпросмотр)
    def recalculate_all(self) -> None:
        try:
            # 9.1 Формируем словарь эффективных коэффициентов для аггрегации.
            # Если для подрядчика глобальный коэффициент отключён, используем 1.0,
            # иначе берём пользовательское значение.
            effective_coeffs: Dict[str, float] = {}
            for v in self.vendors_settings.keys():
                if self.preview_coeff_enabled.get(v, False):
                    # Если коэффициент включён, применяем пользовательское значение для расчёта
                    val = self._coeff_user_values.get(v, self.preview_vendor_coeffs.get(v, 1.0))
                    effective_coeffs[v] = float(val)
                else:
                    # Если коэффициент выключен, используем коэффициент каждой позиции (item.coeff).
                    # Здесь мы не переопределяем коэффициент в effective_coeffs, чтобы аггрегатор
                    # использовал исходные значения позиций. Не обновляем preview_vendor_coeffs.
                    pass

            # 9.1 Аггрегация по подрядчикам
            agg = aggregate_by_vendor(self.items, effective_coeffs)
            # Сохраняем аггрегированные данные для использования в других методах
            self._agg_latest = agg

            # 9.2 Пересчитываем «Общая»
            self._fill_vendors_table(agg)

            # 9.3 Пересчитываем «Внутренняя»
            self._fill_our_discount_table()
            self._fill_profit_table()
            self._fill_expense_table()

            # 9.3b Пересчитываем «Оплаты»
            self._fill_payments_table(agg)

            # 9.4 Итоги доходов/расходов/прибыли
            income_total = self._calc_income_total(agg)
            expense_total = self._calc_expense_total()
            net = income_total - expense_total
            self.lbl_income_total.setText(f"Итого доходы: {round2(income_total):,.2f} ₽".replace(",", " "))
            self.lbl_expense_total.setText(f"Итого расходы: {round2(expense_total):,.2f} ₽".replace(",", " "))
            self.lbl_net_total.setText(f"Итого после вычета расходов: {round2(net):,.2f} ₽".replace(",", " "))

            # 9.5 Сводные суммы для «Общая»
            # Здесь применяем ту же логику, что и в aggregate_by_vendor: если
            # глобальный коэффициент для подрядчика включён, используем
            # пользовательское значение; иначе — возвращаемся к оригинальному
            # коэффициенту позиции (если он хранится в original_coeff), или
            # используем текущий coeff. Это гарантирует, что сводные суммы
            # совпадают с отображаемыми агрегатами.
            total_project = 0.0
            by_dep: Dict[str, float] = {}
            by_cls: Dict[str, float] = {}
            by_zone: Dict[str, float] = {}
            for it in self.items:
                v = it.vendor or "(без подрядчика)"
                eff: Optional[float] = None
                if it.cls == "equipment":
                    if self.preview_coeff_enabled.get(v, False):
                        # Коэффициент включён — берём введённое пользователем значение
                        eff = self._coeff_user_values.get(v, self.preview_vendor_coeffs.get(v, 1.0))
                    else:
                        # Коэффициент отключён — используем original_coeff, если задан
                        if it.original_coeff is not None:
                            eff = it.original_coeff
                        else:
                            eff = None
                # Вычисляем сумму строки
                amt = it.amount(effective_coeff=eff)
                total_project += amt
                # По отделам
                by_dep[it.department] = by_dep.get(it.department, 0.0) + amt
                # По классам (отображаем русское название класса)
                cls_key = CLASS_EN2RU.get(it.cls, it.cls)
                by_cls[cls_key] = by_cls.get(cls_key, 0.0) + amt
                # По зонам
                by_zone[it.zone] = by_zone.get(it.zone, 0.0) + amt
            # Обновляем подписи
            self.lbl_total_project.setText(
                f"Итого проект (без клиентских скидок/комиссий/налога): {round2(total_project):,.2f} ₽".replace(",", " ")
            )
            self.lbl_by_departments.setText(
                "По отделам: " + "; ".join([
                    f"{k}: {round2(v):,.2f} ₽".replace(",", " ") for k, v in sorted(by_dep.items())
                ])
            )
            self.lbl_by_classes.setText(
                "По классам: " + "; ".join([
                    f"{k}: {round2(v):,.2f} ₽".replace(",", " ") for k, v in sorted(by_cls.items())
                ])
            )
            self.lbl_by_zones.setText(
                "По зонам: " + "; ".join([
                    f"{k}: {round2(v):,.2f} ₽".replace(",", " ") for k, v in sorted(by_zone.items())
                ])
            )

        except Exception:
            logger.error("Ошибка пересчёта: %s", traceback.format_exc())
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось пересчитать значения. Подробности в логах.")

    # 10. Заполнение таблицы подрядчиков («Общая»)
    def _fill_vendors_table(self, agg: Dict[str, Dict[str, float]]) -> None:
        self.vendors_table.blockSignals(True)
        self.vendors_table.setRowCount(0)

        for row, vendor in enumerate(sorted(agg.keys())):
            data = agg[vendor]
            equip_sum = data["equip_sum"]
            other_sum = data["other_sum"]
            s = self.vendors_settings.get(vendor, VendorSettings())
            # Значения для предпросмотра
            coeff_user = self._coeff_user_values.get(vendor, s.coeff)
            discount_pct = self.preview_discount_pct.get(vendor, s.discount_pct)
            commission_pct = self.preview_commission_pct.get(vendor, s.commission_pct)
            tax_pct = self.preview_tax_pct.get(vendor, s.tax_pct)

            # Клиентский поток
            discount_amount, commission_amount, tax_amount, subtotal_before_tax, total_with_tax = compute_client_flow(
                equip_sum, other_sum, discount_pct, commission_pct, tax_pct
            )

            # Внутренняя скидка (для вывода)
            # Здесь комиссия вычитается из нашей скидки так же, как и скидка клиента
            internal = compute_internal_discount(
                equip_sum,
                discount_amount,
                commission_amount,
                self.preview_our_discount_pct.get(vendor),
                self.preview_our_discount_sum.get(vendor),
            )

            # Доходы из сметы для этого подрядчика
            profit_from_vendor = sum(p.amount for p in self.profit_items if (p.vendor or "") == (vendor or ""))

            # Сколько должны подрядчику (после налога, минус внутр. скидка и доходы из сметы)
            owe_vendor = total_with_tax - internal - profit_from_vendor

            self.vendors_table.insertRow(row)

            # 0 Подрядчик
            item_vendor = QtWidgets.QTableWidgetItem(vendor)
            item_vendor.setFlags(item_vendor.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 0, item_vendor)

            # 1 Сумма equipment (readonly)
            item_e = QtWidgets.QTableWidgetItem(f"{round2(equip_sum):,.2f}".replace(",", " "))
            item_e.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_e.setFlags(item_e.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 1, item_e)

            # 2 Чекбокс включения глобального коэффициента
            chk = QtWidgets.QCheckBox()
            chk.setChecked(self.preview_coeff_enabled.get(vendor, True))
            # Подпишемся на изменение состояния
            chk.stateChanged.connect(lambda state, v=vendor: self._on_vendor_coeff_enabled_toggled(v, bool(state)))
            # Выравниваем чекбокс по центру
            w_cw = QtWidgets.QWidget()
            h_layout = QtWidgets.QHBoxLayout(w_cw)
            h_layout.setContentsMargins(0, 0, 0, 0)
            h_layout.setAlignment(QtCore.Qt.AlignCenter)
            h_layout.addWidget(chk)
            self.vendors_table.setCellWidget(row, 2, w_cw)

            # 3 Коэф. (редактируемый с отключёнными стрелками). Значение отображается из _coeff_user_values
            w_coeff = SmartDoubleSpinBox()
            w_coeff.setRange(0.0, 1000.0)
            w_coeff.setDecimals(3)
            w_coeff.setSingleStep(0.1)
            # Убираем стрелочки и отключаем автоматическое обновление при наборе
            w_coeff.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_coeff.setKeyboardTracking(False)
            w_coeff.setValue(float(coeff_user))
            # Включаем/отключаем поле в зависимости от состояния чекбокса
            w_coeff.setEnabled(self.preview_coeff_enabled.get(vendor, True))
            # При изменении сохраняем значение в _coeff_user_values и в предпросмотр, если коэффициент активен
            w_coeff.valueChanged.connect(lambda val, v=vendor: self._on_vendor_coeff_changed(v, val))
            self.vendors_table.setCellWidget(row, 3, w_coeff)

            # 4 Скидка % (редактируемая без стрелок)
            w_disc = SmartDoubleSpinBox()
            w_disc.setRange(0.0, 100.0)
            w_disc.setDecimals(2)
            w_disc.setSingleStep(0.5)
            w_disc.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_disc.setKeyboardTracking(False)
            w_disc.setValue(float(discount_pct))
            w_disc.valueChanged.connect(lambda val, v=vendor: self._on_vendor_discount_changed(v, val))
            self.vendors_table.setCellWidget(row, 4, w_disc)

            # 5 Скидка ₽ (readonly)
            item_da = QtWidgets.QTableWidgetItem(f"{round2(discount_amount):,.2f}".replace(",", " "))
            item_da.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_da.setFlags(item_da.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 5, item_da)

            # 6 Комиссия % (редактируемая без стрелок)
            w_comm = SmartDoubleSpinBox()
            w_comm.setRange(0.0, 100.0)
            w_comm.setDecimals(2)
            w_comm.setSingleStep(0.5)
            w_comm.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_comm.setKeyboardTracking(False)
            w_comm.setValue(float(commission_pct))
            w_comm.valueChanged.connect(lambda val, v=vendor: self._on_vendor_commission_changed(v, val))
            self.vendors_table.setCellWidget(row, 6, w_comm)

            # 7 Комиссия ₽ (readonly)
            item_ca = QtWidgets.QTableWidgetItem(f"{round2(commission_amount):,.2f}".replace(",", " "))
            item_ca.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_ca.setFlags(item_ca.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 7, item_ca)

            # 8 Налог % (редактируемый без стрелок)
            w_tax = SmartDoubleSpinBox()
            w_tax.setRange(0.0, 100.0)
            w_tax.setDecimals(2)
            w_tax.setSingleStep(0.5)
            w_tax.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_tax.setKeyboardTracking(False)
            w_tax.setValue(float(tax_pct))
            w_tax.valueChanged.connect(lambda val, v=vendor: self._on_vendor_tax_changed(v, val))
            self.vendors_table.setCellWidget(row, 8, w_tax)

            # 9 Налог ₽ (readonly)
            item_ta = QtWidgets.QTableWidgetItem(f"{round2(tax_amount):,.2f}".replace(",", " "))
            item_ta.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_ta.setFlags(item_ta.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 9, item_ta)

            # 10 Итого с налогом (readonly)
            item_tot = QtWidgets.QTableWidgetItem(f"{round2(total_with_tax):,.2f}".replace(",", " "))
            item_tot.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_tot.setFlags(item_tot.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 10, item_tot)

            # 11 Внутр. скидка ₽ (readonly)
            item_int = QtWidgets.QTableWidgetItem(f"{round2(internal):,.2f}".replace(",", " "))
            item_int.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_int.setFlags(item_int.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 11, item_int)

            # 12 Доходы из сметы ₽ (readonly)
            item_pf = QtWidgets.QTableWidgetItem(f"{round2(profit_from_vendor):,.2f}".replace(",", " "))
            item_pf.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_pf.setFlags(item_pf.flags() ^ QtCore.Qt.ItemIsEditable)
            self.vendors_table.setItem(row, 12, item_pf)

            # 13 Должны подрядчику ₽ (readonly, текст красным цветом без заливки)
            item_owe = QtWidgets.QTableWidgetItem(f"{round2(owe_vendor):,.2f}".replace(",", " "))
            item_owe.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_owe.setFlags(item_owe.flags() ^ QtCore.Qt.ItemIsEditable)
            # Убираем заливку и делаем текст красным
            item_owe.setBackground(QtGui.QColor(QtCore.Qt.transparent))
            item_owe.setForeground(QtGui.QColor(200, 0, 0))
            self.vendors_table.setItem(row, 13, item_owe)

        self.vendors_table.blockSignals(False)

    # 11. Заполнение таблицы «Наша скидка…»
    def _fill_our_discount_table(self) -> None:
        self.tbl_our_discount.blockSignals(True)
        self.tbl_our_discount.setRowCount(0)
        for row, vendor in enumerate(sorted(self.vendors_settings.keys())):
            s = self.vendors_settings[vendor]
            our_pct = self.preview_our_discount_pct.get(vendor, s.our_discount_pct)
            our_sum = self.preview_our_discount_sum.get(vendor, s.our_discount_sum)
            client_pct = self.preview_discount_pct.get(vendor, s.discount_pct)
            # Получаем сумму equipment для расчёта суммы скидки
            equip_sum = 0.0
            try:
                if hasattr(self, "_agg_latest"):
                    equip_sum = self._agg_latest.get(vendor, {}).get("equip_sum", 0.0)
            except Exception:
                equip_sum = 0.0
            # Реальная наша скидка (до учёта клиентской) — для справки
            discount_amount = 0.0
            if our_pct is not None:
                discount_amount = float(equip_sum) * (float(our_pct) / 100.0)
            elif our_sum is not None:
                discount_amount = float(our_sum)

            self.tbl_our_discount.insertRow(row)

            # 0 Подрядчик
            item_v = QtWidgets.QTableWidgetItem(vendor)
            item_v.setFlags(item_v.flags() ^ QtCore.Qt.ItemIsEditable)
            self.tbl_our_discount.setItem(row, 0, item_v)

            # 1 Наша скидка % (редактируемая без стрелок)
            w_pct = SmartDoubleSpinBox()
            w_pct.setRange(0.0, 100.0)
            w_pct.setDecimals(2)
            w_pct.setSpecialValueText("(нет)")
            w_pct.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_pct.setKeyboardTracking(False)
            w_pct.setValue(float(our_pct) if our_pct is not None else 0.0)
            w_pct.valueChanged.connect(lambda val, v=vendor: self._on_our_discount_pct_changed(v, val))
            self.tbl_our_discount.setCellWidget(row, 1, w_pct)

            # 2 Наша скидка ₽ (редактируемая без стрелок)
            w_sum = SmartDoubleSpinBox()
            w_sum.setRange(0.0, 10_000_000_000.0)
            w_sum.setDecimals(2)
            w_sum.setSpecialValueText("(нет)")
            w_sum.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_sum.setKeyboardTracking(False)
            w_sum.setValue(float(our_sum) if our_sum is not None else 0.0)
            w_sum.valueChanged.connect(lambda val, v=vendor: self._on_our_discount_sum_changed(v, val))
            self.tbl_our_discount.setCellWidget(row, 2, w_sum)

            # 3 Клиентская скидка % (readonly, для справки)
            item_c = QtWidgets.QTableWidgetItem(f"{round2(client_pct):,.2f}".replace(",", " "))
            item_c.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_c.setFlags(item_c.flags() ^ QtCore.Qt.ItemIsEditable)
            self.tbl_our_discount.setItem(row, 3, item_c)

            # 4 Сумма скидки ₽ (readonly)
            item_sum = QtWidgets.QTableWidgetItem(f"{round2(discount_amount):,.2f}".replace(",", " "))
            item_sum.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_sum.setFlags(item_sum.flags() ^ QtCore.Qt.ItemIsEditable)
            self.tbl_our_discount.setItem(row, 4, item_sum)

            # 5 Сколько остаётся ₽ (readonly). Рассчитываем как нашу скидку минус
            # клиентскую скидку и комиссию. Для вычисления используем функцию
            # compute_internal_discount() и предварительные значения.
            try:
                # Абсолютная скидка, предоставленная клиенту
                client_discount_amount = float(equip_sum) * (float(client_pct) / 100.0)
                # Комиссия берётся из предпросмотра или настроек подрядчика
                commission_pct = self.preview_commission_pct.get(vendor, s.commission_pct)
                # Комиссия рассчитывается на сумму после клиентской скидки
                commission_amount = (float(equip_sum) - client_discount_amount) * (float(commission_pct) / 100.0)
                # Величина нашей внутренней скидки (то, что нам остаётся)
                internal_remain = compute_internal_discount(
                    float(equip_sum),
                    client_discount_amount,
                    commission_amount,
                    float(our_pct) if our_pct is not None else None,
                    float(our_sum) if our_sum is not None else None,
                )
            except Exception:
                logger.error("Ошибка расчёта внутренней скидки для '%s': %s", vendor, traceback.format_exc())
                internal_remain = 0.0
            item_rem = QtWidgets.QTableWidgetItem(f"{round2(internal_remain):,.2f}".replace(",", " "))
            item_rem.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_rem.setFlags(item_rem.flags() ^ QtCore.Qt.ItemIsEditable)
            self.tbl_our_discount.setItem(row, 5, item_rem)

        self.tbl_our_discount.blockSignals(False)

    # 12. Заполнение таблицы «Доходы из сметы»
    def _fill_profit_table(self) -> None:
        self.tbl_profit.blockSignals(True)
        self.tbl_profit.setRowCount(0)
        for row, p in enumerate(self.profit_items):
            self.tbl_profit.insertRow(row)
            self.tbl_profit.setItem(row, 0, QtWidgets.QTableWidgetItem(p.vendor))
            self.tbl_profit.setItem(row, 1, QtWidgets.QTableWidgetItem(p.description))
            item_amt = QtWidgets.QTableWidgetItem(f"{round2(p.amount):,.2f}".replace(",", " "))
            item_amt.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            self.tbl_profit.setItem(row, 2, item_amt)
        self.tbl_profit.blockSignals(False)

    # 13. Заполнение таблицы «Расходы»
    def _fill_expense_table(self) -> None:
        self.tbl_expense.blockSignals(True)
        self.tbl_expense.setRowCount(0)
        for row, e in enumerate(self.expense_items):
            self.tbl_expense.insertRow(row)
            self.tbl_expense.setItem(row, 0, QtWidgets.QTableWidgetItem(e.name))
            w_qty = QtWidgets.QDoubleSpinBox()
            w_qty.setRange(0.0, 1_000_000.0)
            w_qty.setDecimals(2)
            # Убираем стрелочки для удобства ввода — пользователь может вводить значение вручную
            w_qty.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_qty.setKeyboardTracking(False)  # значение меняется только после подтверждения ввода
            w_qty.setValue(float(e.qty))
            # связываем изменение значения с обновлением данных; при keyboardTracking=False
            # valueChanged сработает только после ввода полного числа, что предотвращает
            # преждевременный пересчёт и пересоздание таблицы при наборе цифр
            w_qty.valueChanged.connect(lambda val, r=row: self._on_expense_qty_changed(r, val))
            self.tbl_expense.setCellWidget(row, 1, w_qty)

            w_price = QtWidgets.QDoubleSpinBox()
            w_price.setRange(0.0, 1_000_000_000.0)
            w_price.setDecimals(2)
            # Убираем стрелочки для удобства ввода
            w_price.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            w_price.setKeyboardTracking(False)
            w_price.setValue(float(e.price))
            w_price.valueChanged.connect(lambda val, r=row: self._on_expense_price_changed(r, val))
            self.tbl_expense.setCellWidget(row, 2, w_price)

            item_total = QtWidgets.QTableWidgetItem(f"{round2(e.total()):,.2f}".replace(",", " "))
            item_total.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item_total.setFlags(item_total.flags() ^ QtCore.Qt.ItemIsEditable)
            self.tbl_expense.setItem(row, 3, item_total)
        self.tbl_expense.blockSignals(False)

    # 13.c. Заполнение таблицы «Оплаты»
    def _fill_payments_table(self, agg: Dict[str, Dict[str, float]]) -> None:
        """
        Заполняет таблицу «Оплаты» для каждого подрядчика. Сумма
        задолженности рассчитывается так же, как и в таблице «Общая»,
        без учёта уже выплаченных денег. Значение «Оплачено» берётся
        из ``self.preview_paid``. Изменения в поле «Оплачено» вызывают
        обработчик, который обновляет состояние предпросмотра и
        пересчитывает таблицу.

        :param agg: агрегированные суммы по подрядчикам (equip_sum и other_sum)
        """
        try:
            self.tbl_payments.blockSignals(True)
            self.tbl_payments.setRowCount(0)
            for row, vendor in enumerate(sorted(agg.keys())):
                data = agg[vendor]
                equip_sum = data["equip_sum"]
                other_sum = data["other_sum"]
                s = self.vendors_settings.get(vendor, VendorSettings())
                # Значения предпросмотра скидок/комиссий/налога
                discount_pct = self.preview_discount_pct.get(vendor, s.discount_pct)
                commission_pct = self.preview_commission_pct.get(vendor, s.commission_pct)
                tax_pct = self.preview_tax_pct.get(vendor, s.tax_pct)
                # Клиентский поток
                discount_amount, commission_amount, tax_amount, subtotal_before_tax, total_with_tax = compute_client_flow(
                    equip_sum, other_sum, discount_pct, commission_pct, tax_pct
                )
                # Внутренняя скидка: вычитаем клиентскую скидку и комиссию
                internal = compute_internal_discount(
                    equip_sum,
                    discount_amount,
                    commission_amount,
                    self.preview_our_discount_pct.get(vendor),
                    self.preview_our_discount_sum.get(vendor),
                )
                # Доходы из сметы для подрядчика
                profit_from_vendor = sum(p.amount for p in self.profit_items if (p.vendor or "") == (vendor or ""))
                # Расчёт суммы, которую должны подрядчику (до вычета оплат)
                owe_vendor = total_with_tax - internal - profit_from_vendor
                # Оплачено
                paid = float(self.preview_paid.get(vendor, 0.0))
                # Остаток
                remain = owe_vendor - paid
                # Добавляем строку
                self.tbl_payments.insertRow(row)
                # 0. Подрядчик (readonly)
                item_v = QtWidgets.QTableWidgetItem(vendor)
                item_v.setFlags(item_v.flags() ^ QtCore.Qt.ItemIsEditable)
                self.tbl_payments.setItem(row, 0, item_v)
                # 1. Должны ₽ (readonly)
                item_owe = QtWidgets.QTableWidgetItem(f"{round2(owe_vendor):,.2f}".replace(",", " "))
                item_owe.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                item_owe.setFlags(item_owe.flags() ^ QtCore.Qt.ItemIsEditable)
                self.tbl_payments.setItem(row, 1, item_owe)
                # 2. Оплачено ₽ (editable via SmartDoubleSpinBox)
                w_paid = SmartDoubleSpinBox()
                w_paid.setRange(0.0, 1_000_000_000.0)
                w_paid.setDecimals(2)
                # Скрываем стрелки и отключаем слежение за вводом
                w_paid.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
                w_paid.setKeyboardTracking(False)
                w_paid.setValue(float(paid))
                # Подключаем обработчик изменения
                w_paid.valueChanged.connect(lambda val, v=vendor: self._on_vendor_paid_changed(v, val))
                self.tbl_payments.setCellWidget(row, 2, w_paid)
                # 3. Остаток ₽ (readonly)
                item_rem = QtWidgets.QTableWidgetItem(f"{round2(remain):,.2f}".replace(",", " "))
                item_rem.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                item_rem.setFlags(item_rem.flags() ^ QtCore.Qt.ItemIsEditable)
                # Покрасим отрицательные остатки красным для наглядности
                if remain < 0:
                    item_rem.setForeground(QtGui.QColor(200, 0, 0))
                self.tbl_payments.setItem(row, 3, item_rem)
            self.tbl_payments.blockSignals(False)
        except Exception:
            logger.error("Ошибка заполнения таблицы оплат: %s", traceback.format_exc())


    # 14. Подсчёт сумм доходов/расходов
    def _calc_income_total(self, agg: Dict[str, Dict[str, float]]) -> float:
        total_internal = 0.0
        for vendor, data in agg.items():
            s = self.vendors_settings.get(vendor, VendorSettings())
            # Значения предпросмотра скидки и комиссии
            discount_pct = self.preview_discount_pct.get(vendor, s.discount_pct)
            commission_pct = self.preview_commission_pct.get(vendor, s.commission_pct)
            # Значения нашей скидки (процент или сумма)
            our_pct = self.preview_our_discount_pct.get(vendor, s.our_discount_pct)
            our_sum = self.preview_our_discount_sum.get(vendor, s.our_discount_sum)
            equip_sum = data["equip_sum"]
            # Абсолютная скидка для клиента
            client_disc_amount = equip_sum * (discount_pct / 100.0)
            # Абсолютная комиссия рассчитывается на сумму после скидки
            commission_amount = (equip_sum - client_disc_amount) * (commission_pct / 100.0)
            # Вычисляем внутреннюю скидку: наша скидка минус скидка клиента и комиссия
            internal = compute_internal_discount(
                equip_sum,
                client_disc_amount,
                commission_amount,
                our_pct,
                our_sum,
            )
            total_internal += internal

        total_profit = sum(p.amount for p in self.profit_items)
        return total_internal + total_profit

    def _calc_expense_total(self) -> float:
        return sum(e.total() for e in self.expense_items)

    # 15. Обработчики изменения значений в «Общая»
    def _on_vendor_coeff_changed(self, vendor: str, value: float) -> None:
        # Сохраняем введённое пользователем значение
        self._coeff_user_values[vendor] = float(value)
        # Если коэффициент включён, используем его в предпросмотре, иначе оставляем 1
        if self.preview_coeff_enabled.get(vendor, True):
            self.preview_vendor_coeffs[vendor] = float(value)
        logger.info("Предпросмотр: изменён коэффициент подрядчика '%s' → %.3f (только equipment)", vendor, value)
        self.recalculate_all()

    def _on_vendor_discount_changed(self, vendor: str, value: float) -> None:
        self.preview_discount_pct[vendor] = float(value)
        logger.info("Предпросмотр: клиентская скидка %% для '%s' → %.2f", vendor, value)
        self.recalculate_all()

    def _on_vendor_commission_changed(self, vendor: str, value: float) -> None:
        self.preview_commission_pct[vendor] = float(value)
        logger.info("Предпросмотр: комиссия %% для '%s' → %.2f", vendor, value)
        self.recalculate_all()

    def _on_vendor_tax_changed(self, vendor: str, value: float) -> None:
        self.preview_tax_pct[vendor] = float(value)
        logger.info("Предпросмотр: налог %% для '%s' → %.2f", vendor, value)
        self.recalculate_all()

    # 16.a. Обработчик включения/отключения глобального коэффициента
    def _on_vendor_coeff_enabled_toggled(self, vendor: str, enabled: bool) -> None:
        """
        При переключении чекбокса 'Глоб. коэф. вкл' этот метод обновляет
        состояние предосмотра: если коэффициент выключен, для аггрегации
        используется 1.0 и поле редактирования блокируется; если включён —
        применяется ранее введённое значение. После смены состояния
        вызывается пересчёт таблиц.
        """
        self.preview_coeff_enabled[vendor] = bool(enabled)
        # Если включаем, восстановим ранее введённый пользователем коэффициент
        if enabled:
            # Восстанавливаем пользовательское значение (оно уже сохранено в _coeff_user_values)
            val = self._coeff_user_values.get(vendor)
            if val is None:
                val = self.vendors_settings.get(vendor, VendorSettings()).coeff
            self.preview_vendor_coeffs[vendor] = float(val)
        else:
            # При выключении коэффициент не применяется, но значение оставляем без изменений
            pass
        logger.info("Предпросмотр: %s глобальный коэффициент для '%s'", "включён" if enabled else "отключён", vendor)
        self.recalculate_all()

    # 16.b. Сброс всех глобальных коэффициентов до исходных значений из сметы
    def _on_reset_coefficients_clicked(self) -> None:
        """
        Возвращает глобальные коэффициенты подрядчиков к значениям, которые были
        импортированы из сметы (хранятся в VendorSettings.coeff), и отключает их
        применение в предпросмотре.
        """
        try:
            # 1. Восстановление глобальных коэффициентов подрядчиков: возвращаем значения,
            # импортированные из сметы (VendorSettings.coeff) и отключаем их применение.
            for v in self.vendors_settings.keys():
                s = self.vendors_settings[v]
                # Пользовательский коэффициент для предпросмотра возвращается к исходному
                self._coeff_user_values[v] = float(s.coeff)
                self.preview_vendor_coeffs[v] = float(s.coeff)
                # Отключаем применение глобального коэффициента: расчёт будет использовать
                # исходные коэффициенты позиций
                self.preview_coeff_enabled[v] = False
            # 2. Восстанавливаем коэффициенты у каждой позиции класса equipment до исходных
            # значений (original_coeff). Это необходимо, если пользователь сохранял проект
            # после применения глобальных коэффициентов: поле Item.coeff в items
            # изменилось, и его нужно вернуть к оригиналу.
            for it in self.items:
                if it.cls == "equipment":
                    # Возвращаем коэффициент к исходному (если известно), иначе оставляем текущий
                    if it.original_coeff is not None:
                        try:
                            it.coeff = float(it.original_coeff)
                        except Exception:
                            pass
            logger.info("Сброс глобальных коэффициентов: восстановлены коэффициенты подрядчиков и позиций")
            # Пересчитываем отображение
            self.recalculate_all()
        except Exception:
            logger.error("Ошибка при сбросе коэффициентов: %s", traceback.format_exc())

    # 16. Обработчики «Наша скидка…»
    def _on_our_discount_pct_changed(self, vendor: str, value: float) -> None:
        # Если задан %, сбрасываем сумму (взаимоисключающие поля)
        self.preview_our_discount_pct[vendor] = float(value) if value > 0.0 else None
        self.preview_our_discount_sum[vendor] = None
        logger.info("Предпросмотр: Наша скидка %% для '%s' → %.2f (сумма сброшена)", vendor, value)
        self.recalculate_all()

    def _on_our_discount_sum_changed(self, vendor: str, value: float) -> None:
        # Если задана сумма, сбрасываем %
        self.preview_our_discount_sum[vendor] = float(value) if value > 0.0 else None
        self.preview_our_discount_pct[vendor] = None
        logger.info("Предпросмотр: Наша скидка ₽ для '%s' → %.2f (процент сброшен)", vendor, value)
        self.recalculate_all()

    # 17. Обработчики таблиц «Доходы из сметы» и «Расходы»
    def on_add_profit_clicked(self) -> None:
        try:
            # Диалог выбора позиций из сметы
            # Передаём также информацию о том, какие глобальные коэффициенты включены,
            # чтобы корректно рассчитать цену за единицу в диалоге выбора.
            dlg = FinanceTab.ProfitSelectDialog(self.items, self.preview_vendor_coeffs, self.preview_coeff_enabled, self)
            if dlg.exec() == QtWidgets.QDialog.Accepted:
                sel = dlg.get_selected()
                for vendor, name, amount in sel:
                    self.profit_items.append(ProfitItem(vendor=vendor, description=name, amount=float(amount)))
                logger.info("Добавлено позиций дохода из сметы: %d", len(sel))
                self.recalculate_all()
        except Exception:
            logger.error("Ошибка добавления доходов из сметы: %s", traceback.format_exc())
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось добавить доход из сметы. См. логи.")

    def on_remove_profit_clicked(self) -> None:
        rows = sorted({idx.row() for idx in self.tbl_profit.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(self.profit_items):
                self.profit_items.pop(r)
        if rows:
            logger.info("Удалено позиций дохода из сметы: %d", len(rows))
            self.recalculate_all()

    def on_add_exp_clicked(self) -> None:
        self.expense_items.append(ExpenseItem(name="", qty=1.0, price=0.0))
        self.recalculate_all()

    def on_remove_exp_clicked(self) -> None:
        rows = sorted({idx.row() for idx in self.tbl_expense.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(self.expense_items):
                self.expense_items.pop(r)
        if rows:
            logger.info("Удалено расходов: %d", len(rows))
            self.recalculate_all()

    # 17.b. Обработчик добавления ручного дохода
    def on_add_manual_profit_clicked(self) -> None:
        """
        Добавляет новый доход вручную. Пользователю предлагается ввести
        подрядчика (можно оставить пустым), описание и сумму. Запись
        добавляется в список ``profit_items``, после чего выполняется
        пересчёт таблиц и сумм. Ошибки отображаются через диалог.
        """
        try:
            # Создаём диалог ввода
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Добавить доход")
            form = QtWidgets.QFormLayout(dlg)
            edit_vendor = QtWidgets.QLineEdit()
            edit_desc = QtWidgets.QLineEdit()
            spin_amount = QtWidgets.QDoubleSpinBox()
            spin_amount.setRange(-1_000_000_000.0, 1_000_000_000.0)
            spin_amount.setDecimals(2)
            spin_amount.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
            form.addRow("Подрядчик:", edit_vendor)
            form.addRow("Описание:", edit_desc)
            form.addRow("Сумма ₽:", spin_amount)
            btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            form.addRow(btn_box)
            btn_box.accepted.connect(dlg.accept)
            btn_box.rejected.connect(dlg.reject)
            if dlg.exec() == QtWidgets.QDialog.Accepted:
                vendor = edit_vendor.text().strip()
                desc = edit_desc.text().strip()
                amt = float(spin_amount.value())
                # Добавляем новую позицию дохода
                self.profit_items.append(ProfitItem(vendor=vendor, description=desc, amount=amt))
                logger.info("Добавлен ручной доход: vendor='%s', desc='%s', amount=%.2f", vendor, desc, amt)
                self.recalculate_all()
        except Exception:
            logger.error("Ошибка ручного добавления дохода", exc_info=True)
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось добавить доход. См. логи.")

    def _on_expense_qty_changed(self, row: int, value: float) -> None:
        if 0 <= row < len(self.expense_items):
            self.expense_items[row].qty = float(value)
            self.recalculate_all()

    def _on_expense_price_changed(self, row: int, value: float) -> None:
        if 0 <= row < len(self.expense_items):
            self.expense_items[row].price = float(value)
            self.recalculate_all()

    def _on_expense_name_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        """
        Обновляет имя расхода в модели при изменении текста в таблице.

        Таблица расходов позволяет редактировать название расхода (столбец 0). Когда
        пользователь вводит или изменяет текст, этот обработчик обновляет
        соответствующий объект ExpenseItem, чтобы имя сохранялось при
        последующем сохранении данных.
        """
        try:
            row = item.row()
            col = item.column()
            # Обрабатываем только столбец 0 («Расход»)
            if col == 0 and 0 <= row < len(self.expense_items):
                self.expense_items[row].name = item.text()
        except Exception:
            logger.error("Ошибка обработки изменения названия расхода", exc_info=True)

    # 17.c. Обработчик изменения суммы оплачено
    def _on_vendor_paid_changed(self, vendor: str, value: float) -> None:
        """
        Обновляет предпросмотрную сумму оплаченных средств для указанного
        подрядчика. После изменения значения выполняется пересчёт всех
        таблиц, чтобы отразить обновившийся остаток задолженности.

        :param vendor: имя подрядчика
        :param value: введённая сумма
        """
        try:
            self.preview_paid[vendor] = float(value)
            logger.info("Предпросмотр: выплата подрядчику '%s' → %.2f", vendor, value)
            self.recalculate_all()
        except Exception:
            logger.error("Ошибка обработки изменения суммы оплачено", exc_info=True)

    # 18. Сохранение изменений
    def on_save_clicked(self) -> None:
        try:
            # 18.1 Применяем предпросмотр в реальные настройки
            for v in self.vendors_settings.keys():
                s = self.vendors_settings[v]
                # Сохраняем коэффициент, введённый пользователем, игнорируя состояние предпросмотра (глобальный coeff)
                s.coeff = float(self._coeff_user_values.get(v, s.coeff))
                s.discount_pct = float(self.preview_discount_pct.get(v, s.discount_pct))
                s.commission_pct = float(self.preview_commission_pct.get(v, s.commission_pct))
                s.tax_pct = float(self.preview_tax_pct.get(v, s.tax_pct))
                s.our_discount_pct = self.preview_our_discount_pct.get(v, s.our_discount_pct)
                s.our_discount_sum = self.preview_our_discount_sum.get(v, s.our_discount_sum)
                # Запоминаем состояние включения глобального коэффициента
                s.coeff_enabled = bool(self.preview_coeff_enabled.get(v, False))
                # Сохраняем сумму оплачено из предпросмотра
                try:
                    s.paid = float(self.preview_paid.get(v, s.paid))
                except Exception:
                    s.paid = 0.0

            # 18.2 Применяем глобальные коэффициенты подрядчиков к позициям класса equipment
            self._apply_vendor_coefficients_to_items()

            # 18.3 Сохраняем всё через провайдер
            self.provider.save_items(self.items)
            self.provider.save_finance(self.vendors_settings, self.profit_items, self.expense_items)

            # 18.4 Пересчёт после сохранения (для актуализации предпросмотра)
            self.preview_vendor_coeffs = {v: s.coeff for v, s in self.vendors_settings.items()}
            self.preview_discount_pct = {v: s.discount_pct for v, s in self.vendors_settings.items()}
            self.preview_commission_pct = {v: s.commission_pct for v, s in self.vendors_settings.items()}
            self.preview_tax_pct = {v: s.tax_pct for v, s in self.vendors_settings.items()}
            self.preview_our_discount_pct = {v: s.our_discount_pct for v, s in self.vendors_settings.items()}
            self.preview_our_discount_sum = {v: s.our_discount_sum for v, s in self.vendors_settings.items()}
            # После сохранения возвращаем отображаемое значение коэффициента и активируем его
            self._coeff_user_values = {v: s.coeff for v, s in self.vendors_settings.items()}
            # После сохранения восстанавливаем состояние включения глобального коэффициента из настроек
            self.preview_coeff_enabled = {v: s.coeff_enabled for v, s in self.vendors_settings.items()}

            # Обновляем предпросмотр сумм оплат после сохранения
            self.preview_paid = {v: float(s.paid) for v, s in self.vendors_settings.items()}

            self.recalculate_all()

            logger.info("Изменения сохранены успешно.")
            QtWidgets.QMessageBox.information(self, "Готово", "Изменения сохранены.")
            self.saved.emit()

            # Излучаем сигнал finance_changed для уведомления ProjectPage
            try:
                page_obj = getattr(self, "_page", None)
                if page_obj is not None and hasattr(page_obj, "finance_changed"):
                    page_obj.finance_changed.emit()
            except Exception:
                logger.error("Ошибка отправки сигнала finance_changed: %s", traceback.format_exc())


        except Exception:
            logger.error("Ошибка сохранения: %s", traceback.format_exc())
            QtWidgets.QMessageBox.critical(self, "Ошибка", "Не удалось сохранить изменения. Подробности в логах.")

    # 19. Применение глобальных коэффициентов к позициям (только класс equipment)
    def _apply_vendor_coefficients_to_items(self) -> None:
        """
        Применяет глобальные коэффициенты подрядчиков к позициям класса equipment.
        Если коэффициент у подрядчика отключён (coeff_enabled == False), то значение
        coeff для позиций не перезаписывается. При повторном включении глобального
        коэффициента восстанавливается оригинальное значение, если оно было
        сохранено ранее. Также если глобальный коэффициент равен 1.0, восстанавливаем
        исходный коэффициент позиции.
        """
        for it in self.items:
            v = it.vendor or "(без подрядчика)"
            if it.cls == "equipment" and v in self.vendors_settings:
                settings = self.vendors_settings[v]
                # Проверяем, нужно ли применять глобальный коэффициент
                if settings.coeff_enabled:
                    # Сохраняем оригинальный коэффициент при первом замещении
                    if it.original_coeff is None:
                        it.original_coeff = it.coeff
                    # Замещаем coeff глобальным значением
                    it.coeff = float(settings.coeff)
                else:
                    # Глобальный коэффициент отключён — восстановим оригинальный при наличии
                    if it.original_coeff is not None:
                        it.coeff = float(it.original_coeff)
                        it.original_coeff = None
            # Для классов, отличных от equipment, ничего не делаем

        # Восстановление, если глобальный коэффициент стал 1.0 и ранее менялся
        # Когда коэффициент = 1.0 и был применён, отменяем переопределение
        for it in self.items:
            v = it.vendor or "(без подрядчика)"
            if it.cls == "equipment" and v in self.vendors_settings:
                settings = self.vendors_settings[v]
                if settings.coeff_enabled and settings.coeff == 1.0:
                    if it.original_coeff is not None:
                        it.coeff = float(it.original_coeff)
                        it.original_coeff = None

    # 20. Диалог выбора позиций для доходов из сметы
    class SimpleItemsModel(QtCore.QAbstractTableModel):
        """Простейшая модель для отображения позиций сметы в диалоге выбора."""
        def __init__(self, rows: List[Tuple[str, str, str, float]], parent=None) -> None:
            super().__init__(parent)
            self._rows = rows
            self._headers = ["Подрядчик", "Класс", "Наименование", "Сумма ₽"]

        def rowCount(self, parent=QtCore.QModelIndex()) -> int:
            return len(self._rows)

        def columnCount(self, parent=QtCore.QModelIndex()) -> int:
            return 4

        def data(self, index, role=QtCore.Qt.DisplayRole):
            if not index.isValid():
                return None
            r, c = index.row(), index.column()
            if role == QtCore.Qt.DisplayRole:
                val = self._rows[r][c]
                if c == 3:
                    return f"{round2(float(val)):,.2f}".replace(",", " ")
                return val
            return None

        def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
            if role == QtCore.Qt.DisplayRole and orientation == QtCore.Qt.Horizontal:
                return self._headers[section]
            return None

        def get_row(self, r: int) -> Tuple[str, str, str, float]:
            return self._rows[r]

    class ProfitSelectDialog(QtWidgets.QDialog):
        def __init__(self, items: List[Item], preview_vendor_coeffs: Dict[str, float], preview_coeff_enabled: Optional[Dict[str, bool]] = None, parent=None) -> None:
            super().__init__(parent)
            self.setWindowTitle("Выбор позиций для дохода")
            self.resize(1000, 600)
            layout = QtWidgets.QVBoxLayout(self)

            # 1. Формируем данные для модели: (vendor, class_ru, name, qty_available, unit_price, eff)
            #    В список добавляем также коэффициент `eff` (фактический коэффициент,
            #    с которым считается цена позиции). Значение `eff` показывает,
            #    как была получена цена за единицу (price * eff). Это поле не
            #    отображается в основной таблице, но используется в диалоге
            #    выбора количества для отображения столбца «Коэффициент».
            data_rows: List[Tuple[str, str, str, float, float, float]] = []
            unique_vendors: set[str] = set()
            unique_classes: set[str] = set()
            for it in items:
                # Рассчитываем коэффициент подрядчика (vendor_eff) для цены. При предпросмотре
                # он учитывается только если включён, иначе берём коэффициент 1.0.
                vendor_eff = 1.0
                if it.cls == "equipment" and it.vendor in preview_vendor_coeffs:
                    if preview_coeff_enabled is not None:
                        # Коэффициент vendor_eff активен только если флажок включён
                        if preview_coeff_enabled.get(it.vendor, False):
                            vendor_eff = preview_vendor_coeffs[it.vendor]
                    else:
                        vendor_eff = preview_vendor_coeffs[it.vendor]
                cls_ru = CLASS_EN2RU.get(it.cls, it.cls)
                # Цена за единицу зависит только от коэффициента подрядчика
                unit_amount = float(it.price) * float(vendor_eff)
                qty_avail = float(it.qty)
                # Используем коэффициент позиции (it.coeff) как ограничение для редактируемого поля
                eff_val = float(it.coeff)
                data_rows.append((it.vendor, cls_ru, it.name, qty_avail, unit_amount, eff_val))
                unique_vendors.add(it.vendor)
                unique_classes.add(cls_ru)

            # 2. Определяем модель данных: только отображение (vendor, class_ru, name, qty_available, unit_price).
            class ProfitItemsModel(QtCore.QAbstractTableModel):
                """
                Табличная модель для списка позиций сметы. Колонки:
                  0 – Подрядчик
                  1 – Класс (перевод на русский)
                  2 – Наименование
                  3 – Доступное количество
                  4 – Цена за единицу (с учётом глобального коэффициента)

                Модель не допускает редактирование ячеек; количество
                выбирается в отдельном диалоге при подтверждении выбора.
                """
                headers = ["Подрядчик", "Класс", "Наименование", "Доступно", "Цена ₽"]
                def __init__(self, rows: List[Tuple[str, str, str, float, float, float]], parent=None) -> None:
                    super().__init__(parent)
                    # Каждая строка: (vendor, class_ru, name, qty_available, unit_price, eff)
                    self._rows = rows
                def rowCount(self, parent=QtCore.QModelIndex()) -> int:
                    return len(self._rows)
                def columnCount(self, parent=QtCore.QModelIndex()) -> int:
                    # Отображаем только пять колонок: подрядчик, класс, наименование, доступно, цена
                    return 5
                def data(self, index: QtCore.QModelIndex, role: int = QtCore.Qt.DisplayRole):
                    if not index.isValid():
                        return None
                    r, c = index.row(), index.column()
                    # Распаковываем все поля; коэффициент eff (6‑е значение) в данном методе не нужен
                    vendor, cls_ru, name, qty_avail, unit_price, _eff = self._rows[r]
                    if role == QtCore.Qt.DisplayRole:
                        if c == 0:
                            return vendor
                        if c == 1:
                            return cls_ru
                        if c == 2:
                            return name
                        if c == 3:
                            # Показать целочисленное количество или дробное без лишних нулей
                            s = f"{qty_avail:.2f}".rstrip("0").rstrip(".")
                            return s
                        if c == 4:
                            return f"{unit_price:.2f}".replace(".", ",").rstrip("0").rstrip(",")
                    return None
                def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):
                    if role == QtCore.Qt.DisplayRole and orientation == QtCore.Qt.Horizontal:
                        return self.headers[section]
                    return None
                def flags(self, index: QtCore.QModelIndex) -> QtCore.Qt.ItemFlags:
                    if not index.isValid():
                        return QtCore.Qt.NoItemFlags
                    # Позволяем выделять строки, но не редактировать
                    return QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled
                def get_available(self, row: int) -> float:
                    """Возвращает доступное количество для выбранной строки."""
                    if 0 <= row < len(self._rows):
                        return self._rows[row][3]
                    return 0.0
                def get_unit_price(self, row: int) -> float:
                    """Возвращает цену за единицу для выбранной строки (с учётом коэффициента)."""
                    if 0 <= row < len(self._rows):
                        return self._rows[row][4]
                    return 0.0

                def get_eff(self, row: int) -> float:
                    """Возвращает коэффициент, с которым рассчитана цена за единицу."""
                    if 0 <= row < len(self._rows):
                        return self._rows[row][5]
                    return 1.0
                def get_vendor_name(self, row: int) -> Tuple[str, str]:
                    """Возвращает (vendor, name) для выбранной строки."""
                    if 0 <= row < len(self._rows):
                        return self._rows[row][0], self._rows[row][2]
                    return "", ""

            self.model = ProfitItemsModel(data_rows, self)

            # 3. Фильтр-модель для поиска/фильтрации
            class FilterProxyModel(QtCore.QSortFilterProxyModel):
                def __init__(self, parent=None) -> None:
                    super().__init__(parent)
                    self.search_text = ""
                    self.vendor_filter = "Все"
                    self.class_filter = "Все"
                def set_search(self, text: str) -> None:
                    self.search_text = text.lower().strip()
                    self.invalidateFilter()
                def set_vendor_filter(self, vendor: str) -> None:
                    self.vendor_filter = vendor
                    self.invalidateFilter()
                def set_class_filter(self, cls: str) -> None:
                    self.class_filter = cls
                    self.invalidateFilter()
                def filterAcceptsRow(self, source_row: int, parent: QtCore.QModelIndex) -> bool:
                    vendor = self.sourceModel().index(source_row, 0, parent).data(QtCore.Qt.DisplayRole)
                    cls_ru = self.sourceModel().index(source_row, 1, parent).data(QtCore.Qt.DisplayRole)
                    name = self.sourceModel().index(source_row, 2, parent).data(QtCore.Qt.DisplayRole)
                    if self.vendor_filter != "Все" and vendor != self.vendor_filter:
                        return False
                    if self.class_filter != "Все" and cls_ru != self.class_filter:
                        return False
                    if self.search_text:
                        text = f"{vendor} {cls_ru} {name}".lower()
                        return self.search_text in text
                    return True

            self.proxy = FilterProxyModel(self)
            self.proxy.setSourceModel(self.model)

            # 4. Панель фильтров и поиска
            controls = QtWidgets.QHBoxLayout()
            vendor_lbl = QtWidgets.QLabel("Подрядчик:")
            self.vendor_combo = QtWidgets.QComboBox()
            self.vendor_combo.addItem("Все")
            for vn in sorted(unique_vendors):
                self.vendor_combo.addItem(vn)
            self.vendor_combo.currentTextChanged.connect(self.proxy.set_vendor_filter)
            class_lbl = QtWidgets.QLabel("Класс:")
            self.class_combo = QtWidgets.QComboBox()
            self.class_combo.addItem("Все")
            for cl in sorted(unique_classes):
                self.class_combo.addItem(cl)
            self.class_combo.currentTextChanged.connect(self.proxy.set_class_filter)
            search_lbl = QtWidgets.QLabel("Поиск:")
            self.search_edit = QtWidgets.QLineEdit()
            self.search_edit.setPlaceholderText("Введите текст для поиска")
            self.search_edit.textChanged.connect(self.proxy.set_search)
            controls.addWidget(vendor_lbl)
            controls.addWidget(self.vendor_combo)
            controls.addWidget(class_lbl)
            controls.addWidget(self.class_combo)
            controls.addWidget(search_lbl)
            controls.addWidget(self.search_edit, 1)
            layout.addLayout(controls, 0)

            # 5. Таблица выбора
            self.view = QtWidgets.QTableView(self)
            self.view.setModel(self.proxy)
            # Разрешаем множественный выбор строк; пользователь отмечает позиции, которые будут добавлены.
            # Количество будет введено на следующем шаге (в отдельном диалоге).
            self.view.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
            self.view.setSelectionMode(QtWidgets.QTableView.MultiSelection)
            self.view.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
            layout.addWidget(self.view, 1)

            # 6. Кнопки OK/Cancel
            buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            layout.addWidget(buttons, 0)
            buttons.accepted.connect(self.accept)
            buttons.rejected.connect(self.reject)

        def get_selected(self) -> List[Tuple[str, str, float]]:
            """
            Возвращает список выбранных позиций (vendor, name, amount).

            Пользователь выбирает одну или несколько строк в таблице и нажимает OK.
            После этого открывается дополнительный диалог, где для каждой выбранной
            позиции нужно указать количество (целое число) с ограничением по
            доступному количеству. Если пользователь отменяет ввод или не
            вводит положительное количество, позиция не добавляется.
            """
            res: List[Tuple[str, str, float]] = []
            try:
                # 1. Индексы строк из исходной модели, которые выбрал пользователь
                selected_source_rows = [self.proxy.mapToSource(idx).row() for idx in self.view.selectionModel().selectedRows()]
                # Ничего не выбрано – возвращаем пустой список
                if not selected_source_rows:
                    return res
                # 2. Создаём диалог выбора количеств
                dlg = QtWidgets.QDialog(self)
                dlg.setWindowTitle("Количество позиций для дохода")
                dlg_layout = QtWidgets.QVBoxLayout(dlg)
                # Таблица с позициями, количеством, коэффициентом, скидкой и суммой
                # Колонки: подрядчик, наименование, доступно, кол-во, цена, коэффициент, срезаем, сумма
                table = QtWidgets.QTableWidget(len(selected_source_rows), 8, dlg)
                table.setHorizontalHeaderLabels([
                    "Подрядчик", "Наименование", "Доступно", "Кол-во", "Цена ₽", "Коэфф.", "Срезаем ₽", "Сумма ₽"])  # type: ignore[list-item]

                # Словари для хранения текущих значений количества и скидки
                current_qty: Dict[int, int] = {}
                current_cut: Dict[int, float] = {}
                # Текущие коэффициенты по каждой строке (позволяют
                # пользователю уменьшать коэффициент позиции, но не
                # увеличивать его выше исходного. Цена за единицу не
                # изменяется при изменении коэффициента; коэффициент
                # применяется лишь при расчёте дохода.)
                current_coeff: Dict[int, float] = {}
                # Сохраняем исходные цены за единицу для каждой строки (цены
                # не меняются при изменении коэффициента)
                base_prices: Dict[int, float] = {}
                # Суммы по каждой строке
                current_amounts: Dict[int, float] = {}
                # Метка для отображения общей суммы всех выбранных позиций
                sum_label = QtWidgets.QLabel()

                # Вспомогательная функция пересчёта суммы по строке и общего итога
                def recalc_row_total(r: int) -> None:
                    """Пересчитывает итог для строки r на основе текущего количества и скидки.

                    Правила расчёта:
                        - если скидка указана (>0), берём эту скидку как цену за единицу;
                        - если скидка = 0, берём исходную цену позиции;
                        - далее цена за единицу умножается на коэффициент
                          (который пользователь может уменьшить).
                    """
                    try:
                        qty_val = current_qty.get(r, 0)
                        cut_val = current_cut.get(r, 0.0)
                        # Исходная цена за единицу (не зависит от коэффициента)
                        base_price_val = base_prices.get(r, 0.0)
                        coeff_val = current_coeff.get(r, 0.0)
                        # Выбираем цену за единицу: скидку или полную цену
                        price_component = float(cut_val) if cut_val > 0.0 else base_price_val
                        per_unit_amount = coeff_val * price_component
                        new_amount = float(qty_val) * per_unit_amount
                        current_amounts[r] = new_amount
                        # Обновляем отображение суммы в таблице
                        item = table.item(r, 7)
                        if item is not None:
                            disp = f"{new_amount:.2f}".replace(".", ",").rstrip("0").rstrip(",")
                            item.setText(disp)
                        # Обновляем агрегированную сумму
                        total = sum(current_amounts.values())
                        sum_label.setText(f"Сумма к добавлению: {total:,.2f} ₽".replace(",", " "))
                    except Exception:
                        logger.error("Ошибка пересчёта суммы по строке", exc_info=True)

                # Создаём строки таблицы
                for row_idx, src_row in enumerate(selected_source_rows):
                    vendor, name = self.model.get_vendor_name(src_row)
                    qty_avail = self.model.get_available(src_row)
                    unit_price = self.model.get_unit_price(src_row)
                    eff_val = self.model.get_eff(src_row)
                    # Подрядчик
                    item0 = QtWidgets.QTableWidgetItem(vendor)
                    item0.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    table.setItem(row_idx, 0, item0)
                    # Наименование
                    item1 = QtWidgets.QTableWidgetItem(name)
                    item1.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    table.setItem(row_idx, 1, item1)
                    # Доступно (форматируем без лишних нулей)
                    disp_avail = f"{qty_avail:.2f}".rstrip("0").rstrip(".")
                    item2 = QtWidgets.QTableWidgetItem(disp_avail)
                    item2.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    table.setItem(row_idx, 2, item2)
                    # Количество (редактируемое) – QSpinBox
                    spin_qty = QtWidgets.QSpinBox()
                    spin_qty.setMinimum(0)
                    spin_qty.setMaximum(int(qty_avail))
                    # По умолчанию 1 (если есть доступные) или 0
                    default_qty = 1 if qty_avail >= 1 else 0
                    spin_qty.setValue(default_qty)
                    table.setCellWidget(row_idx, 3, spin_qty)
                    current_qty[row_idx] = default_qty
                    # Цена за единицу (не редактируемая)
                    disp_price = f"{unit_price:.2f}".replace(".", ",").rstrip("0").rstrip(",")
                    item4 = QtWidgets.QTableWidgetItem(disp_price)
                    item4.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    table.setItem(row_idx, 4, item4)
                    # Сохраняем исходную цену без учёта коэффициента (цена за единицу)
                    base_prices[row_idx] = unit_price
                    # Коэффициент – редактируемый DoubleSpinBox с ограничением сверху исходного coeff
                    spin_coeff = QtWidgets.QDoubleSpinBox()
                    spin_coeff.setMinimum(0.0)
                    spin_coeff.setMaximum(eff_val)
                    spin_coeff.setDecimals(2)
                    spin_coeff.setSingleStep(0.1)
                    spin_coeff.setValue(eff_val)
                    spin_coeff.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
                    spin_coeff.setKeyboardTracking(False)
                    table.setCellWidget(row_idx, 5, spin_coeff)
                    # Сохраняем текущий коэффициент
                    current_coeff[row_idx] = eff_val
                    # Срезаем с цены – DoubleSpinBox для ввода абсолютной скидки
                    spin_cut = QtWidgets.QDoubleSpinBox()
                    spin_cut.setMinimum(0.0)
                    # Максимальная скидка на единицу не может превышать цену за единицу
                    spin_cut.setMaximum(unit_price)
                    spin_cut.setDecimals(2)
                    spin_cut.setSingleStep(1.0)
                    spin_cut.setValue(0.0)
                    # Убираем стрелки для чистоты интерфейса
                    spin_cut.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
                    spin_cut.setKeyboardTracking(False)
                    table.setCellWidget(row_idx, 6, spin_cut)
                    current_cut[row_idx] = 0.0
                    # Итоговая сумма для текущей строки: изначально 0
                    current_amounts[row_idx] = 0.0
                    disp_amount = f"0".replace(".", ",")
                    item7 = QtWidgets.QTableWidgetItem(disp_amount)
                    item7.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    table.setItem(row_idx, 7, item7)

                    # Обработчики изменений количества и скидки
                    def on_qty_changed(val: int, r=row_idx) -> None:
                        # При изменении количества обновляем текущий qty и пересчитываем строку
                        current_qty[r] = val
                        recalc_row_total(r)
                    spin_qty.valueChanged.connect(on_qty_changed)

                    def on_cut_changed(val: float, r=row_idx) -> None:
                        # При изменении скидки обновляем текущую скидку и пересчитываем строку
                        current_cut[r] = float(val)
                        recalc_row_total(r)
                    spin_cut.valueChanged.connect(on_cut_changed)

                    def on_coeff_changed(val: float, r=row_idx) -> None:
                        # При изменении коэффициента обновляем сохранённый коэффициент и пересчитываем строку
                        try:
                            new_coeff = float(val)
                            current_coeff[r] = new_coeff
                            recalc_row_total(r)
                        except Exception:
                            logger.error("Ошибка изменения коэффициента", exc_info=True)
                    spin_coeff.valueChanged.connect(on_coeff_changed)

                # Автоматическое растягивание колонок
                table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
                dlg_layout.addWidget(table, 1)
                # Итоговая сумма всех выбранных позиций
                total_amount = sum(current_amounts.values())
                sum_label.setText(f"Сумма к добавлению: {total_amount:,.2f} ₽".replace(",", " "))
                dlg_layout.addWidget(sum_label, 0)
                # Кнопки OK/Cancel
                btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
                dlg_layout.addWidget(btn_box)
                btn_box.accepted.connect(dlg.accept)
                btn_box.rejected.connect(dlg.reject)
                # Показываем диалог и формируем результат
                if dlg.exec() == QtWidgets.QDialog.Accepted:
                    for row_idx, src_row in enumerate(selected_source_rows):
                        qty = current_qty.get(row_idx, 0)
                        cut_per_unit = current_cut.get(row_idx, 0.0)
                        if qty > 0:
                            vendor, name = self.model.get_vendor_name(src_row)
                            # Получаем исходную цену за единицу для строки
                            base_price_val = base_prices.get(row_idx, self.model.get_unit_price(src_row))
                            # Получаем выбранный коэффициент для строки
                            coeff_val = current_coeff.get(row_idx, 0.0)
                            # Выбираем цену за единицу: скидку или полную цену
                            price_component = cut_per_unit if cut_per_unit > 0.0 else base_price_val
                            per_unit = coeff_val * price_component
                            amount = float(qty) * float(per_unit)
                            res.append((vendor, name, amount))
                return res
            except Exception:
                # В случае ошибки возвращаем пустой список и записываем в лог
                logger.error("ProfitSelectDialog.get_selected: %s", traceback.format_exc())
                return res

# 21. Wrapper function for backward compatibility
def build_finance_tab(page: object, tab: QtWidgets.QWidget) -> None:
    """Строит вкладку «Бухгалтерия» для ProjectPage.

    В отличие от версии по умолчанию, если у страницы есть доступ к
    базе данных (page.db) и задан project_id, используется
    DBDataProvider: позиции и финансовые настройки загружаются прямо
    из базы и сохраняются обратно в projects.finance_json и items.

    Если база отсутствует, используется FileDataProvider (по умолчанию).

    Виджет FinanceTab сохраняется в атрибуте page.tab_finance_widget.
    """
    try:
        # Определяем подходящий провайдер: из базы или файловый
        provider = None
        try:
            if getattr(page, "db", None) and getattr(page, "project_id", None):
                provider = DBDataProvider(page)
                logger.info("build_finance_tab: использую DBDataProvider для project_id=%s", page.project_id)
        except Exception:
            logger.error("build_finance_tab: ошибка определения DBDataProvider: %s", traceback.format_exc())
            provider = None
        # Создаём виджет вкладки с провайдером
        finance_widget = FinanceTab(data_provider=provider)
        # Передаём ссылку на страницу в виджет (если нужно для логики)
        setattr(finance_widget, "_page", page)
        # Размещаем виджет на вкладке
        layout = QtWidgets.QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(finance_widget)
        # сохраняем ссылку на виджет для дальнейшего доступа
        setattr(page, "tab_finance_widget", finance_widget)
        logger.info("Вкладка 'Бухгалтерия' построена: использован %s", type(provider).__name__ if provider else "FileDataProvider")
        # Попытаемся связать сигнал сохранения с обновлением сводных показателей во вкладке «Информация»
        try:
            # Подключаем сигнал сохранения к обновлению сводных показателей и сметы
            from .info_tab import update_financial_summary  # type: ignore
            from .summary_tab import reload_zone_tabs  # type: ignore
            if hasattr(finance_widget, "saved"):
                # При сохранении обновляем «Информацию» и перестраиваем «Сводную смету»
                finance_widget.saved.connect(lambda: update_financial_summary(page))
                finance_widget.saved.connect(lambda: reload_zone_tabs(page))
                logger.info("Сигнал сохранения 'Бухгалтерии' привязан к обновлению информации и сводной сметы")

            # Подключаем обновление «Бухгалтерии» к сигналу изменения сводной сметы
            try:
                if hasattr(page, "summary_changed") and hasattr(finance_widget, "on_summary_changed"):
                    page.summary_changed.connect(finance_widget.on_summary_changed)
                    logger.info("Сигнал summary_changed подключён к обновлению вкладки 'Бухгалтерия'")
            except Exception:
                logger.warning("Не удалось связать сигнал summary_changed", exc_info=True)
        except Exception:
            # Неудача связывания не является критичной, поэтому просто логируем
            logger.warning("Не удалось привязать сигнал сохранения к обновлению сводных показателей и сметы", exc_info=True)
        # 21.1 Определяем функцию пересчёта для связи со «Сводной сметой»
        #
        # В старой версии (ACT10) метод recalc_finance находился в модуле
        # finance_tab и добавлялся в объект страницы. Эта функция перезагружала
        # данные из базы и обновляла таблицы бухгалтерии. Для обеспечения
        # совместимости и синхронизации со сводной сметой, здесь создаём
        # аналогичную функцию, которая будет доступна через page.recalc_finance.
        def _recalc_finance_from_summary() -> None:
            """Перезагружает данные для вкладки «Бухгалтерия» из базы и пересчитывает её.

            Вызывается из вкладки «Сводная смета» после любого изменения.
            Загружает список позиций из базы проекта, передаёт их в
            finance_widget и инициирует пересчёт всех таблиц.
            Если база недоступна, использует провайдер виджета.
            """
            try:
                items: List[Item] = []
                # Если доступна база и выбран проект — читаем из базы
                if getattr(page, "db", None) and getattr(page, "project_id", None):
                    try:
                        rows = page.db.list_items(page.project_id)
                        # Используем класс Item, определённый в этом модуле
                        for row in rows:
                            try:
                                item_id = str(row["id"]) if "id" in row.keys() else ""
                                vendor = ""
                                if "vendor" in row.keys() and row["vendor"] is not None:
                                    vendor = str(row["vendor"]).strip()
                                vendor = vendor or "(без подрядчика)"
                                cls = "equipment"
                                if "type" in row.keys() and row["type"]:
                                    cls = row["type"]
                                department = ""
                                if "department" in row.keys() and row["department"]:
                                    department = row["department"]
                                zone = ""
                                if "zone" in row.keys() and row["zone"]:
                                    zone = row["zone"]
                                name = ""
                                if "name" in row.keys() and row["name"]:
                                    name = row["name"]
                                unit_price = 0.0
                                if "unit_price" in row.keys() and row["unit_price"] is not None:
                                    unit_price = float(row["unit_price"])
                                qty = 0.0
                                if "qty" in row.keys() and row["qty"] is not None:
                                    qty = float(row["qty"])
                                coeff = 1.0
                                if "coeff" in row.keys() and row["coeff"] is not None:
                                    coeff = float(row["coeff"])
                                items.append(Item(
                                    id=item_id,
                                    vendor=vendor,
                                    cls=cls,
                                    department=department,
                                    zone=zone,
                                    name=name,
                                    price=unit_price,
                                    qty=qty,
                                    coeff=coeff,
                                ))
                            except Exception:
                                logger.error("_recalc_finance_from_summary: ошибка обработки строки: %s", traceback.format_exc())
                                continue
                    except Exception:
                        items = []
                else:
                    # Иначе используем текущий провайдер виджета
                    prov = getattr(finance_widget, "provider", None)
                    if prov is not None:
                        try:
                            items = prov.load_items() or []
                        except Exception:
                            items = []
                # Передаём позиции в виджет
                try:
                    finance_widget.set_items(items)
                    # Полный пересчёт
                    if hasattr(finance_widget, "recalculate_all"):
                        finance_widget.recalculate_all()
                except Exception:
                    pass
                # Обновляем вкладку «Информация»
                try:
                    from .info_tab import update_financial_summary  # type: ignore
                    update_financial_summary(page)
                except Exception:
                    pass
            except Exception:
                logger.error("Ошибка пересчёта в _recalc_finance_from_summary", exc_info=True)

        # Добавляем функцию в объект страницы под именем recalc_finance
        setattr(page, "recalc_finance", _recalc_finance_from_summary)
        # Выполним пересчёт один раз после создания вкладки
        try:
            _recalc_finance_from_summary()
        except Exception:
            pass
    except Exception:
        logger.error("Ошибка при построении вкладки 'Бухгалтерия': %s", traceback.format_exc())
        QtWidgets.QMessageBox.critical(tab, "Ошибка", "Не удалось построить вкладку 'Бухгалтерия'. Подробности в логах.")


# 22. Точка входа для отладки виджета отдельно
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    # Пример запуска с файловым провайдером по умолчанию (project_root = .., project_id = debug)
    w = FinanceTab(project_id="debug")
    w.resize(1280, 800)
    w.setWindowTitle("Бухгалтерия — предпросмотр")
    w.show()
    sys.exit(app.exec())
