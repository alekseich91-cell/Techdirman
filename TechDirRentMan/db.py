"""
Назначение файла:
    Слой доступа к данным (SQLite) и схема базы приложения TechDirRentMan.

Принцип работы (кратко):
    - Подключение к SQLite (WAL + foreign_keys).
    - Безопасная инициализация/миграция схемы: CREATE TABLE → ensure_column (ALTER) → ensure_index.
    - Таблицы:
        * projects   — проекты;
        * items      — позиции проекта (vendor/department/zone/power_watts/import_batch);
        * catalog    — глобальная база (name, unit_price, class, vendor, power_watts, department).
    - Методы:
        * Проекты/позиции: CRUD, выборки с фильтрами, суммирование, откат импорта по batch,
          массовые вставки, точечные и множественные апдейты полей.
        * Каталог: импорт/экспорт CSV, доп. выборки, смена класса/мощности, дубли, средняя/максимальная цена/мощность.
        * Новое: project_sync_from_catalog(project_id) — обновляет в проекте «type» (класс) и «power_watts»
          по наименованиям в соответствии с базой (класс = наиболее частый, мощность = максимум).
          Используется для синхронизации сводной сметы после правок в базе.

Стиль:
    - Код разбит на пронумерованные секции; у ключевых операций краткие комментарии.
"""

# 1. Импорт
import sqlite3
from pathlib import Path
from typing import Iterable, Any, Optional, Dict, List, Tuple
import datetime
import csv
import logging


# 2. Класс DB
class DB:
    # 2.1 Конструктор
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self._conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA journal_mode=WAL;")
        self._conn.execute("PRAGMA foreign_keys=ON;")

    # 2.2 Инициализация схемы (безопасный порядок)
    def init_schema(self):
        cur = self._conn.cursor()
        # 2.2.1 Базовые таблицы и индексы, не зависящие от новых столбцов
        cur.executescript(
            """
            -- Проекты
            CREATE TABLE IF NOT EXISTS projects(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            );

            -- Позиции проекта
            CREATE TABLE IF NOT EXISTS items(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
                type TEXT NOT NULL DEFAULT 'equipment',
                group_name TEXT NOT NULL DEFAULT 'Аренда оборудования',
                name TEXT NOT NULL,
                qty REAL NOT NULL DEFAULT 1,
                coeff REAL NOT NULL DEFAULT 1,
                amount REAL NOT NULL DEFAULT 0,
                unit_price REAL NOT NULL DEFAULT 0,
                source_file TEXT,
                created_at TEXT NOT NULL
                -- Новые поля добавляются далее через ALTER (см. ensure_column)
            );
            CREATE INDEX IF NOT EXISTS idx_items_project ON items(project_id);
            CREATE INDEX IF NOT EXISTS idx_items_name ON items(name);

            -- Глобальная база (каталог)
            CREATE TABLE IF NOT EXISTS catalog(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                unit_price REAL NOT NULL DEFAULT 0,
                class TEXT NOT NULL DEFAULT 'equipment',
                vendor TEXT,
                power_watts REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                department TEXT
            );
            -- Уникальность: (name, vendor, unit_price)
            CREATE UNIQUE INDEX IF NOT EXISTS uq_catalog_name_vendor_price
                ON catalog(name, COALESCE(vendor,''), unit_price);

            -- Утилитарные индексы
            CREATE INDEX IF NOT EXISTS idx_catalog_name ON catalog(name);
            CREATE INDEX IF NOT EXISTS idx_catalog_class ON catalog(class);
            CREATE INDEX IF NOT EXISTS idx_catalog_vendor ON catalog(vendor);
            CREATE INDEX IF NOT EXISTS idx_catalog_department ON catalog(department);
            """
        )
        self._conn.commit()

        # 2.2.2 Добавляем недостающие столбцы к items
        self._ensure_column("items", "vendor",       "ALTER TABLE items ADD COLUMN vendor TEXT;")
        self._ensure_column("items", "department",   "ALTER TABLE items ADD COLUMN department TEXT;")
        self._ensure_column("items", "zone",         "ALTER TABLE items ADD COLUMN zone TEXT;")
        self._ensure_column("items", "power_watts",  "ALTER TABLE items ADD COLUMN power_watts REAL NOT NULL DEFAULT 0;")
        self._ensure_column("items", "import_batch", "ALTER TABLE items ADD COLUMN import_batch TEXT;")

        # 2.2.2a Добавляем недостающий столбец тайминга к таблице проектов
        #    Используется для хранения конфигурации тайминга в формате JSON.
        self._ensure_column(
            "projects",
            "timing_json",
            "ALTER TABLE projects ADD COLUMN timing_json TEXT;"
        )

        # 2.2.2b Добавляем недостающий столбец финансовых данных к таблице проектов
        #     Используется для хранения конфигурации вкладки «Бухгалтерия» в формате JSON.
        self._ensure_column(
            "projects",
            "finance_json",
            "ALTER TABLE projects ADD COLUMN finance_json TEXT;"
        )

        # 2.2.3 Индексы по новым столбцам (создаём только после ensure_column)
        self._ensure_index("idx_items_batch", "CREATE INDEX IF NOT EXISTS idx_items_batch ON items(import_batch);")
        self._ensure_index("idx_items_vendor", "CREATE INDEX IF NOT EXISTS idx_items_vendor ON items(vendor);")
        self._ensure_index("idx_items_department", "CREATE INDEX IF NOT EXISTS idx_items_department ON items(department);")
        self._ensure_index("idx_items_zone", "CREATE INDEX IF NOT EXISTS idx_items_zone ON items(zone);")
        self._ensure_index("idx_items_type", "CREATE INDEX IF NOT EXISTS idx_items_type ON items(type);")

        # 2.2.4 Расширение глобального каталога: добавляем колонку для учёта складских остатков
        #    Колонка stock_qty содержит количество оборудования, находящееся на складе у подрядчика.
        #    Значение REAL выбрано для поддержки дробных единиц. По умолчанию 0.
        #    ensure_column безопасно добавит её только один раз.
        self._ensure_column(
            "catalog",
            "stock_qty",
            "ALTER TABLE catalog ADD COLUMN stock_qty REAL NOT NULL DEFAULT 0;"
        )

    # 2.3 Вспомогательные: обеспечение столбцов и индексов
    def _ensure_column(self, table: str, column: str, ddl: str):
        cur = self._conn.cursor()
        cur.execute(f"PRAGMA table_info({table})")
        cols = [r[1] for r in cur.fetchall()]
        if column not in cols:
            cur.execute(ddl)
            self._conn.commit()

    def _ensure_index(self, name: str, ddl: str):
        cur = self._conn.cursor()
        cur.execute("PRAGMA index_list(items)")
        names = [r[1] for r in cur.fetchall()]
        if name not in names:
            cur.execute(ddl)
            self._conn.commit()

    # 2.4 ------- Методы ПРОЕКТОВ -------
    def add_project(self, name: str) -> int:
        now = datetime.datetime.utcnow().isoformat()
        cur = self._conn.cursor()
        cur.execute("INSERT INTO projects(name, created_at) VALUES(?, ?)", (name, now))
        self._conn.commit()
        return cur.lastrowid

    def list_projects(self):
        cur = self._conn.cursor()
        cur.execute("SELECT * FROM projects ORDER BY created_at DESC")
        return cur.fetchall()

    def delete_project(self, project_id: int):
        self._conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
        self._conn.commit()

    # 2.4.5 Переименование проекта
    def rename_project(self, project_id: int, new_name: str) -> None:
        """
        Переименовывает проект в таблице projects.

        Принимает идентификатор проекта и новое имя. Если в базе уже
        существует запись с таким именем, SQLite поднимет исключение
        IntegrityError из-за ограничения UNIQUE. Вызывающему коду
        рекомендуется перехватывать это исключение, чтобы показать
        пользователю сообщение об ошибке. После обновления выполняется commit.

        :param project_id: id существующего проекта
        :param new_name: новое название
        """
        cur = self._conn.cursor()
        cur.execute(
            "UPDATE projects SET name=? WHERE id=?",
            (new_name, project_id),
        )
        self._conn.commit()

    # 2.4.1 Получение JSON тайминга проекта
    def get_project_timing(self, project_id: int) -> Optional[str]:
        """Возвращает строку JSON с таймингом проекта.

        Если поле timing_json отсутствует или равно NULL, возвращает None.
        """
        cur = self._conn.cursor()
        cur.execute("SELECT timing_json FROM projects WHERE id=?", (project_id,))
        row = cur.fetchone()
        if row is None:
            return None
        return row[0]

    # 2.4.2 Сохранение JSON тайминга проекта
    def set_project_timing(self, project_id: int, timing_json: str) -> None:
        """Сохраняет строку JSON с таймингом в таблицу projects.

        Использует простой UPDATE. После обновления выполняет commit.
        """
        cur = self._conn.cursor()
        cur.execute(
            "UPDATE projects SET timing_json=? WHERE id=?",
            (timing_json, project_id),
        )
        self._conn.commit()

    # 2.4.3 Получение JSON финансов проекта
    def get_project_finance(self, project_id: int) -> Optional[str]:
        """Возвращает строку JSON с конфигурацией вкладки «Бухгалтерия».

        Если поле finance_json отсутствует или равно NULL, возвращает None.
        """
        cur = self._conn.cursor()
        cur.execute("SELECT finance_json FROM projects WHERE id=?", (project_id,))
        row = cur.fetchone()
        if row is None:
            return None
        return row[0]

    # 2.4.4 Сохранение JSON финансов проекта
    def set_project_finance(self, project_id: int, finance_json: str) -> None:
        """Сохраняет строку JSON с конфигурацией вкладки «Бухгалтерия» в таблицу projects.

        Использует UPDATE; после обновления выполняет commit.
        """
        cur = self._conn.cursor()
        cur.execute(
            "UPDATE projects SET finance_json=? WHERE id=?",
            (finance_json, project_id),
        )
        self._conn.commit()

    def add_items_bulk(self, items: Iterable[dict]):
        """
        Ожидаемые ключи в словаре item:
        project_id, type, group_name, name, qty, coeff, amount, unit_price, source_file,
        vendor, department, zone, power_watts, import_batch
        """
        cur = self._conn.cursor()
        now = datetime.datetime.utcnow().isoformat()
        cur.executemany(
            """
            INSERT INTO items(project_id, type, group_name, name, qty, coeff, amount, unit_price,
                              source_file, created_at, vendor, department, zone, power_watts, import_batch)
            VALUES(:project_id, :type, :group_name, :name, :qty, :coeff, :amount, :unit_price,
                   :source_file, :created_at, :vendor, :department, :zone, :power_watts, :import_batch)
            """,
            [
                {
                    "project_id": it["project_id"],
                    "type": it.get("type", "equipment"),
                    "group_name": it.get("group_name", "Аренда оборудования"),
                    "name": it["name"],
                    "qty": float(it.get("qty", 1) or 1),
                    "coeff": float(it.get("coeff", 1) or 1),
                    "amount": float(it.get("amount", 0) or 0),
                    "unit_price": float(it.get("unit_price", 0) or 0),
                    "source_file": it.get("source_file"),
                    "created_at": it.get("created_at") or now,
                    "vendor": (it.get("vendor") or "").strip(),
                    "department": (it.get("department") or "").strip(),
                    "zone": (it.get("zone") or "").strip(),
                    "power_watts": float(it.get("power_watts", 0) or 0),
                    "import_batch": it.get("import_batch"),
                }
                for it in items
            ],
        )
        self._conn.commit()

    def list_items(self, project_id: int):
        cur = self._conn.cursor()
        cur.execute("SELECT * FROM items WHERE project_id=? ORDER BY id", (project_id,))
        return cur.fetchall()

    def list_items_filtered(
        self,
        project_id: int,
        vendor: Optional[str] = None,
        department: Optional[str] = None,
        zone: Optional[str] = None,
        class_en: Optional[str] = None,
        name_like: Optional[str] = None,
    ):
        """
        Возвращает позиции проекта с опциональными фильтрами.
        Пустая строка в zone означает 'Без зоны'.
        """
        sql = "SELECT * FROM items WHERE project_id=?"
        args: List[Any] = [project_id]
        # Фильтр подрядчика: сравниваем без учёта регистра
        if vendor and vendor != "<ALL>":
            sql += " AND LOWER(COALESCE(vendor,'')) = LOWER(?)"; args.append(vendor)
        # Фильтр отдела: сравниваем без учёта регистра
        if department and department != "<ALL>":
            sql += " AND LOWER(COALESCE(department,'')) = LOWER(?)"; args.append(department)
        if zone is not None:
            # если zone == "<ALL>" — без фильтра; иначе сравниваем без учёта регистра
            if zone != "<ALL>":
                sql += " AND LOWER(COALESCE(zone,'')) = LOWER(?)"; args.append(zone)
        if class_en and class_en != "<ALL>":
            sql += " AND type = ?"; args.append(class_en)
        # Если задана строка для поиска по наименованию — используем поиск без учёта регистра.
        # В SQLite оператор LIKE по умолчанию может быть чувствителен к регистру в зависимости от сборки,
        # поэтому явно задаём COLLATE NOCASE. Это обеспечивает корректный поиск для кириллицы и латиницы.
        if name_like:
            sql += " AND name LIKE ? COLLATE NOCASE"; args.append(f"%{name_like}%")
        sql += " ORDER BY name COLLATE NOCASE"
        cur = self._conn.cursor(); cur.execute(sql, args)
        return cur.fetchall()

    def update_item_field(self, item_id: int, field: str, value: Any):
        assert field in {"type", "group_name", "name", "qty", "coeff", "amount", "unit_price", "vendor", "department", "zone", "power_watts"}
        self._conn.execute(f"UPDATE items SET {field}=? WHERE id=?", (value, item_id))
        self._conn.commit()

    def update_item_fields(self, item_id: int, fields: Dict[str, Any]):
        """
        Массовый апдейт нескольких полей (qty, amount, zone и т.д.).
        """
        allowed = {"type","group_name","name","qty","coeff","amount","unit_price","vendor","department","zone","power_watts"}
        pairs = [(k, v) for k, v in fields.items() if k in allowed]
        if not pairs:
            return
        set_sql = ", ".join([f"{k}=?" for k, _ in pairs])
        args = [v for _, v in pairs] + [item_id]
        self._conn.execute(f"UPDATE items SET {set_sql} WHERE id=?", args)
        self._conn.commit()

    def get_item_by_id(self, item_id: int) -> Optional[sqlite3.Row]:
        cur = self._conn.cursor()
        cur.execute("SELECT * FROM items WHERE id=?", (item_id,))
        return cur.fetchone()

    def delete_items(self, item_ids: Iterable[int]):
        self._conn.executemany("DELETE FROM items WHERE id=?", [(i,) for i in item_ids])
        self._conn.commit()

    def delete_items_by_import_batch(self, project_id: int, batch: str) -> int:
        # Откат последнего импорта по batch-id
        cur = self._conn.cursor()
        cur.execute("DELETE FROM items WHERE project_id=? AND import_batch=?", (project_id, batch))
        self._conn.commit()
        return cur.rowcount

    def project_total(self, project_id: int) -> float:
        cur = self._conn.cursor()
        cur.execute("SELECT COALESCE(SUM(amount),0) as total FROM items WHERE project_id=?", (project_id,))
        row = cur.fetchone()
        return float(row["total"] or 0)

    def project_distinct_values(self, project_id: int, field: str) -> List[str]:
        """
        Возвращает список уникальных значений для заданного поля внутри проекта.

        Все значения приводятся к нормализованному регистру: каждое слово в
        значении начинается с заглавной буквы, остальные буквы делаются
        строчными. Это позволяет устранить дубликаты, отличающиеся только
        регистром (например, «Отдел» и «отдел» будут считаться одним
        значением). Если поле пустое (NULL или пустая строка), результатом
        будет пустая строка, которая всегда помещается в начало списка.

        :param project_id: идентификатор проекта
        :param field: одно из {"vendor", "department", "zone"}
        :return: отсортированный список уникальных значений
        """
        assert field in {"vendor", "department", "zone"}
        cur = self._conn.cursor()
        cur.execute(
            f"""
            SELECT COALESCE({field},'')
            FROM items
            WHERE project_id=?
            """,
            (project_id,),
        )
        raw_values: List[str] = []
        for (val,) in cur.fetchall():
            # Обрабатываем None как пустую строку
            text = val or ""
            # Нормализуем: каждое слово с заглавной буквы, остальные строчные
            parts = text.split()
            normalized = " ".join(p.capitalize() for p in parts)
            raw_values.append(normalized)
        # Удаляем дубликаты без учёта регистра
        unique = []
        seen_lower = set()
        for val in raw_values:
            low = val.lower()
            if low not in seen_lower:
                unique.append(val)
                seen_lower.add(low)
        # Сортируем в алфавитном порядке без учёта регистра
        unique.sort(key=lambda s: s.lower())
        # Перемещаем пустое значение в начало
        if "" in unique:
            unique.remove("")
            unique.insert(0, "")
        return unique

    # 2.X Переименование зоны
    def rename_zone(self, project_id: int, old_name: Optional[str], new_name: Optional[str]) -> int:
        """
        Переименовывает зону проекта.
        Если old_name пустая строка/None — обновляются позиции без зоны (NULL/пустая).
        Если new_name пустая строка/None — позиции переводятся в состояние "без зоны".
        Возвращает количество затронутых позиций.
        """
        cur = self._conn.cursor()
        # Подготавливаем критерий выборки
        if not old_name:
            where = "project_id=? AND (zone IS NULL OR zone='')"
            params_sel = (project_id,)
        else:
            where = "project_id=? AND zone=?"
            params_sel = (project_id, old_name)
        try:
            # Считаем затронутые строки до апдейта
            cur.execute(f"SELECT COUNT(1) FROM items WHERE {where}", params_sel)
            count_before = int(cur.fetchone()[0])
            # Выполняем обновление
            new_val = new_name if (new_name is not None and new_name != '') else None
            if not old_name:
                cur.execute("UPDATE items SET zone=? WHERE project_id=? AND (zone IS NULL OR zone='')", (new_val, project_id))
            else:
                cur.execute("UPDATE items SET zone=? WHERE project_id=? AND zone=?", (new_val, project_id, old_name))
            self._conn.commit()
            try:
                logging.getLogger(__name__).info("Переименование зоны: '%s' -> '%s' (проект %s, затронуто позиций: %d)",
                                                 old_name if old_name else "<без зоны>",
                                                 new_name if new_name else "<без зоны>",
                                                 project_id, count_before)
            except Exception:
                pass
            return count_before
        except Exception as ex:
            logging.getLogger(__name__).error("Ошибка переименования зоны '%s' -> '%s' (проект %s): %s",
                                              old_name, new_name, project_id, ex, exc_info=True)
            raise

    def project_distinct_item_names(self, project_id: int) -> List[str]:
        """Уникальные наименования в проекте (для синхронизации)."""
        cur = self._conn.cursor()
        cur.execute("SELECT DISTINCT name FROM items WHERE project_id=? ORDER BY name COLLATE NOCASE", (project_id,))
        return [r[0] for r in cur.fetchall()]

    # 2.5 ------- Методы КАТАЛОГА -------
    def catalog_add_or_ignore(self, rows: Iterable[dict]):
        # Пачечная вставка в глобальную базу с игнорированием дублей по (name, vendor, unit_price)
        cur = self._conn.cursor()
        now = datetime.datetime.utcnow().isoformat()
        cur.executemany(
            """
            INSERT OR IGNORE INTO catalog(name, unit_price, class, vendor, power_watts, created_at, department)
            VALUES(:name, :unit_price, :class, :vendor, :power_watts, :created_at, :department)
            """,
            [
                {
                    "name": (r.get("name") or "").strip(),
                    "unit_price": float((r.get("unit_price", 0) or 0)),
                    "class": (r.get("class") or "equipment").strip(),
                    "vendor": (r.get("vendor") or "").strip(),
                    "power_watts": float((r.get("power_watts", 0) or 0)),
                    "created_at": r.get("created_at") or now,
                    "department": (r.get("department") or "").strip(),
                }
                for r in rows
                if r.get("name")
            ],
        )
        self._conn.commit()

    def catalog_import_csv(self, csv_path: Path) -> int:
        added = 0
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            batch = []
            for row in reader:
                batch.append({
                    "name": row.get("name", "").strip(),
                    "unit_price": row.get("unit_price", "0").replace(" ", "").replace(",", "."),
                    "class": row.get("class", "equipment").strip() or "equipment",
                    "vendor": row.get("vendor", "").strip(),
                    "power_watts": row.get("power_watts", "0").replace(" ", "").replace(",", "."),
                    "department": row.get("department", "").strip(),
                })
                if len(batch) >= 1000:
                    self.catalog_add_or_ignore(batch); added += len(batch); batch.clear()
            if batch:
                self.catalog_add_or_ignore(batch); added += len(batch)
        return added

    def catalog_export_csv(self, csv_path: Path, filters: Optional[Dict[str, Any]] = None) -> int:
        rows = self.catalog_list(filters or {})
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["name", "unit_price", "class", "vendor", "power_watts", "department", "created_at"])
            for r in rows:
                w.writerow([r["name"], r["unit_price"], r["class"], r["vendor"] or "", r["power_watts"] or 0, r["department"] or "", r["created_at"]])
        return len(rows)

    def catalog_distinct_values(self, field: str) -> List[str]:
        assert field in {"class", "vendor", "department"}
        cur = self._conn.cursor()
        cur.execute(f"SELECT DISTINCT {field} FROM catalog WHERE COALESCE({field},'')<>'' ORDER BY {field} COLLATE NOCASE")
        return [r[0] for r in cur.fetchall()]

    def catalog_list(self, filters: Dict[str, Any]) -> list[sqlite3.Row]:
        """
        Возвращает список строк каталога по заданным фильтрам.

        Поиск по наименованию выполняется без учёта регистра и ищет
        подстроку в поле ``name``. Если указан фильтр ``class``
        (не равный "<ALL>"), фильтр ``vendor`` или ``department``,
        то выборка ограничивается соответствующим значением.

        :param filters: словарь с ключами name, class, vendor, department
        :return: список строк ``sqlite3.Row``
        """
        name_like = (filters.get("name") or "").strip()
        class_eq = filters.get("class") or None
        vendor_eq = filters.get("vendor") or None
        department_eq = filters.get("department") or None

        sql = "SELECT * FROM catalog WHERE 1=1"
        args: List[Any] = []
        if name_like:
            # Для поиска по наименованию используем COLLATE NOCASE, чтобы
            # совпадения не зависели от регистра. Применяем подстановочный
            # поиск по подстроке.
            sql += " AND name LIKE ? COLLATE NOCASE"; args.append(f"%{name_like}%")
        if class_eq and class_eq != "<ALL>":
            sql += " AND class = ?"; args.append(class_eq)
        if vendor_eq and vendor_eq != "<ALL>":
            sql += " AND COALESCE(vendor,'') = ?"; args.append(vendor_eq)
        if department_eq and department_eq != "<ALL>":
            sql += " AND COALESCE(department,'') = ?"; args.append(department_eq)

        sql += " ORDER BY name COLLATE NOCASE, unit_price"
        cur = self._conn.cursor(); cur.execute(sql, args)
        return cur.fetchall()

    def catalog_update_field(self, row_id: int, field: str, value: Any):
        # Разрешаем менять класс и мощность
        assert field in {"class", "power_watts"}
        self._conn.execute(f"UPDATE catalog SET {field}=? WHERE id=?", (value, row_id))
        self._conn.commit()

    def catalog_bulk_update_class(self, ids: Iterable[int], new_class: str) -> int:
        """
        Массовая смена класса в каталоге.
        Возвращает число затронутых строк.
        """
        cur = self._conn.cursor()
        cur.executemany("UPDATE catalog SET class=? WHERE id=?", [(new_class, i) for i in ids])
        self._conn.commit()
        return cur.rowcount

    def catalog_find_duplicates(self) -> Dict[Tuple[str, str, float], List[int]]:
        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT name, COALESCE(vendor,''), unit_price, GROUP_CONCAT(id) AS ids, COUNT(*) AS cnt
            FROM catalog
            GROUP BY name, COALESCE(vendor,''), unit_price
            HAVING COUNT(*) > 1
            """
        )
        dups: Dict[Tuple[str, str, float], List[int]] = {}
        for name, vendor, price, ids_str, _cnt in cur.fetchall():
            ids = [int(x) for x in str(ids_str).split(",")]
            dups[(name, vendor, float(price))] = ids
        return dups

    def catalog_delete_ids(self, ids: Iterable[int]) -> int:
        cur = self._conn.cursor()
        cur.executemany("DELETE FROM catalog WHERE id=?", [(i,) for i in ids])
        self._conn.commit()
        return cur.rowcount

    def catalog_delete_duplicates(self) -> int:
        dups = self.catalog_find_duplicates()
        to_del: List[int] = []
        for _key, ids in dups.items():
            ids_sorted = sorted(ids)
            to_del.extend(ids_sorted[1:])  # удаляем все, кроме минимального id
        if not to_del:
            return 0
        return self.catalog_delete_ids(to_del)

    def catalog_get_class_by_name(self, name: str) -> Optional[str]:
        """
        Возвращает наиболее часто встречаемый класс (или первый попавшийся) для наименования из каталога.
        Если не найдено — None.
        """
        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT class, COUNT(*) as cnt
            FROM catalog
            WHERE name = ?
            GROUP BY class
            ORDER BY cnt DESC
            LIMIT 1
            """,
            (name.strip(),),
        )
        row = cur.fetchone()
        return row["class"] if row else None

    def catalog_avg_price_by_name(self, name: str) -> float:
        """Средняя цена по наименованию из каталога."""
        cur = self._conn.cursor()
        cur.execute("SELECT AVG(unit_price) FROM catalog WHERE name = ?", (name.strip(),))
        val = cur.fetchone()[0]
        return float(val or 0.0)

    def catalog_max_power_by_name(self, name: str) -> float:
        """Максимальная мощность по наименованию (любые подрядчики)."""
        cur = self._conn.cursor()
        cur.execute("SELECT MAX(power_watts) FROM catalog WHERE name = ?", (name.strip(),))
        val = cur.fetchone()[0]
        return float(val or 0.0)

    def catalog_distinct_powers_by_name_vendor(self, name: str, vendor: str) -> List[float]:
        """Набор уникальных значений мощности по (name, vendor)."""
        cur = self._conn.cursor()
        cur.execute(
            "SELECT DISTINCT power_watts FROM catalog WHERE name=? AND COALESCE(vendor,'')=?",
            (name.strip(), vendor.strip()),
        )
        return [float(r[0] or 0.0) for r in cur.fetchall()]

    def catalog_update_power_by_name_vendor(self, name: str, vendor: str, new_power_w: float) -> int:
        """Обновить мощность у всех записей каталога с данным (name, vendor)."""
        cur = self._conn.cursor()
        cur.execute(
            "UPDATE catalog SET power_watts=? WHERE name=? AND COALESCE(vendor,'')=?",
            (float(new_power_w or 0), name.strip(), vendor.strip()),
        )
        self._conn.commit()
        return cur.rowcount

    def catalog_update_stock_by_name_vendor(self, name: str, vendor: str, stock_qty: float) -> int:
        """
        Обновляет количество оборудования на складе (stock_qty) для записей каталога
        с указанными наименованием и подрядчиком.

        Этот метод устанавливает значение stock_qty для всех строк каталога,
        удовлетворяющих условию (name, vendor). Если подрядчик не указан (пустая строка),
        совпадение выполняется по полю vendor=NULL.

        :param name: наименование позиции (с учётом регистра и пробелов будет
                     приведено к строке и обрезано)
        :param vendor: имя подрядчика; пустая строка означает отсутствие подрядчика
        :param stock_qty: количество оборудования на складе; отрицательные значения
                          приводятся к 0
        :return: число обновлённых строк
        """
        # Корректируем входные данные: имя и подрядчик обрезаем, складское
        # количество приводим к float и не допускаем отрицательных значений.
        name_s = (name or "").strip()
        vendor_s = (vendor or "").strip()
        qty = float(stock_qty or 0)
        if qty < 0:
            qty = 0.0
        cur = self._conn.cursor()
        cur.execute(
            "UPDATE catalog SET stock_qty=? WHERE name=? AND COALESCE(vendor,'')=?",
            (qty, name_s, vendor_s),
        )
        self._conn.commit()
        return cur.rowcount

    # 2.6 Синхронизация проекта с каталогом (класс/мощность)
    def project_sync_from_catalog(self, project_id: int) -> Tuple[int, int]:
        """
        Обновляет в items: type и power_watts по совпадающим наименованиям из каталога.
        - type = наиболее часто встречаемый класс для name в каталоге;
        - power_watts = максимальная мощность для name в каталоге.
        Возвращает: (кол-во строк, у которых обновился класс; кол-во строк, у которых обновилась мощность).
        """
        names = self.project_distinct_item_names(project_id)
        if not names:
            return (0, 0)

        cur = self._conn.cursor()
        upd_class = 0
        upd_power = 0

        for nm in names:
            cls = self.catalog_get_class_by_name(nm)
            if cls:
                cur.execute(
                    "UPDATE items SET type=? WHERE project_id=? AND name=? AND type<>?",
                    (cls, project_id, nm, cls),
                )
                upd_class += cur.rowcount

            max_pw = self.catalog_max_power_by_name(nm)
            if max_pw and max_pw > 0:
                cur.execute(
                    "UPDATE items SET power_watts=? WHERE project_id=? AND name=? AND COALESCE(power_watts,0)<>?",
                    (max_pw, project_id, nm, max_pw),
                )
                upd_power += cur.rowcount

        self._conn.commit()
        return (upd_class, upd_power)

    # 2.7 Прочее
    def commit(self):
        self._conn.commit()

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass

    def delete_items_by_vendor_zone(self, project_id: int, vendor: str, zone: str) -> list[dict]:
        """
        Массовое удаление позиций по подрядчику и зоне с возвратом удалённых строк
        в виде словарей для возможного восстановления (undo).
        Пустая зона интерпретируется как '' (без зоны).
        """
        cur = self._conn.cursor()
        v = vendor or ""
        z = zone or ""
        # Читаем удаляемые строки для UNDO
        cur.execute(
            "SELECT * FROM items WHERE project_id=? AND COALESCE(vendor,'')=? AND COALESCE(zone,'')=?",
            (project_id, v, z)
        )
        rows = cur.fetchall()
        # Преобразуем в словари с только валидными ключами для повторной вставки
        to_restore = []
        for r in rows:
            d = dict(r)
            # Удаляем авто-поле id и временные поля, выставим заново при вставке
            d.pop("id", None)
            d["project_id"] = project_id
            d["vendor"] = d.get("vendor") or ""
            d["zone"] = d.get("zone") or ""
            # import_batch при восстановлении будет новым
            d["import_batch"] = None
            to_restore.append(d)
        # Удаляем
        cur.execute(
            "DELETE FROM items WHERE project_id=? AND COALESCE(vendor,'')=? AND COALESCE(zone,'')=?",
            (project_id, v, z)
        )
        self._conn.commit()
        return to_restore
