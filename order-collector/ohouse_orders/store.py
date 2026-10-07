"""수집한 주문을 SQLite에 누적 저장한다 (같은 주문은 상태/송장만 갱신)."""

import datetime as dt
import json
import sqlite3

from .columns import FIELDS, INT_FIELDS

def _sql_type(field):
    return "INTEGER" if field in INT_FIELDS else "TEXT"


SCHEMA = f"""
CREATE TABLE IF NOT EXISTS orders (
    order_key TEXT PRIMARY KEY,
    {", ".join(f"{f} {_sql_type(f)}" for f in FIELDS)},
    extra TEXT,
    source_file TEXT,
    first_collected_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""

INDEXES = """
CREATE INDEX IF NOT EXISTS idx_orders_ordered_at ON orders(ordered_at);
CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
"""

COLUMNS = FIELDS + ["extra", "source_file", "first_collected_at", "updated_at"]


def order_key(record):
    """주문 1줄의 고유키. 주문옵션번호가 있으면 그것을, 없으면 주문상품번호+상품명+옵션을 쓴다."""
    if record.get("item_no"):
        return f"{record['order_no']}#{record['item_no']}"
    return "#".join(
        [record["order_no"]]
        + [record.get(f) or "" for f in ("product_order_no", "product_name", "option")]
    )


def _now():
    return dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")


class OrderStore:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
        self._migrate()
        self.conn.executescript(INDEXES)

    def _migrate(self):
        """이전 버전 DB에 없는 컬럼을 추가한다."""
        existing = {row["name"] for row in self.conn.execute("PRAGMA table_info(orders)")}
        with self.conn:
            for field in FIELDS:
                if field not in existing:
                    self.conn.execute(f"ALTER TABLE orders ADD COLUMN {field} {_sql_type(field)}")

    def close(self):
        self.conn.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def upsert(self, records):
        """레코드를 저장하고 (신규 수, 변경 수, 동일 수)를 반환한다."""
        new = changed = same = 0
        now = _now()
        with self.conn:
            for record in records:
                key = order_key(record)
                values = {f: record.get(f) for f in FIELDS}
                values["extra"] = json.dumps(record.get("extra") or {}, ensure_ascii=False)
                values["source_file"] = record.get("source_file")
                existing = self.conn.execute(
                    "SELECT * FROM orders WHERE order_key = ?", (key,)
                ).fetchone()
                if existing is None:
                    values.update(order_key=key, first_collected_at=now, updated_at=now)
                    cols = ", ".join(values)
                    marks = ", ".join("?" for _ in values)
                    self.conn.execute(
                        f"INSERT INTO orders ({cols}) VALUES ({marks})", list(values.values())
                    )
                    new += 1
                    continue
                # 새 파일에 비어 있는 값으로 기존 값을 지우지 않는다
                updates = {
                    k: v for k, v in values.items()
                    if v not in (None, "", "{}") and v != existing[k]
                }
                # 파일명만 다른 것은 변경으로 보지 않는다
                if not set(updates) - {"source_file"}:
                    same += 1
                    continue
                updates["updated_at"] = now
                assignments = ", ".join(f"{k} = ?" for k in updates)
                self.conn.execute(
                    f"UPDATE orders SET {assignments} WHERE order_key = ?",
                    [*updates.values(), key],
                )
                changed += 1
        return new, changed, same

    def query(self, status=None, since=None, until=None, collected_since=None):
        sql = "SELECT * FROM orders WHERE 1=1"
        params = []
        if status:
            sql += f" AND status IN ({', '.join('?' for _ in status)})"
            params += list(status)
        if since:
            sql += " AND ordered_at >= ?"
            params.append(since)
        if until:
            # 날짜만 주면 그날 하루 전체 포함
            sql += " AND ordered_at < ?"
            params.append(_next_day(until) if len(until) == 10 else until)
        if collected_since:
            sql += " AND first_collected_at >= ?"
            params.append(collected_since)
        sql += " ORDER BY ordered_at, order_no, product_order_no, item_no"
        return [dict(row) for row in self.conn.execute(sql, params)]

    def status_counts(self):
        rows = self.conn.execute(
            "SELECT COALESCE(NULLIF(status, ''), '(없음)') AS status, COUNT(*) AS n "
            "FROM orders GROUP BY 1 ORDER BY n DESC"
        )
        return [(row["status"], row["n"]) for row in rows]


def _next_day(date_text):
    return (dt.date.fromisoformat(date_text) + dt.timedelta(days=1)).isoformat()
