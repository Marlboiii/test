import csv
import datetime as dt
import json
import tempfile
import unittest
from pathlib import Path

import openpyxl

from ohouse_orders.cli import main
from ohouse_orders.reader import OrderFileError, read_orders
from ohouse_orders.store import OrderStore

HEADERS = ["주문번호", "주문상품번호", "주문일시", "주문상태", "상품명", "옵션", "수량",
           "결제금액", "주문자", "수령인", "수령인 연락처", "우편번호", "주소", "배송메모", "사은품"]


def make_xlsx(path, rows, title=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    if title:
        ws.append(["오늘의집 주문내역 (2026-10-07 다운로드)"])
        ws.append([])
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    wb.save(path)


ROW_A = [202610070001, 1, dt.datetime(2026, 10, 7, 9, 30), "결제완료", "원목 협탁", "내추럴 / 대형",
         2, "59,800원", "홍길동", "김철수", 1012345678, 4524, "서울 중구 세종대로 1", "문 앞", "쿠션"]
ROW_B = ["202610070002", "1", "2026.10.06 18:05", "배송준비중", "러그", "", "1",
         "32000", "이영희", "이영희", "010-9876-5432", "06236", "서울 강남구 테헤란로 2", "", ""]


class ReaderTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_xlsx_normalizes_fields(self):
        path = self.dir / "orders.xlsx"
        make_xlsx(path, [ROW_A, ROW_B, [None] * len(HEADERS)])
        orders = read_orders(path)
        self.assertEqual(len(orders), 2)
        a = orders[0]
        self.assertEqual(a["order_no"], "202610070001")
        self.assertEqual(a["item_no"], "1")
        self.assertEqual(a["ordered_at"], "2026-10-07 09:30:00")
        self.assertEqual(a["amount"], 59800)
        self.assertEqual(a["quantity"], 2)
        self.assertEqual(a["receiver_phone"], "010-1234-5678")
        self.assertEqual(a["zipcode"], "04524")
        self.assertEqual(a["extra"], {"사은품": "쿠션"})
        self.assertEqual(orders[1]["ordered_at"], "2026-10-06 18:05:00")

    def test_csv_cp949(self):
        path = self.dir / "orders.csv"
        with open(path, "w", newline="", encoding="cp949") as f:
            w = csv.writer(f)
            w.writerow(HEADERS)
            w.writerow(ROW_B)
        orders = read_orders(path)
        self.assertEqual(orders[0]["product_name"], "러그")
        self.assertEqual(orders[0]["amount"], 32000)

    def test_custom_mapping(self):
        path = self.dir / "orders.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append(["주문번호", "받는사람 핸드폰"])
        wb.active.append(["X1", "01011112222"])
        wb.save(path)
        mapping = self.dir / "mapping.json"
        mapping.write_text(json.dumps({"receiver_phone": ["받는사람 핸드폰"]}), encoding="utf-8")
        orders = read_orders(path, mapping_path=mapping)
        self.assertEqual(orders[0]["receiver_phone"], "010-1111-2222")

    def test_missing_header(self):
        path = self.dir / "bad.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append(["a", "b"])
        wb.save(path)
        with self.assertRaises(OrderFileError):
            read_orders(path)


class StoreAndCliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.db = str(self.dir / "orders.db")

    def tearDown(self):
        self.tmp.cleanup()

    def test_reimport_updates_status_without_duplicates(self):
        first = self.dir / "day1.xlsx"
        make_xlsx(first, [ROW_A, ROW_B])
        second = self.dir / "day2.xlsx"
        updated = list(ROW_A)
        updated[3] = "배송중"
        make_xlsx(second, [updated, ROW_B])

        with OrderStore(self.db) as store:
            self.assertEqual(store.upsert(read_orders(first)), (2, 0, 0))
            self.assertEqual(store.upsert(read_orders(second)), (0, 1, 1))
            rows = store.query()
            self.assertEqual(len(rows), 2)
            self.assertEqual({r["status"] for r in rows}, {"배송중", "배송준비중"})
            self.assertEqual(len(store.query(status=["배송중"])), 1)
            self.assertEqual(len(store.query(since="2026-10-07")), 1)
            self.assertEqual(len(store.query(until="2026-10-06")), 1)

    def test_cli_import_and_export(self):
        src = self.dir / "downloads"
        src.mkdir()
        make_xlsx(src / "a.xlsx", [ROW_A, ROW_B])
        out = self.dir / "out.xlsx"
        self.assertEqual(main(["--db", self.db, "import", str(src)]), 0)
        self.assertEqual(main(["--db", self.db, "export", "-o", str(out), "--status", "결제완료", "--extra"]), 0)
        ws = openpyxl.load_workbook(out).active
        header = [c.value for c in ws[1]]
        self.assertEqual(header[0], "주문번호")
        self.assertIn("사은품", header)
        self.assertEqual(ws.max_row, 2)
        self.assertEqual(ws.cell(2, 1).value, "202610070001")

        out_csv = self.dir / "out.csv"
        self.assertEqual(main(["--db", self.db, "export", "-o", str(out_csv)]), 0)
        with open(out_csv, encoding="utf-8-sig") as f:
            self.assertEqual(len(list(csv.reader(f))), 3)


if __name__ == "__main__":
    unittest.main()
