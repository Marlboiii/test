import csv
import datetime as dt
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import openpyxl

from ohouse_orders.cli import main
from ohouse_orders.reader import OrderFileError, read_orders
from ohouse_orders.store import OrderStore

# 오늘의집 파트너센터 주문 엑셀의 실제 헤더 (데이터는 가짜)
HEADERS = [
    "주문번호", "주문상품번호", "주문옵션번호", "상품아이디", "상품자체관리코드", "상품명",
    "옵션아이디", "옵션자체관리코드", "옵션명", "수량", "주문메모", "오늘출발여부", "조립신청여부",
    "주문결제완료일", "출고예정일", "배송예정일", "희망배송일", "빠른배송/오늘의집배송일", "구매확정일",
    "개인통관부호", "인증번호", "판매가", "판매가 * 수량", "조립비", "배송비",
    "판매가*수량 + 조립비 + 배송비", "정산예정금액", "주문자명", "주문자 연락처", "수취인명",
    "수취인 연락처", "수취인 우편번호", "수취인 주소", "수취인 주소상세", "배송방법",
    "배송정보(택배사)", "운송장번호", "제주/도서산간 여부", "제주/도서산간 배송비", "배송메모",
    "배송조회 URL", "묶음배송그룹", "주문상태",
]


def order_row(**values):
    defaults = {
        "주문번호": "100000001", "주문상품번호": "200000001", "주문옵션번호": "300000001",
        "상품아이디": "4000001", "상품명": "테스트 원목 침대", "옵션아이디": "5000001",
        "옵션명": "SS 프레임", "수량": "1", "주문결제완료일": "2026-10-04 17:00:59",
        "출고예정일": "2026-10-13", "판매가": "282000", "판매가 * 수량": "282000", "조립비": "0",
        "배송비": "40000", "판매가*수량 + 조립비 + 배송비": "322000", "정산예정금액": "245340",
        "주문자명": "홍길동", "주문자 연락처": "010-0000-0001", "수취인명": "홍길동",
        "수취인 연락처": "010-0000-0001", "수취인 우편번호": "04524",
        "수취인 주소": "서울특별시 중구 세종대로 1", "수취인 주소상세": "101동 101호",
        "배송방법": "업체 직접 배달", "제주/도서산간 배송비": "0", "배송메모": "배송 전 연락 주세요",
        "묶음배송그룹": "G1", "주문상태": "배송준비",
    }
    defaults.update(values)
    return [defaults.get(h, "") for h in HEADERS]


def make_xlsx(path, rows, headers=HEADERS):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


# 같은 주문상품번호 안에 옵션(추가상품)이 2줄
ROW_A = order_row()
ROW_A2 = order_row(**{"주문옵션번호": "300000002", "옵션아이디": "5000002",
                      "옵션명": "추가상품 - 매트리스SS", "판매가": "170000", "판매가 * 수량": "170000",
                      "판매가*수량 + 조립비 + 배송비": "210000", "정산예정금액": "147900"})
ROW_B = order_row(**{"주문번호": "100000002", "주문상품번호": "200000002", "주문옵션번호": "300000003",
                     "상품명": "러그", "옵션명": "", "수량": "2", "주문결제완료일": "2026-10-06 18:05:00",
                     "판매가": "16000", "판매가 * 수량": "32000", "배송비": "0",
                     "판매가*수량 + 조립비 + 배송비": "32000", "정산예정금액": "28000",
                     "배송정보(택배사)": "CJ대한통운", "운송장번호": "123456789012",
                     "주문상태": "배송중"})


class ReaderTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_ohouse_export_columns(self):
        path = self.dir / "orders.xlsx"
        make_xlsx(path, [ROW_A, ROW_A2, ROW_B, [None] * len(HEADERS)])
        orders = read_orders(path)
        self.assertEqual(len(orders), 3)
        a = orders[0]
        self.assertEqual((a["order_no"], a["product_order_no"], a["item_no"]),
                         ("100000001", "200000001", "300000001"))
        self.assertEqual(a["ordered_at"], "2026-10-04 17:00:59")
        self.assertEqual(a["ship_due_date"], "2026-10-13")
        self.assertEqual(a["option"], "SS 프레임")
        self.assertEqual((a["unit_price"], a["amount"], a["shipping_fee"], a["total_amount"],
                          a["settlement_amount"]), (282000, 282000, 40000, 322000, 245340))
        self.assertEqual(a["receiver_name"], "홍길동")
        self.assertEqual(a["receiver_phone"], "010-0000-0001")
        self.assertEqual(a["zipcode"], "04524")
        self.assertEqual(a["address_detail"], "101동 101호")
        self.assertEqual(a["delivery_method"], "업체 직접 배달")
        self.assertEqual(a["status"], "배송준비")
        self.assertEqual(a["extra"], {"조립비": "0", "제주/도서산간 배송비": "0", "묶음배송그룹": "G1"})
        b = orders[2]
        self.assertEqual((b["courier"], b["tracking_no"], b["quantity"]), ("CJ대한통운", "123456789012", 2))

    def test_numbers_and_title_rows(self):
        # 엑셀이 숫자로 바꾼 값, 제목 행, 다른 양식의 헤더도 처리
        path = self.dir / "other.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["주문내역 다운로드"])
        ws.append([])
        ws.append(["주문번호", "주문일시", "결제금액", "수령인 연락처", "우편번호"])
        ws.append([202610070001, dt.datetime(2026, 10, 7, 9, 30), "59,800원", 1012345678, 4524])
        wb.save(path)
        o = read_orders(path)[0]
        self.assertEqual(o["order_no"], "202610070001")
        self.assertEqual(o["ordered_at"], "2026-10-07 09:30:00")
        self.assertEqual(o["total_amount"], 59800)
        self.assertEqual(o["receiver_phone"], "010-1234-5678")
        self.assertEqual(o["zipcode"], "04524")

    def test_csv_cp949(self):
        path = self.dir / "orders.csv"
        with open(path, "w", newline="", encoding="cp949") as f:
            w = csv.writer(f)
            w.writerow(HEADERS)
            w.writerow(ROW_B)
        orders = read_orders(path)
        self.assertEqual(orders[0]["product_name"], "러그")
        self.assertEqual(orders[0]["amount"], 32000)
        self.assertEqual(orders[0]["courier"], "CJ대한통운")

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
        make_xlsx(first, [ROW_A, ROW_A2, ROW_B])
        second = self.dir / "day2.xlsx"
        shipped = order_row(**{"주문상태": "배송중", "배송정보(택배사)": "직접배송"})
        make_xlsx(second, [shipped, ROW_A2, ROW_B])

        with OrderStore(self.db) as store:
            # 같은 주문상품번호의 옵션 2줄이 따로 저장되어야 한다
            self.assertEqual(store.upsert(read_orders(first)), (3, 0, 0))
            self.assertEqual(store.upsert(read_orders(second)), (0, 1, 2))
            rows = store.query()
            self.assertEqual(len(rows), 3)
            self.assertEqual([r["status"] for r in rows], ["배송중", "배송준비", "배송중"])
            self.assertEqual(len(store.query(status=["배송준비"])), 1)
            self.assertEqual(len(store.query(since="2026-10-05")), 1)
            self.assertEqual(len(store.query(until="2026-10-04")), 2)

    def test_migrates_old_db(self):
        conn = sqlite3.connect(self.db)
        conn.execute("CREATE TABLE orders (order_key TEXT PRIMARY KEY, order_no TEXT, extra TEXT, "
                     "source_file TEXT, first_collected_at TEXT NOT NULL, updated_at TEXT NOT NULL)")
        conn.commit()
        conn.close()
        path = self.dir / "a.xlsx"
        make_xlsx(path, [ROW_A])
        with OrderStore(self.db) as store:
            self.assertEqual(store.upsert(read_orders(path)), (1, 0, 0))
            self.assertEqual(store.query()[0]["settlement_amount"], 245340)

    def test_cli_import_and_export(self):
        src = self.dir / "downloads"
        src.mkdir()
        make_xlsx(src / "a.xlsx", [ROW_A, ROW_B])
        out = self.dir / "out.xlsx"
        self.assertEqual(main(["--db", self.db, "import", str(src)]), 0)
        self.assertEqual(main(["--db", self.db, "export", "-o", str(out), "--status", "배송준비", "--extra"]), 0)
        ws = openpyxl.load_workbook(out).active
        header = [c.value for c in ws[1]]
        self.assertEqual(header[0], "주문번호")
        self.assertIn("묶음배송그룹", header)
        self.assertEqual(ws.max_row, 2)
        self.assertEqual(ws.cell(2, 1).value, "100000001")

        out_csv = self.dir / "out.csv"
        self.assertEqual(main(["--db", self.db, "export", "-o", str(out_csv)]), 0)
        with open(out_csv, encoding="utf-8-sig") as f:
            self.assertEqual(len(list(csv.reader(f))), 3)


if __name__ == "__main__":
    unittest.main()
