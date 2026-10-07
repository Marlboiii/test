"""명령행 인터페이스.

    python -m ohouse_orders import 다운로드/*.xlsx
    python -m ohouse_orders export -o 주문목록.xlsx --status 결제완료
    python -m ohouse_orders summary
"""

import argparse
import csv
import datetime as dt
import glob
import json
import os
import sys
from pathlib import Path

from .reader import OrderFileError, read_orders
from .store import OrderStore

DEFAULT_DB = "ohouse_orders.db"

# 내보내기 컬럼 순서와 한글 헤더
EXPORT_COLUMNS = [
    ("order_no", "주문번호"),
    ("product_order_no", "주문상품번호"),
    ("item_no", "주문옵션번호"),
    ("ordered_at", "주문결제완료일"),
    ("ship_due_date", "출고예정일"),
    ("status", "주문상태"),
    ("product_id", "상품아이디"),
    ("product_code", "상품자체관리코드"),
    ("product_name", "상품명"),
    ("option_id", "옵션아이디"),
    ("option_code", "옵션자체관리코드"),
    ("option", "옵션명"),
    ("quantity", "수량"),
    ("unit_price", "판매가"),
    ("amount", "판매가*수량"),
    ("shipping_fee", "배송비"),
    ("total_amount", "합계금액"),
    ("settlement_amount", "정산예정금액"),
    ("buyer_name", "주문자명"),
    ("buyer_phone", "주문자 연락처"),
    ("receiver_name", "수취인명"),
    ("receiver_phone", "수취인 연락처"),
    ("zipcode", "우편번호"),
    ("address", "주소"),
    ("address_detail", "상세주소"),
    ("delivery_memo", "배송메모"),
    ("delivery_method", "배송방법"),
    ("courier", "택배사"),
    ("tracking_no", "운송장번호"),
    ("first_collected_at", "최초수집일시"),
    ("updated_at", "최종갱신일시"),
    ("source_file", "원본파일"),
]


def _expand(patterns):
    files = []
    for pattern in patterns:
        matches = sorted(glob.glob(pattern)) or [pattern]
        for match in matches:
            path = Path(match)
            if path.is_dir():
                files += sorted(
                    p for p in path.iterdir()
                    if p.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv")
                    and not p.name.startswith("~$")
                )
            else:
                files.append(path)
    return files


def cmd_import(args):
    files = _expand(args.files)
    if not files:
        print("가져올 파일이 없습니다.", file=sys.stderr)
        return 1
    password = args.password or os.environ.get("OHOUSE_EXCEL_PASSWORD")
    total_new = total_changed = 0
    failed = 0
    with OrderStore(args.db) as store:
        for path in files:
            try:
                orders = read_orders(path, password=password, mapping_path=args.mapping)
            except (OrderFileError, FileNotFoundError) as exc:
                print(f"[실패] {exc}", file=sys.stderr)
                failed += 1
                continue
            new, changed, same = store.upsert(orders)
            total_new += new
            total_changed += changed
            print(f"[완료] {path.name}: {len(orders)}건 읽음 → 신규 {new}, 변경 {changed}, 동일 {same}")
    print(f"합계: 신규 {total_new}건, 변경 {total_changed}건 (DB: {args.db})")
    return 1 if failed and failed == len(files) else 0


def _write_xlsx(path, orders, include_extra):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    extra_keys = _extra_keys(orders) if include_extra else []
    headers = [h for _, h in EXPORT_COLUMNS] + extra_keys

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "주문목록"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
    for order in orders:
        extra = json.loads(order.get("extra") or "{}")
        ws.append([order.get(f) for f, _ in EXPORT_COLUMNS] + [extra.get(k, "") for k in extra_keys])
    # 주문번호/연락처 등이 숫자로 바뀌지 않도록 텍스트 서식 지정
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str):
                cell.number_format = "@"
    for i, header in enumerate(headers, start=1):
        width = max([len(str(header))] + [len(str(c.value or "")) for c in ws[get_column_letter(i)][1:200]])
        ws.column_dimensions[get_column_letter(i)].width = min(max(width * 1.4, 8), 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def _write_csv(path, orders, include_extra):
    extra_keys = _extra_keys(orders) if include_extra else []
    # 엑셀에서 한글이 깨지지 않도록 BOM 포함 UTF-8
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([h for _, h in EXPORT_COLUMNS] + extra_keys)
        for order in orders:
            extra = json.loads(order.get("extra") or "{}")
            writer.writerow(
                ["" if order.get(f) is None else order.get(f) for f, _ in EXPORT_COLUMNS]
                + [extra.get(k, "") for k in extra_keys]
            )


def _extra_keys(orders):
    keys = []
    for order in orders:
        for key in json.loads(order.get("extra") or "{}"):
            if key not in keys:
                keys.append(key)
    return keys


def cmd_export(args):
    with OrderStore(args.db) as store:
        orders = store.query(
            status=args.status,
            since=args.since,
            until=args.until,
            collected_since=args.collected_since,
        )
    out = args.output or f"오늘의집_주문_{dt.date.today():%Y%m%d}.xlsx"
    if out.lower().endswith(".csv"):
        _write_csv(out, orders, args.extra)
    else:
        _write_xlsx(out, orders, args.extra)
    print(f"{len(orders)}건을 {out} 로 저장했습니다.")
    return 0


def cmd_summary(args):
    with OrderStore(args.db) as store:
        counts = store.status_counts()
    if not counts:
        print("저장된 주문이 없습니다.")
        return 0
    total = sum(n for _, n in counts)
    print(f"전체 {total}건")
    for status, n in counts:
        print(f"  {status}: {n}")
    return 0


def build_parser():
    parser = argparse.ArgumentParser(
        prog="ohouse_orders",
        description="오늘의집 파트너센터 주문 엑셀을 모아 하나의 주문목록으로 관리합니다.",
    )
    parser.add_argument("--db", default=DEFAULT_DB, help=f"주문 DB 경로 (기본: {DEFAULT_DB})")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("import", help="주문 파일(xlsx/xls/csv)을 DB에 수집")
    p.add_argument("files", nargs="+", help="파일, 폴더 또는 와일드카드 (예: downloads/*.xlsx)")
    p.add_argument("--password", help="엑셀 파일 비밀번호 (또는 OHOUSE_EXCEL_PASSWORD 환경변수)")
    p.add_argument("--mapping", help='컬럼 매핑 JSON (예: {"receiver_phone": ["받는사람 핸드폰"]})')
    p.set_defaults(func=cmd_import)

    p = sub.add_parser("export", help="수집된 주문을 엑셀/CSV로 내보내기")
    p.add_argument("-o", "--output", help="출력 파일 (.xlsx 또는 .csv, 기본: 오늘의집_주문_YYYYMMDD.xlsx)")
    p.add_argument("--status", nargs="+", help="주문상태 필터 (예: --status 결제완료 배송준비중)")
    p.add_argument("--since", help="주문일시 시작 (YYYY-MM-DD)")
    p.add_argument("--until", help="주문일시 끝, 해당일 포함 (YYYY-MM-DD)")
    p.add_argument("--collected-since", help="이 시각 이후 처음 수집된 주문만 (YYYY-MM-DD[ HH:MM:SS])")
    p.add_argument("--extra", action="store_true", help="매핑되지 않은 원본 컬럼도 함께 출력")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("summary", help="주문상태별 건수 보기")
    p.set_defaults(func=cmd_summary)
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    return args.func(args)
