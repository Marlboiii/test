"""파트너센터에서 내려받은 주문 파일(xlsx/xls/csv)을 읽어 표준 주문 레코드로 변환한다."""

import csv
import datetime as dt
import io
import re
from pathlib import Path

from .columns import FIELDS, REQUIRED_FIELDS, build_column_map, load_aliases, looks_like_header

HEADER_SCAN_ROWS = 20


class OrderFileError(Exception):
    pass


# ---------------------------------------------------------------- raw rows

def _read_csv_rows(path):
    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "cp949"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise OrderFileError(f"{path}: CSV 인코딩을 판별할 수 없습니다 (UTF-8/CP949 아님)")
    return [row for row in csv.reader(io.StringIO(text))]


def _decrypt(path, password):
    try:
        import msoffcrypto
    except ImportError as exc:
        raise OrderFileError(
            f"{path}: 비밀번호가 걸린 파일입니다. `pip install msoffcrypto-tool` 후 --password 로 실행하세요"
        ) from exc
    with open(path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        if not office.is_encrypted():
            return None
        if not password:
            raise OrderFileError(f"{path}: 비밀번호가 걸린 파일입니다. --password 옵션을 지정하세요")
        out = io.BytesIO()
        try:
            office.load_key(password=password)
            office.decrypt(out)
        except Exception as exc:
            raise OrderFileError(f"{path}: 파일 암호 해제 실패 (비밀번호 확인): {exc}") from exc
    out.seek(0)
    return out


def _is_ole(path):
    # 암호화된 xlsx와 구형 xls는 OLE 컨테이너(D0 CF 11 E0)로 저장된다
    with open(path, "rb") as f:
        return f.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _read_xlsx_rows(source):
    import openpyxl

    wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _read_xls_rows(source):
    try:
        import xlrd
    except ImportError as exc:
        raise OrderFileError("구형 .xls 파일입니다. `pip install xlrd` 후 다시 실행하세요") from exc
    if isinstance(source, io.BytesIO):
        book = xlrd.open_workbook(file_contents=source.getvalue())
    else:
        book = xlrd.open_workbook(source)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(sheet.nrows):
        row = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_DATE:
                row.append(xlrd.xldate_as_datetime(cell.value, book.datemode))
            else:
                row.append(cell.value)
        rows.append(row)
    return rows


def read_rows(path, password=None):
    """파일 형식에 맞게 읽어 2차원 리스트로 반환."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _read_csv_rows(path)
    if suffix not in (".xlsx", ".xlsm", ".xls"):
        raise OrderFileError(f"{path}: 지원하지 않는 형식입니다 (xlsx, xls, csv 만 가능)")

    if not _is_ole(path):
        return _read_xlsx_rows(path)
    # OLE 컨테이너: 구형 xls 이거나 비밀번호가 걸린 xlsx/xls
    if suffix == ".xls" and not password:
        return _read_xls_rows(path)
    decrypted = _decrypt(path, password)
    if decrypted is None:
        return _read_xls_rows(path)
    if decrypted.getvalue()[:2] == b"PK":
        return _read_xlsx_rows(decrypted)
    return _read_xls_rows(decrypted)


# ------------------------------------------------------------- normalizing

DATE_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
    "%Y.%m.%d %H:%M:%S",
    "%Y.%m.%d %H:%M",
    "%Y.%m.%d",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y/%m/%d",
    "%Y%m%d%H%M%S",
    "%Y%m%d",
)


def parse_datetime(value):
    if value in (None, ""):
        return None
    if isinstance(value, dt.datetime):
        return value.replace(microsecond=0).isoformat(sep=" ")
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day).isoformat(sep=" ")
    text = re.sub(r"\s+", " ", str(value).strip())
    text = text.replace("T", " ")
    text = re.sub(r"\.\d+$", "", text)  # 초 이하 제거
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).isoformat(sep=" ")
        except ValueError:
            continue
    return text  # 알 수 없는 형식은 원문 유지


def parse_int(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return int(round(value))
    digits = re.sub(r"[^\d\-.]", "", str(value))
    if digits in ("", "-", "."):
        return None
    try:
        return int(round(float(digits)))
    except ValueError:
        return None


def to_text(value):
    """엑셀이 숫자로 바꿔버린 주문번호 등을 원래 문자열로 되돌린다."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    return str(value).strip()


def normalize_phone(value):
    text = to_text(value)
    if not text:
        return ""
    digits = re.sub(r"\D", "", text)
    # 엑셀이 앞자리 0을 지운 경우 복원 (1012345678 -> 01012345678)
    if len(digits) in (9, 10) and digits[0] == "1":
        digits = "0" + digits
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    if len(digits) == 10:
        if digits.startswith("02"):
            return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return text


def normalize_zipcode(value):
    text = to_text(value)
    if text.isdigit() and len(text) < 5:
        text = text.zfill(5)
    return text


def normalize_record(raw):
    record = {}
    for field in FIELDS:
        value = raw.get(field)
        if field == "ordered_at":
            record[field] = parse_datetime(value)
        elif field in ("quantity", "amount", "shipping_fee"):
            record[field] = parse_int(value)
        elif field in ("buyer_phone", "receiver_phone"):
            record[field] = normalize_phone(value)
        elif field == "zipcode":
            record[field] = normalize_zipcode(value)
        else:
            record[field] = to_text(value)
    return record


# ----------------------------------------------------------------- public

def find_header(rows, aliases):
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        if looks_like_header(row, aliases):
            return i
    return None


def read_orders(path, password=None, mapping_path=None):
    """주문 파일을 읽어 표준 레코드 리스트를 반환한다.

    각 레코드는 FIELDS 키와, 매핑되지 않은 컬럼을 담은 'extra'(dict)를 가진다.
    """
    aliases = load_aliases(mapping_path)
    rows = read_rows(path, password=password)
    header_idx = find_header(rows, aliases)
    if header_idx is None:
        raise OrderFileError(
            f"{path}: 헤더 행(주문번호 컬럼)을 찾지 못했습니다. --mapping 으로 컬럼명을 지정하세요"
        )

    headers = [to_text(h) for h in rows[header_idx]]
    column_map = build_column_map(headers, aliases)
    missing = [f for f in REQUIRED_FIELDS if f not in column_map.values()]
    if missing:
        raise OrderFileError(f"{path}: 필수 컬럼 누락: {', '.join(missing)}")

    orders = []
    for row in rows[header_idx + 1:]:
        if not any(cell not in (None, "") for cell in row):
            continue
        raw, extra = {}, {}
        for i, cell in enumerate(row):
            if i in column_map:
                raw[column_map[i]] = cell
            elif i < len(headers) and headers[i] and cell not in (None, ""):
                extra[headers[i]] = to_text(cell)
        record = normalize_record(raw)
        if not record["order_no"]:
            continue  # 합계 행 등
        record["extra"] = extra
        record["source_file"] = Path(path).name
        orders.append(record)
    return orders
