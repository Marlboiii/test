"""엑셀 헤더(한글 컬럼명)를 표준 필드로 매핑한다.

오늘의집 파트너센터의 주문 엑셀 양식은 바뀔 수 있으므로, 컬럼명은
여러 별칭을 허용하고 --mapping 옵션(JSON)으로 덮어쓸 수 있게 한다.
"""

import json
import re

# 표준 필드 -> 허용하는 헤더 별칭 (공백/괄호 제거 후 비교)
FIELD_ALIASES = {
    "order_no": ["주문번호", "주문ID", "주문아이디"],
    "item_no": ["주문상품번호", "상품주문번호", "주문옵션번호", "주문상세번호", "옵션주문번호"],
    "ordered_at": ["주문일시", "주문일", "주문일자", "결제일시", "결제일", "결제완료일"],
    "status": ["주문상태", "상태", "배송상태", "처리상태"],
    "product_id": ["상품번호", "상품ID", "상품코드"],
    "product_name": ["상품명", "주문상품명", "상품이름"],
    "option": ["옵션", "옵션명", "옵션정보", "주문옵션", "선택옵션"],
    "quantity": ["수량", "주문수량", "구매수량"],
    "amount": ["결제금액", "주문금액", "총결제금액", "상품결제금액", "판매금액", "판매가"],
    "shipping_fee": ["배송비", "배송비금액"],
    "buyer_name": ["주문자", "주문자명", "구매자", "구매자명"],
    "buyer_phone": ["주문자연락처", "주문자휴대폰", "주문자전화번호", "구매자연락처"],
    "receiver_name": ["수령인", "수령인명", "수령자", "수취인", "수취인명", "받는분"],
    "receiver_phone": ["수령인연락처", "수령인휴대폰", "수령인전화번호", "수취인연락처", "수취인휴대폰", "받는분연락처", "연락처", "휴대폰"],
    "zipcode": ["우편번호", "수령인우편번호"],
    "address": ["주소", "배송지", "배송주소", "수령인주소", "수취인주소", "받는분주소"],
    "delivery_memo": ["배송메모", "배송요청사항", "배송메세지", "배송메시지", "요청사항"],
    "courier": ["택배사", "배송업체", "택배회사"],
    "tracking_no": ["송장번호", "운송장번호", "운송장"],
}

FIELDS = list(FIELD_ALIASES)
REQUIRED_FIELDS = ["order_no"]


def normalize_header(value):
    """'수령인 연락처(휴대폰)' -> '수령인연락처' 처럼 비교용으로 정리."""
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\(.*?\)|\[.*?\]", "", text)
    text = re.sub(r"[\s_\-/·.:*]", "", text)
    return text.lower()


def load_aliases(mapping_path=None):
    """기본 별칭에 사용자 매핑(JSON: {"필드": ["헤더", ...]})을 우선 적용한다."""
    aliases = {field: list(names) for field, names in FIELD_ALIASES.items()}
    if mapping_path:
        with open(mapping_path, encoding="utf-8") as f:
            custom = json.load(f)
        for field, names in custom.items():
            if field not in aliases:
                raise ValueError(f"알 수 없는 필드: {field} (사용 가능: {', '.join(FIELDS)})")
            if isinstance(names, str):
                names = [names]
            aliases[field] = list(names) + aliases[field]
    return aliases


def build_column_map(headers, aliases):
    """헤더 목록에서 {컬럼 인덱스: 필드} 매핑을 만든다.

    별칭 목록의 앞쪽일수록 우선한다. 매핑되지 않은 컬럼은 extra로 보관된다.
    """
    normalized = [normalize_header(h) for h in headers]
    column_map = {}
    for field, names in aliases.items():
        for name in names:
            target = normalize_header(name)
            idx = next(
                (i for i, h in enumerate(normalized) if h == target and i not in column_map),
                None,
            )
            if idx is not None:
                column_map[idx] = field
                break
    return column_map


def looks_like_header(row, aliases):
    """행이 헤더 행인지 판단 (주문번호 컬럼이 있으면 헤더로 본다)."""
    order_names = {normalize_header(n) for n in aliases["order_no"]}
    return any(normalize_header(cell) in order_names for cell in row)
