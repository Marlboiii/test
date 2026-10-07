# 오늘의집 주문수집기

오늘의집 파트너센터에서 내려받은 **주문 엑셀(xlsx/xls/csv)** 을 읽어
하나의 주문 DB(SQLite)에 모으고, 상태·기간별로 정리된 엑셀/CSV로 내보냅니다.

> 오늘의집 Open API는 승인된 연동 솔루션(이지어드민, 사방넷 등)에만 제공되므로,
> 이 도구는 파트너센터의 **엑셀 다운로드 파일**을 입력으로 사용합니다.

## 설치

```bash
cd order-collector
pip install -r requirements.txt
```

`openpyxl`만 있으면 기본 기능이 동작합니다. 비밀번호가 걸린 엑셀은 `msoffcrypto-tool`,
구형 `.xls`는 `xlrd`가 필요합니다.

## 사용법

1. 오늘의집 파트너센터 → 주문/배송 관리에서 주문 목록을 엑셀로 다운로드해 `downloads/` 폴더에 저장
2. 수집

```bash
python -m ohouse_orders import downloads/
# 비밀번호가 걸린 파일이라면
python -m ohouse_orders import downloads/ --password 1234
#   (또는 OHOUSE_EXCEL_PASSWORD 환경변수 사용)
```

같은 주문을 여러 번 가져와도 중복 저장되지 않고, 주문상태·송장번호 등 바뀐 값만 갱신됩니다.

```
[완료] 주문내역_1007.xlsx: 42건 읽음 → 신규 15, 변경 3, 동일 24
합계: 신규 15건, 변경 3건 (DB: ohouse_orders.db)
```

3. 확인 / 내보내기

```bash
python -m ohouse_orders summary                                  # 상태별 건수
python -m ohouse_orders export                                   # 오늘의집_주문_YYYYMMDD.xlsx
python -m ohouse_orders export --status 결제완료 -o 발주.xlsx     # 신규 주문만
python -m ohouse_orders export --since 2026-10-01 --until 2026-10-07 -o 10월1주.csv
python -m ohouse_orders export --collected-since "2026-10-07 09:00"   # 오늘 새로 들어온 주문
python -m ohouse_orders export --extra                           # 원본의 기타 컬럼까지 포함
```

## 컬럼 인식

헤더 행은 파일 상단 20행 안에서 `주문번호` 컬럼을 찾아 자동 인식합니다(제목 행이 있어도 됨).
`주문번호, 주문상품번호, 주문일시, 주문상태, 상품명, 옵션, 수량, 결제금액, 주문자, 수령인,
수령인 연락처, 우편번호, 주소, 배송메모, 택배사, 송장번호` 등의 일반적인 컬럼명과 별칭을 인식하며
(`ohouse_orders/columns.py` 참고), 인식되지 않은 컬럼은 `--extra` 로 그대로 내보낼 수 있습니다.

파트너센터 양식의 컬럼명이 다르다면 매핑 파일을 만들어 지정하세요.

```json
{
  "receiver_phone": ["받는사람 핸드폰"],
  "item_no": ["주문옵션 ID"]
}
```

```bash
python -m ohouse_orders import downloads/ --mapping mapping.json
```

정리 기능: 엑셀이 숫자로 바꾼 주문번호 복원, 앞자리 0이 빠진 전화번호/우편번호 복원,
`59,800원` → `59800`, 다양한 날짜 형식 → `YYYY-MM-DD HH:MM:SS`, CSV의 UTF-8/CP949 자동 판별.

## 테스트

```bash
python -m unittest discover -s tests -t .
```

## 주의

주문 파일과 DB에는 고객 개인정보(이름·연락처·주소)가 들어 있습니다.
`.gitignore`에서 `*.db`, `*.xlsx`, `*.csv`, `downloads/` 를 제외해 두었으니 저장소에 올리지 마세요.
