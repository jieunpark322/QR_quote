# B2B 계약서/견적서 자동화 시스템

엑셀에서 품목을 드롭다운으로 선택하면 가격이 자동 반영되고, 한 번의 명령으로 **DOCX와 PDF** 가 동시에 생성되는 화이트라벨 시스템. **견적서**와 **계약서** 두 가지 문서를 같은 엔진으로 생성하고, 브랜드(회사명·색상·로고)는 설정 파일로 갈아 끼울 수 있습니다.

## 지원하는 문서

| 문서 | 입력 방식 | 핵심 구성 |
|---|---|---|
| **견적서 (Quote)** | 엑셀 입력 폼 또는 JSON | 공급자/수신처, VAT 별도 표시 내역, 공급가액·부가세·합계, 기타 안내(유효기간 자동) |
| **계약서 (Contract)** | JSON | 전문, 갑/을 정보표, 계약 개요(기간·금액), 조항 7개(모듈식), 양측 서명란 |

## 사용 방법 — 3가지

### 🌐 방법 1: 웹 인터페이스 (가장 쉬움)

**로컬 PC 에서 사용:**

```powershell
.\.venv\Scripts\python.exe -m src.cli web
```

→ 브라우저가 자동으로 열리고 `http://localhost:8501` 에 견적서 작성 폼이 표시됩니다. 좌측 사이드바의 **카탈로그 빠른 추가** 버튼으로 상품을 클릭해 넣고, 수신처와 건명 채운 뒤 **견적서 생성** 버튼 → DOCX/PDF 다운로드.

페이지는 3개:
- 📋 견적서 작성
- 📦 카탈로그 관리 (상품 추가/수정/삭제)
- ⚙ 설정 (브랜드 정보, 양식 라벨/문구 편집)

**클라우드(인터넷)에 무료로 배포해서 어디서나 사용:**

→ **[CLOUD_DEPLOY.md](CLOUD_DEPLOY.md)** 단계별 가이드 참고. Streamlit Community Cloud 로 무료 배포 가능.

### 📊 방법 2: 엑셀 입력

```
1. 빈 엑셀 템플릿 생성  →  2. 엑셀 작성 (Excel/LibreOffice)  →  3. DOCX/PDF 자동 생성
```

### 📝 방법 3: JSON 직접 작성 (개발자/배치 처리용)

`data/quotes/` 에 JSON 파일을 만들어 처리. 자세한 형식은 `data/quotes/sample.json` 참고.

## 빠른 사용법

PowerShell에서 `contract-system` 폴더로 이동 후:

### ① 빈 엑셀 입력 템플릿 만들기
```powershell
.\.venv\Scripts\python.exe -m src.cli template --out "input\견적서_작성.xlsx"
```

### ② 엑셀에서 작성
- 엑셀로 `input\견적서_작성.xlsx` 열기
- **품목** 컬럼은 셀 클릭 시 ▼ 드롭다운이 나타남 → 카탈로그 품목 선택
- 선택 시 **설명·단가가 자동 입력** 됨 (수정 가능)
- 드롭다운에 없는 품목은 **직접 입력**, 단가도 직접 입력
- 발행일·유효기간·수신처 정보 채우고 저장

### ③ DOCX + PDF 생성
```powershell
.\.venv\Scripts\python.exe -m src.cli render --input "input\견적서_작성.xlsx" --out output
```

생성물: `output/Q-YYYYMMDD-<고객명>.docx` 와 `.pdf`

## 폴더 구조

```
contract-system/
├── brands/              # 브랜드별 설정 (회사명·색상·로고·담당자)
│   └── softment/
│       └── brand.json
├── catalog/             # 품목 카탈로그 (엑셀 드롭다운 + 가격 자동입력 원본)
│   └── products.json
├── clauses/             # (계약서용) 모듈식 조항
├── data/                # JSON 직접 입력 방식 (선택적)
├── input/               # 엑셀 입력 파일
├── output/              # 생성된 DOCX/PDF
├── src/                 # Python 소스코드
└── requirements.txt
```

## 품목 카탈로그 관리

엑셀 드롭다운과 가격 자동입력의 원본은 **`catalog/products.json`** 입니다. 새 상품을 추가하거나 가격을 변경하려면 이 파일을 편집한 뒤, 엑셀 템플릿을 다시 생성하면 됩니다.

```json
{
  "products": [
    {
      "code": "ENT-ANNUAL",
      "name": "Enterprise Plan (연간 구독)",
      "description": "사용자 50명, 24/7 기술지원 포함",
      "unit_price": 6000000,
      "currency": "KRW"
    }
  ]
}
```

엑셀 안의 `품목카탈로그` 시트에서도 가격을 즉석 수정할 수 있지만, **영구 변경은 `products.json` 에 반영**해야 다음 템플릿 생성 시에도 유지됩니다.

## 브랜드(화이트라벨) 추가

`brands/<브랜드ID>/brand.json` 을 만들고 회사명·색상·서명자·담당자 정보를 설정. 엑셀의 "브랜드 ID" 셀에 해당 ID를 입력하거나, CLI에 `--brand <ID>` 를 주면 그 브랜드로 문서가 생성됩니다.

현재 등록된 브랜드:
- **softment** — Softment Blue (#5A8CDC) / Mint (#64C8BE), 공식 CI 적용

브랜드별 로고는 `brands/<브랜드ID>/assets/logo.png` 에 두면 문서 상단에 자동 삽입됩니다.

## 계약서 사용 예시

```powershell
.\.venv\Scripts\python.exe -m src.cli render --input "data\contracts\sample.json" --out output
```

계약서는 JSON으로 입력하며 (`data/contracts/sample.json` 참고), 다음 조항들이 모듈식으로 들어있습니다:

| 조항 ID | 제목 | 변수 |
|---|---|---|
| `service_scope` | 계약의 목적 및 공급 범위 | — |
| `contract_term` | 계약 기간 및 갱신 | — |
| `payment` | 대금 및 지급 조건 | `payment_days` |
| `ip_ownership` | 지적재산권 | — |
| `confidentiality` | 비밀유지 | `period_years` |
| `termination` | 계약의 해지 | `notice_days` |
| `governing_law` | 준거법 및 분쟁 해결 | — |

`clauses` 배열에서 ID를 빼거나 순서를 바꾸면, 본문의 "제 N조" 번호가 자동으로 재정렬됩니다. 새 조항을 추가하려면 `clauses/contract/` 폴더에 같은 형식의 markdown 파일을 만들면 됩니다.

## CLI 명령 요약

| 명령 | 용도 |
|---|---|
| `web` | **웹 인터페이스 실행** (브라우저에서 폼 작성) |
| `template --out <경로>` | 빈 엑셀 입력 템플릿 생성 |
| `render --input <엑셀 또는 JSON>` | 입력 파일에서 DOCX(+PDF) 생성 |
| `render ... --no-pdf` | PDF 생성 건너뛰기 (DOCX만) |
| `render ... --brand <ID>` | 브랜드 강제 지정 |

## JSON 직접 입력 방식 (선택)

엑셀 없이 JSON 으로 작성하고 싶다면 `data/quotes/sample.json` 형식 참고:

```powershell
.\.venv\Scripts\python.exe -m src.cli render --input "data\quotes\sample.json" --out output
```

## 양식 수정 가이드 (비개발자용)

자주 바꿀 만한 항목은 모두 **JSON 파일 한두 개를 편집**하는 것으로 끝납니다. Python 코드는 건드릴 필요 없습니다. **수정 후 견적서/계약서를 다시 생성**(`python -m src.cli render ...`)하면 반영됩니다.

### 📄 `config/labels.json` — 양식의 모든 라벨·기본값

이 파일이 **'내부 양식의 기준'** 을 모은 곳입니다. 견적서/계약서에 나오는 거의 모든 텍스트가 여기서 옵니다.

| 바꾸고 싶은 것 | 편집 위치 |
|---|---|
| **유효기간 자동 문구** ("본 견적의 유효기간은 ... 까지입니다.") | `quote.auto_notices.validity_template` |
| **입금 계좌 자동 문구** ("입금 계좌: ...") | `quote.auto_notices.bank_account_template` |
| **모든 견적서에 항상 들어가는 고정 안내문구** (예: "본 견적은 부가세 별도입니다.") | `quote.static_notices` (배열에 추가) |
| **VAT율 변경** (10% → 다른 비율) | `quote.vat_rate` (예: `0.10` → `0.05`) |
| **유효기간 기본 일수** (30일 → 다른 값) | `quote.default_validity_days` |
| **합계 라벨** ("공급가액", "부가세", "합계 금액") | `quote.labels.subtotal` / `.vat` / `.total` |
| **표 컬럼명** ("품목", "설명", "수량", "기간(회)", "단가", "금액", "비고") | `quote.table_headers` |
| **"수 신" 섹션 라벨** | `quote.labels.counterparty_section` |
| **"건명:" prefix** | `quote.labels.subject_prefix` |
| **"기타 안내" 섹션 라벨** | `quote.labels.etc_notice_section` |
| **"* VAT 별도" 표시 문구** | `quote.labels.vat_separate_notice` |
| **견적서/계약서 제목** | `quote.title` / `contract.title` |
| **계약서 전문 문장** | `contract.preamble_template` (`{supplier_name}`, `{counterparty_name}` 치환됨) |
| **계약서 갑/을 박스 라벨** | `contract.labels.party_a` / `.party_b` |
| **계약 개요 라벨** ("계약 시작일" 등) | `contract.labels.effective_date` 등 |

> 💡 **자동 문구 끄기**: 템플릿 문자열을 `""` (빈 따옴표) 로 두면 해당 자동 안내가 표시되지 않습니다.

### 📄 `brands/softment/brand.json` — 회사 정보·CI

| 바꾸고 싶은 것 | 편집 위치 |
|---|---|
| 사업자등록번호, 대표자, 주소 | `company.registration_number` / `.ceo` / `.address` |
| 색상 (CI 변경 시) | `branding.colors.primary` / `.accent` / `.text` |
| 로고 이미지 | `brands/softment/assets/logo.png` 파일 자체를 교체 |
| 담당자 정보 (박지은) | `contact_person.name` / `.title` / `.phone` / `.email` |
| 입금 계좌 | `bank_account.bank` / `.account_number` / `.account_holder` |

### 📄 `catalog/products.json` — 판매 상품 목록 (엑셀 드롭다운 원본)

상품을 추가/수정하려면 이 파일의 `products` 배열에 항목을 넣거나 수정하면 됩니다. 엑셀 템플릿을 다시 생성(`python -m src.cli template ...`) 하면 드롭다운에 반영됩니다.

### 📄 `clauses/contract/*.md` — 계약서 조항

각 조항이 markdown 파일 하나입니다. 본문 텍스트나 `{{ payment_days }}` 같은 변수 자리를 자유롭게 편집할 수 있습니다.

### ⚠ 편집 시 주의사항

- JSON 파일은 **쉼표·중괄호·따옴표** 가 정확해야 합니다. VS Code에서 열면 오타를 빨간 줄로 알려줍니다.
- 변경 후 **항상 재생성** (`python -m src.cli render ...`) 해야 반영됩니다.
- 헷갈리면 변경 전 파일을 복사해두고 시도하세요.

## 의존성

- Python 3.11+
- LibreOffice (PDF 변환용, 백그라운드에서만 사용)
- 패키지: `python-docx`, `openpyxl`, `jinja2`, `python-frontmatter`, `pydantic`, `typer`, `Pillow`
