"""양식 라벨·기본값 설정 로더.

`config/labels.json` 의 값을 읽어 견적서/계약서 렌더링에 사용합니다.
설정 파일이 없거나 일부 키가 빠지면 아래 기본값이 사용됩니다.
"""
from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, Field


class QuoteTextLabels(BaseModel):
    counterparty_section: str = "수 신"
    subject_prefix: str = "건명"
    etc_notice_section: str = "기타 안내"
    vat_separate_notice: str = "* VAT 별도"
    subtotal: str = "공급가액"
    vat: str = "부가세"
    total: str = "합계 금액"


class QuoteTableHeaders(BaseModel):
    name: str = "항목"
    description: str = "설명"
    qty: str = "수량"
    period: str = "기간(횟수)"
    unit_price: str = "단가"
    amount: str = "공급가"
    notes: str = "비고"


class QuoteAutoNotices(BaseModel):
    # 빈 문자열 = 자동 표시 안 함
    validity_template: str = "본 견적의 유효기간은 {valid_until}까지입니다."
    bank_account_template: str = "입금 계좌: {bank} {account_number} (예금주: {account_holder})"


class QuoteConfig(BaseModel):
    title: str = "견 적 서"
    vat_rate: float = 0.10
    default_validity_days: int = 30
    labels: QuoteTextLabels = Field(default_factory=QuoteTextLabels)
    table_headers: QuoteTableHeaders = Field(default_factory=QuoteTableHeaders)
    auto_notices: QuoteAutoNotices = Field(default_factory=QuoteAutoNotices)
    # 항상 기타 안내에 함께 표시되는 고정 문구 목록
    static_notices: list[str] = Field(default_factory=list)


class ContractTextLabels(BaseModel):
    overview_section: str = "계약 개요"
    effective_date: str = "계약 시작일"
    contract_term: str = "계약 기간"
    amount_pre_vat: str = "계약 금액 (VAT 별도)"
    amount_with_vat: str = "계약 금액 (VAT 포함)"
    party_a: str = "갑 (공급자)"
    party_b: str = "을 (수신자)"


class ContractConfig(BaseModel):
    title: str = "계 약 서"
    preamble_template: str = (
        '{supplier_name}(이하 "갑"이라 한다)과(와) '
        '{counterparty_name}(이하 "을"이라 한다)은(는) 위 사안에 관하여 '
        '다음과 같이 합의하여 본 계약을 체결한다.'
    )
    labels: ContractTextLabels = Field(default_factory=ContractTextLabels)


class DocumentLabels(BaseModel):
    quote: QuoteConfig = Field(default_factory=QuoteConfig)
    contract: ContractConfig = Field(default_factory=ContractConfig)


def load_labels(project_root: Path) -> DocumentLabels:
    """`config/labels.json` 을 로드. 파일이 없으면 모든 기본값 사용."""
    path = project_root / "config" / "labels.json"
    if not path.exists():
        return DocumentLabels()
    return DocumentLabels.model_validate_json(path.read_text(encoding="utf-8"))
