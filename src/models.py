from __future__ import annotations

from datetime import date
from typing import Literal

from pydantic import BaseModel, Field


class CompanyInfo(BaseModel):
    name_ko: str
    name_en: str | None = None
    registration_number: str
    ceo: str
    address: str
    phone: str | None = None
    email: str | None = None


class BrandColors(BaseModel):
    primary: str = "#0B3D91"
    accent: str = "#F5A623"
    text: str = "#1A1A1A"


class BrandingAssets(BaseModel):
    logo_path: str | None = None
    colors: BrandColors = Field(default_factory=BrandColors)
    font_family: str = "맑은 고딕"


class SignatureInfo(BaseModel):
    signer_name: str
    signer_title: str


class ContactPerson(BaseModel):
    name: str
    title: str | None = None
    phone: str | None = None
    email: str | None = None


class BankAccount(BaseModel):
    bank: str
    account_number: str
    account_holder: str | None = None


class Brand(BaseModel):
    brand_id: str
    company: CompanyInfo
    branding: BrandingAssets = Field(default_factory=BrandingAssets)
    signature: SignatureInfo
    contact_person: ContactPerson | None = None
    bank_account: BankAccount | None = None
    footer_text: str | None = None


class Counterparty(BaseModel):
    name: str
    registration_number: str | None = None
    address: str | None = None
    ceo: str | None = None
    contact_name: str | None = None
    contact_title: str | None = None
    email: str | None = None


class LineItem(BaseModel):
    name: str
    description: str | None = None
    qty: float | None = None
    period: float | None = None
    unit_price: float
    discount_rate: float | None = None     # 항목별 할인율 0~1 (예: 0.4 = 40%)
    discount_amount: float | None = None   # 항목별 할인 금액 (양수)
    currency: str = "KRW"
    notes: str | None = None
    # 청구 방식 메타. 'deferred_percent' 면 후불(QR결제 %) 항목 — PDF 에서
    # 단가/공급가 셀을 '-' 로 표시하고 합계에서 제외
    billing_type: str | None = None

    @property
    def gross_amount(self) -> float:
        """할인 적용 전 금액 (수량 × 기간 × 단가)."""
        q = self.qty if (self.qty is not None and self.qty != 0) else 1
        p = self.period if (self.period is not None and self.period != 0) else 1
        return q * p * self.unit_price

    @property
    def discount_value(self) -> float:
        """실제 차감되는 할인액. 할인금액 > 할인율 우선."""
        if self.discount_amount and self.discount_amount > 0:
            return float(self.discount_amount)
        if self.discount_rate:
            return self.gross_amount * float(self.discount_rate)
        return 0.0

    @property
    def amount(self) -> float:
        """할인 적용 후 금액 (PDF/합계에 사용)."""
        return self.gross_amount - self.discount_value


class Totals(BaseModel):
    subtotal: float | None = None
    vat_rate: float | None = None
    vat: float | None = None
    total: float | None = None
    currency: str = "KRW"


def ensure_totals(document, default_vat_rate: float) -> None:
    """문서의 totals에서 누락된 값을 자동 계산합니다.

    계산 순서:
      - 항목별 할인 후 금액 합 = items_sum
      - 전체 할인율 차감 = items_sum × (1 - total_discount_rate)  → subtotal
      - 부가세 별도 = subtotal × vat_rate
      - 합계 = subtotal + vat
    """
    t = document.totals
    if t.subtotal is None:
        items_sum = sum(item.amount for item in document.line_items)
        td = getattr(document, "total_discount_rate", None) or 0
        t.subtotal = items_sum * (1 - td)
    if t.vat_rate is None:
        t.vat_rate = default_vat_rate
    if t.vat is None:
        t.vat = round(t.subtotal * t.vat_rate)
    if t.total is None:
        t.total = t.subtotal + t.vat


class ClauseRef(BaseModel):
    id: str
    vars: dict[str, object] = Field(default_factory=dict)


class QuoteDocument(BaseModel):
    document_id: str
    document_type: Literal["quote", "contract"]
    brand_id: str
    issued_date: date
    valid_until: date | None = None
    effective_date: date | None = None
    contract_term: str | None = None
    counterparty: Counterparty
    subject: str
    line_items: list[LineItem]
    total_discount_rate: float | None = None  # 전체 일괄 할인율 0~1 (예: 0.1 = 10%)
    totals: Totals = Field(default_factory=Totals)
    clauses: list[ClauseRef] = Field(default_factory=list)
    notes: str | None = None
    custom_variables: dict[str, object] = Field(default_factory=dict)


class Clause(BaseModel):
    clause_id: str
    title: str
    category: str
    required: bool = False
    variables: list[str] = Field(default_factory=list)
    body_template: str
