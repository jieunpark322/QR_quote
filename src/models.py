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
    currency: str = "KRW"
    notes: str | None = None

    @property
    def amount(self) -> float:
        # 수량/기간이 None 또는 0(의미 없는 값) 이면 1로 처리
        q = self.qty if (self.qty is not None and self.qty != 0) else 1
        p = self.period if (self.period is not None and self.period != 0) else 1
        return q * p * self.unit_price


class Totals(BaseModel):
    subtotal: float | None = None
    vat_rate: float | None = None
    vat: float | None = None
    total: float | None = None
    currency: str = "KRW"


def ensure_totals(document, default_vat_rate: float) -> None:
    """문서의 totals에서 누락된 값을 자동 계산합니다.

    - subtotal: 누락 시 line_items의 amount 합으로
    - vat_rate: 누락 시 labels의 기본값으로
    - vat: 누락 시 subtotal × vat_rate 반올림
    - total: 누락 시 subtotal + vat
    """
    t = document.totals
    if t.subtotal is None:
        t.subtotal = sum(item.amount for item in document.line_items)
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
