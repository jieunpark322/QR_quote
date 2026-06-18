"""멤버십 클라우드 견적서 데이터 모델.

기존 QR 견적서(평면 line_items)와 달리, 다단계 계층 구조를 가짐:

  MembershipQuoteDocument
   └ scenarios: list[MembershipScenario]            (시나리오 — 엑셀 시트 1개)
      └ sections: list[MembershipSection]           (구분 — 예: 멤버십 클라우드)
         └ categories: list[MembershipCategory]     (분류 — 초기구축비/사용료/옵션)
            └ items: list[MembershipLineItem]       (상세 구분 — 실제 행)
               └ sub_items: list[MembershipSubItem] (종량제 하위 단가들)

unit_price 가 null 이고 unit_price_text 가 있으면 단가 칸을 텍스트로 표시.
default_amount_text 가 있으면 금액 칸을 텍스트로 표시 ("후청구", "협의 금액" 등).
"""
from __future__ import annotations

from datetime import date
from typing import Literal

from pydantic import BaseModel, Field


class MembershipSubItem(BaseModel):
    """종량제 하위 단가 (SMS, LMS, PUSH 등). 한 행 안에 줄바꿈으로 표시됨."""
    label: str = ""   # 예: "PUSH 발송 시"  (빈 문자열이면 줄 시작 인덴트만)
    spec: str         # 예: "9.9원/건", "(SMS) 9.9원/건"


class MembershipLineItem(BaseModel):
    """상세 구분 한 행."""
    name: str
    name_detail: str | None = None     # 이름 아래 보조 설명 (예: "- CS 대응\n- 현장 지원")
    billing_period: str | None = None  # 기간: "1회성", "매월", "발생시", "발생월", "1개당"
    unit_price: float | None = None    # 단가 (숫자). null 이면 unit_price_text 사용
    unit_price_text: str | None = None # 단가 텍스트 (예: "투입기간 X SW개발자 임금 단가")
    discount_rate: float | None = None      # 항목별 할인율 0~1 (0.4 = 40%)
    discount_amount: float | None = None    # 항목별 할인 금액 (양수)
    amount: float | None = None        # 금액 직접 지정. 없으면 자동 계산
    amount_text: str | None = None     # 금액 텍스트 (예: "후청구", "협의 금액", "무상제공")
    notes: str | None = None           # 비고
    sub_items: list[MembershipSubItem] = Field(default_factory=list)

    def gross_amount(self) -> float | None:
        """할인 적용 전 금액. unit_price 가 없으면 None."""
        if self.unit_price is None:
            return None
        return float(self.unit_price)

    def discount_value(self) -> float:
        """실제 차감 할인액. 할인금액 > 할인율."""
        if self.discount_amount and self.discount_amount > 0:
            return float(self.discount_amount)
        if self.discount_rate:
            g = self.gross_amount()
            if g is None:
                return 0.0
            return g * float(self.discount_rate)
        return 0.0

    def effective_amount(self) -> float | None:
        """소계에 포함될 실제 숫자 금액. 텍스트/None 이면 None."""
        if self.amount is not None:
            return self.amount
        if self.amount_text:
            return None
        g = self.gross_amount()
        if g is None:
            return None
        return g - self.discount_value()


class MembershipCategory(BaseModel):
    """분류 (초기구축비/사용료/옵션 등). 분류별 소계 행이 자동 삽입됨."""
    name: str  # "초기구축비" / "사용료" / "옵션"
    items: list[MembershipLineItem] = Field(default_factory=list)
    show_subtotal: bool = True      # "○○ 합계" 행 표시 여부
    subtotal_label: str | None = None  # 기본: "{name} 합계"


class MembershipSection(BaseModel):
    """구분 (멤버십 클라우드 등). 구분별 예상 총 금액이 자동 표시됨."""
    name: str
    categories: list[MembershipCategory] = Field(default_factory=list)
    show_section_total: bool = True  # "예상 총 금액" 표시 여부


class MembershipParty(BaseModel):
    """발행 정보의 한쪽 (제휴사 또는 회사). 양측이 좌우 배치됨."""
    label: str = "회사"           # 헤더 라벨 (예: "제휴사", "회사")
    name: str                     # 회사명
    address: str | None = None
    ceo: str | None = None        # 대표이사
    contact: str | None = None    # 담당자 (이름 + 직책)
    contact_phone: str | None = None
    contact_email: str | None = None


class MembershipScenario(BaseModel):
    """한 시나리오 (= 엑셀 시트 1개 / PDF에서 별도 페이지)."""
    name: str                                          # 시나리오 명 (예: "앱&POS 연동")
    subject: str | None = None                         # 항목 설명 부제
    sections: list[MembershipSection] = Field(default_factory=list)
    show_grand_total: bool = True                      # "전체 서비스 이용 금액" 표시


class MembershipQuoteDocument(BaseModel):
    """전체 멤버십 견적서."""
    document_id: str
    document_type: Literal["membership_quote"] = "membership_quote"
    brand_id: str = "softment"
    issued_date: date

    title: str = "멤버십 클라우드 견적서"            # 문서 메인 타이틀
    counterparty: MembershipParty                      # 제휴사(고객)
    supplier: MembershipParty | None = None            # 회사(우리). 없으면 brand에서 자동 채움

    scenarios: list[MembershipScenario] = Field(default_factory=list)

    total_discount_rate: float | None = None       # 전체 일괄 할인율 0~1 (시나리오별 적용)
    unit_notice: str = "(단위 : 원, 부가세별도)"      # 표 우측 상단 안내

    remarks: list[str] = Field(default_factory=lambda: [
        "견적유효 : 견적일로부터 15일",
        "결제조건 : 현금결제 (귀사 결제조건)",
    ])


# ─── 헬퍼 함수: 계산 ────────────────────────────────────────

def category_subtotal(cat: MembershipCategory) -> float:
    """분류 합계."""
    return sum((it.effective_amount() or 0) for it in cat.items)


def section_subtotals_by_period(sec: MembershipSection) -> dict[str, float]:
    """구분의 기간(1회성/매월 등)별 소계.

    엑셀의 '예상 총 금액' 형태로 결과 산출 (예: 초기구축비 / 월 사용료 분리).
    """
    by_period: dict[str, float] = {}
    for cat in sec.categories:
        for it in cat.items:
            amt = it.effective_amount()
            if amt is None:
                continue
            key = (it.billing_period or "기타").strip()
            by_period[key] = by_period.get(key, 0) + amt
    return by_period


def scenario_grand_total_by_period(sc: MembershipScenario) -> dict[str, float]:
    """시나리오 전체 기간별 총계."""
    total: dict[str, float] = {}
    for sec in sc.sections:
        for k, v in section_subtotals_by_period(sec).items():
            total[k] = total.get(k, 0) + v
    return total


def scenario_subtotal(sc: MembershipScenario) -> float:
    """시나리오 전체 공급가액 (모든 기간 합산)."""
    return sum(scenario_grand_total_by_period(sc).values())


def scenario_vat_and_total(sc: MembershipScenario,
                           vat_rate: float = 0.10,
                           total_discount_rate: float | None = None,
                           ) -> tuple[float, float, float]:
    """시나리오의 (공급가액, 부가세, 합계 금액) 계산.
    전체 할인율 적용 후 부가세 계산.
    """
    items_sum = scenario_subtotal(sc)
    td = total_discount_rate or 0
    subtotal = items_sum * (1 - td)
    vat = round(subtotal * vat_rate)
    return subtotal, vat, subtotal + vat
