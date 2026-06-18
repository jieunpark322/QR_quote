"""멤버십 클라우드 견적서 DOCX 렌더러.

기존 quote 렌더러(평면 line_items)와 별도로 동작. 시나리오마다 페이지 분리,
계층 표(구분/분류/항목 + 종량제 하위), 자동 소계·총계.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .membership_models import (
    MembershipCategory,
    MembershipLineItem,
    MembershipQuoteDocument,
    MembershipScenario,
    MembershipSection,
    category_subtotal,
    scenario_grand_total_by_period,
    section_subtotals_by_period,
)
from .labels import load_labels
from .models import Brand


# ─── 색상/스타일 헬퍼 ────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _apply_font(run, font_name: str, *, size_pt: float | None = None,
                bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold:
        run.bold = True
    if color is not None:
        run.font.color.rgb = color


def _set_cell_bg(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tc_pr.append(shd)


def _vcenter(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:vAlign")):
        tc_pr.remove(old)
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), "center")
    tc_pr.append(v)


def _set_table_borders(table, color: str = "9FB0CC", size: int = 4) -> None:
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(size))
        b.set(qn("w:color"), color.lstrip("#"))
        borders.append(b)
    tbl_pr.append(borders)


def _merge_vertical(table, col: int, start_row: int, end_row: int) -> None:
    """col(0-indexed) 컬럼의 start_row..end_row 를 수직 병합."""
    if end_row <= start_row:
        return
    first = table.cell(start_row, col)
    last = table.cell(end_row, col)
    first.merge(last)


def _add_paragraph(doc, text: str, *, font: str, size_pt: float = 10,
                   bold: bool = False, alignment=None, color: RGBColor | None = None,
                   space_after_pt: float = 4, space_before_pt: float = 0):
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.space_before = Pt(space_before_pt)
    run = p.add_run(text)
    _apply_font(run, font, size_pt=size_pt, bold=bold, color=color)
    return p


def _format_won(amount: float) -> str:
    return f"{int(round(amount)):,}"


def _money(amount: float) -> str:
    """부호 우선 통화 표기 (-₩1,000,000 / ₩1,000,000)."""
    val = int(round(amount))
    if val < 0:
        return f"-₩{abs(val):,}"
    return f"₩{val:,}"


def _compute_section_widths(section: MembershipSection,
                            usable_cm: float = 18.8) -> list:
    """콘텐츠 길이 + 중요도 기반으로 컬럼 너비를 유동 산출.

    [구분, 분류, 상세 구분, 기간, 단가, 할인율, 금액, 비고]
    """
    cats = section.categories
    items = [it for cat in cats for it in cat.items]

    # 각 컬럼의 최대 콘텐츠 글자 수 (한 줄 기준, 가장 긴 줄)
    def _max_line_len(strings):
        m = 0
        for s in strings:
            if not s:
                continue
            for ln in str(s).split("\n"):
                m = max(m, len(ln))
        return m

    # 헤더는 기본 한국어 2~3자 (구분, 분류, 기간 등)
    sec_len = _max_line_len([section.name, "구분"])
    cat_len = _max_line_len([c.name for c in cats] + ["분류"])

    # 상세 구분: name + name_detail + sub_items 줄들 최대치
    detail_strings = []
    for it in items:
        detail_strings.append(it.name)
        if it.name_detail:
            detail_strings.append(it.name_detail)
        if it.sub_items:
            # 종량제 줄 합치기 패턴 그대로
            sub_lines = []
            for s in it.sub_items:
                if s.label:
                    sub_lines.append(f"· {s.label}  {s.spec}")
                elif sub_lines:
                    sub_lines[-1] += f"  {s.spec}"
                else:
                    sub_lines.append(f"   {s.spec}")
            detail_strings.extend(sub_lines)
    detail_len = max(_max_line_len(detail_strings + ["상세 구분"]), 8)

    period_len = _max_line_len([it.billing_period for it in items] + ["기간"])

    unit_price_strings = []
    for it in items:
        if it.unit_price is not None:
            unit_price_strings.append(f"₩{int(it.unit_price):,}")
        elif it.unit_price_text:
            unit_price_strings.append(it.unit_price_text)
    unit_price_len = _max_line_len(unit_price_strings + ["단가"])

    has_discount = any(it.discount_rate for it in items)

    amount_strings = []
    for it in items:
        if it.amount_text:
            amount_strings.append(it.amount_text)
        elif it.effective_amount() is not None:
            amount_strings.append(f"₩{int(it.effective_amount()):,}")
        else:
            amount_strings.append("-")
    # 분류 합계도 금액 컬럼에 표시
    for cat in cats:
        st = category_subtotal(cat)
        if st:
            amount_strings.append(f"₩{int(st):,}")
    amount_len = _max_line_len(amount_strings + ["금액"])

    notes_len = _max_line_len([it.notes or "" for it in items] + ["비고"])

    # 글자수 → cm 환산 (작은 폰트라 한 글자 ~0.15cm 가정)
    CHAR_CM = 0.16
    raw = [
        sec_len * CHAR_CM,
        cat_len * CHAR_CM,
        detail_len * CHAR_CM,
        period_len * CHAR_CM,
        unit_price_len * CHAR_CM,
        2 * CHAR_CM,  # 할인율 ("40%" 또는 "-")
        amount_len * CHAR_CM,
        notes_len * CHAR_CM,
    ]

    # 최소·최대 제약
    min_widths = [1.4, 1.0, 3.5, 0.9, 1.5, 0.7, 1.5, 1.0]
    max_widths = [3.0, 2.2, 7.5, 1.8, 4.0, 1.0, 3.5, 3.0]
    if not has_discount:
        max_widths[5] = 0.8  # 할인율 모두 없으면 더 좁게

    constrained = [
        min(max(raw_w, min_widths[i]), max_widths[i])
        for i, raw_w in enumerate(raw)
    ]

    # 합계가 사용 가능한 폭을 넘으면 비율 스케일다운
    total = sum(constrained)
    if total > usable_cm:
        scale = usable_cm / total
        constrained = [w * scale for w in constrained]
    else:
        # 남는 공간은 콘텐츠 폭이 큰 컬럼(상세구분, 단가, 금액)에 가중 분배
        slack = usable_cm - total
        slack_weights = [0.3, 0.2, 2.0, 0.2, 1.0, 0.0, 0.8, 0.5]
        sw_sum = sum(slack_weights)
        if sw_sum > 0:
            constrained = [
                w + slack * sw / sw_sum
                for w, sw in zip(constrained, slack_weights)
            ]

    return [Cm(round(w, 2)) for w in constrained]


# ─── 렌더링 함수들 ────────────────────────────────────────

def _render_logo(doc, brand: Brand, project_root: Path) -> None:
    if not brand.branding.logo_path:
        return
    logo_path = project_root / "brands" / brand.brand_id / brand.branding.logo_path
    if not logo_path.exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(str(logo_path), width=Cm(2.4))


def _render_title(doc, document: MembershipQuoteDocument,
                  scenario: MembershipScenario, brand: Brand) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)
    suffix = f" ({scenario.name})" if scenario.name else ""
    _add_paragraph(
        doc, f"{document.title}{suffix}",
        font=font, size_pt=15, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        color=primary, space_after_pt=1,
    )
    if scenario.subject:
        _add_paragraph(
            doc, scenario.subject,
            font=font, size_pt=8.5,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_after_pt=6,
        )


def _render_parties(doc, document: MembershipQuoteDocument, brand: Brand) -> None:
    """양측 발행 정보 (제휴사 | 회사)."""
    font = brand.branding.font_family
    primary_hex = brand.branding.colors.primary.lstrip("#")

    cp = document.counterparty
    sup = document.supplier
    if sup is None:
        # brand 정보로 자동 채움
        from .membership_models import MembershipParty
        cpn = brand.contact_person
        contact_label = None
        if cpn:
            contact_label = (
                f"{cpn.name} ({cpn.title})" if cpn.title else cpn.name
            )
        sup = MembershipParty(
            label="회사",
            name=brand.company.name_ko,
            address=brand.company.address,
            ceo=brand.company.ceo,
            contact=contact_label,
            contact_phone=(cpn.phone if cpn else None),
            contact_email=(cpn.email if cpn else None),
        )

    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(table, color="BFBFBF", size=4)

    # 헤더 (라벨)
    for col, party in enumerate([cp, sup]):
        c = table.rows[0].cells[col]
        _vcenter(c)
        _set_cell_bg(c, primary_hex)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(party.label)
        _apply_font(r, font, size_pt=8.5, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF))
    table.rows[0].height = Cm(0.45)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    # 정보
    for col, party in enumerate([cp, sup]):
        c = table.rows[1].cells[col]
        _vcenter(c)
        # 첫 줄: 회사명 (굵게)
        first = c.paragraphs[0]
        first.paragraph_format.space_before = Pt(0)
        first.paragraph_format.space_after = Pt(0)
        fr = first.add_run(party.name)
        _apply_font(fr, font, size_pt=9, bold=True)
        # 보조 정보
        info_lines = []
        if party.address:
            info_lines.append(f"주소 : {party.address}")
        if party.ceo:
            info_lines.append(f"대표이사 : {party.ceo}")
        if party.contact:
            info_lines.append(f"담당자 : {party.contact}")
        if party.contact_phone:
            info_lines.append(f"연락처 : {party.contact_phone}")
        if party.contact_email:
            info_lines.append(f"이메일 : {party.contact_email}")
        for line in info_lines:
            p = c.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(line)
            _apply_font(r, font, size_pt=7.5)


def _render_table_header_row(table, row_idx: int, headers: list[str], widths: list[Cm],
                             font: str, primary_hex: str) -> None:
    row = table.rows[row_idx]
    row.height = Cm(0.5)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for c_idx, (h, w) in enumerate(zip(headers, widths)):
        c = row.cells[c_idx]
        c.width = w
        _vcenter(c)
        _set_cell_bg(c, primary_hex)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        _apply_font(r, font, size_pt=8, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))


def _fill_cell(cell, text: str, *, font: str, size_pt: float = 9,
               bold: bool = False, align=WD_ALIGN_PARAGRAPH.LEFT,
               color: RGBColor | None = None, bg: str | None = None) -> None:
    _vcenter(cell)
    if bg:
        _set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if i == 0:
            r = p.add_run(ln)
        else:
            p2 = cell.add_paragraph()
            p2.alignment = align
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(0)
            r = p2.add_run(ln)
        _apply_font(r, font, size_pt=size_pt, bold=bold, color=color)


def _render_item_row(table, row_idx: int, item: MembershipLineItem, *,
                     widths: list[Cm], font: str) -> None:
    row = table.rows[row_idx]
    row.height = Cm(0.8)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    cells = row.cells

    # 상세 구분 (이름 + 보조 + 종량제 하위)
    name_text = item.name
    if item.name_detail:
        name_text += "\n" + item.name_detail
    if item.sub_items:
        # 종량제: 본행은 비우고 하위만 보여주거나, 본행 + 줄바꿈으로 하위 나열
        sub_lines = []
        for s in item.sub_items:
            if s.label:
                sub_lines.append(f"· {s.label}  {s.spec}")
            else:
                sub_lines.append(f"   {s.spec}")
        name_text += "\n" + "\n".join(sub_lines)
    _fill_cell(cells[0], name_text, font=font, size_pt=8.5,
               align=WD_ALIGN_PARAGRAPH.LEFT)

    # 기간
    _fill_cell(cells[1], item.billing_period or "-",
               font=font, size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 단가 (숫자 또는 텍스트)
    if item.unit_price is not None:
        up_text = f"₩{_format_won(item.unit_price)}"
    elif item.unit_price_text:
        up_text = item.unit_price_text
    else:
        up_text = "-"
    _fill_cell(cells[2], up_text, font=font, size_pt=8.5,
               align=WD_ALIGN_PARAGRAPH.RIGHT)

    # 할인율
    if item.discount_rate:
        dr_text = f"{int(item.discount_rate * 100)}%"
    else:
        dr_text = "-"
    _fill_cell(cells[3], dr_text, font=font, size_pt=8.5,
               align=WD_ALIGN_PARAGRAPH.CENTER)

    # 금액
    eff = item.effective_amount()
    if item.amount_text:
        amt_text = item.amount_text
    elif eff is not None:
        amt_text = f"₩{_format_won(eff)}"
    else:
        amt_text = "-"
    _fill_cell(cells[4], amt_text, font=font, size_pt=8.5, bold=True,
               align=WD_ALIGN_PARAGRAPH.RIGHT)

    # 비고
    _fill_cell(cells[5], item.notes or "-",
               font=font, size_pt=8.5, align=WD_ALIGN_PARAGRAPH.LEFT)


def _render_section_table(doc, section: MembershipSection, brand: Brand) -> None:
    """한 구분의 표를 렌더링. 구분 컬럼은 좌측에 병합."""
    font = brand.branding.font_family
    primary_hex = brand.branding.colors.primary.lstrip("#")
    light_bg = "F0F2F8"

    # 표 컬럼: [구분 | 분류 | 상세구분 | 기간 | 단가 | 할인율 | 금액 | 비고]
    # 단순화를 위해 [구분, 분류, 상세 구분 | 기간 | 단가 | 할인율 | 금액 | 비고] 8개
    # 첫 컬럼(구분)은 첫 행에만 표시하고 나머지는 수직 병합

    # 행 수 계산: 헤더(1) + 각 분류별 [항목들 + (옵션) 합계 1행] + 섹션 합계 1행
    body_rows = 0
    for cat in section.categories:
        body_rows += len(cat.items)
        if cat.show_subtotal:
            body_rows += 1
    if section.show_section_total:
        body_rows += 1
    total_rows = 1 + body_rows  # +1 헤더

    cols_n = 8
    table = doc.add_table(rows=total_rows, cols=cols_n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(table, color="BFBFBF", size=4)

    headers = ["구분", "분류", "상세 구분", "기간", "단가", "할인율", "금액", "비고"]
    # 콘텐츠 길이 + 중요도 기반 유동 컬럼 너비
    widths = _compute_section_widths(section)
    _render_table_header_row(table, 0, headers, widths, font, primary_hex)

    # 본문 채우기
    r = 1
    section_first_row = r  # 구분 컬럼 병합 시작점

    for cat in section.categories:
        cat_first_row = r  # 분류 컬럼 병합 시작점

        for item in cat.items:
            # 컬럼 0,1 은 분류 row에서만 텍스트 (병합 후 빈 값)
            _fill_cell(table.cell(r, 0), "", font=font)
            _fill_cell(table.cell(r, 1), "", font=font)
            # 컬럼 2~7: 항목 데이터
            _render_item_row_in_table(table, r, item, widths, font)
            r += 1

        # 분류 컬럼 텍스트 (병합 후의 첫 행 = cat_first_row)
        _fill_cell(table.cell(cat_first_row, 1), cat.name,
                   font=font, size_pt=8, bold=True,
                   align=WD_ALIGN_PARAGRAPH.CENTER, bg=light_bg)
        # 분류 컬럼 수직 병합 (옵션 분류는 합계 행 없을 수 있음)
        cat_last_row = r - 1  # 마지막 항목 row
        if cat_first_row != cat_last_row:
            _merge_vertical(table, 1, cat_first_row, cat_last_row)

        # 분류 합계 행 (col 2~5 병합으로 라벨 넓게, col 6 금액, col 7 비고)
        if cat.show_subtotal:
            subtotal = category_subtotal(cat)
            label = cat.subtotal_label or f"{cat.name} 합계"
            subtotal_row = table.rows[r]
            subtotal_row.height = Cm(0.45)
            subtotal_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            _fill_cell(table.cell(r, 0), "", font=font)
            _fill_cell(table.cell(r, 1), "", font=font)
            label_cell = table.cell(r, 2).merge(table.cell(r, 5))
            _fill_cell(label_cell, label, font=font, size_pt=8, bold=True,
                       align=WD_ALIGN_PARAGRAPH.RIGHT, bg=light_bg)
            _fill_cell(table.cell(r, 6),
                       _money(subtotal) if subtotal else "-",
                       font=font, size_pt=8.5, bold=True,
                       align=WD_ALIGN_PARAGRAPH.RIGHT, bg=light_bg)
            _fill_cell(table.cell(r, 7), "-", font=font, size_pt=8,
                       align=WD_ALIGN_PARAGRAPH.CENTER, bg=light_bg)
            r += 1

    # 구분 컬럼 텍스트 + 수직 병합 (섹션 합계 행 직전까지)
    _fill_cell(table.cell(section_first_row, 0), section.name,
               font=font, size_pt=8.5, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, bg=light_bg)
    section_body_last = r - 1
    if section_first_row != section_body_last:
        _merge_vertical(table, 0, section_first_row, section_body_last)

    # 섹션 예상 총 금액 행 (col 2~7 가로 병합으로 한 줄에 표시)
    if section.show_section_total:
        by_period = section_subtotals_by_period(section)
        parts = [f"{period}: {_money(amt)}" for period, amt in by_period.items()]
        total_text = "  /  ".join(parts) if parts else "-"

        section_total_row = table.rows[r]
        section_total_row.height = Cm(0.5)
        section_total_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        _fill_cell(table.cell(r, 0), "", font=font, bg=light_bg)
        _fill_cell(table.cell(r, 1), "예상 총 금액",
                   font=font, size_pt=8, bold=True,
                   align=WD_ALIGN_PARAGRAPH.CENTER,
                   bg=brand.branding.colors.primary.lstrip("#"),
                   color=RGBColor(0xFF, 0xFF, 0xFF))
        # col 2~7 가로 병합해서 넓은 한 셀로
        merged_cell = table.cell(r, 2).merge(table.cell(r, 7))
        _fill_cell(merged_cell, total_text,
                   font=font, size_pt=9, bold=True,
                   align=WD_ALIGN_PARAGRAPH.LEFT, bg=light_bg)
        r += 1


def _render_item_row_in_table(table, row_idx: int, item: MembershipLineItem,
                              widths: list[Cm], font: str) -> None:
    """8컬럼 표에서 상세구분 시작 컬럼(idx=2)부터 채움."""
    row = table.rows[row_idx]
    row.height = Cm(0.45)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    # 컬럼 2: 상세 구분 (이름 + 보조 + 종량제 하위) — sub_items 만 더 작은 폰트로
    name_cell = table.cell(row_idx, 2)
    _vcenter(name_cell)
    name_p = name_cell.paragraphs[0]
    name_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    name_p.paragraph_format.space_before = Pt(0)
    name_p.paragraph_format.space_after = Pt(0)
    nr = name_p.add_run(item.name)
    _apply_font(nr, font, size_pt=7.5)
    if item.name_detail:
        for ln in item.name_detail.split("\n"):
            p_detail = name_cell.add_paragraph()
            p_detail.paragraph_format.space_before = Pt(0)
            p_detail.paragraph_format.space_after = Pt(0)
            r_detail = p_detail.add_run(ln)
            _apply_font(r_detail, font, size_pt=7.5)
    if item.sub_items:
        sub_lines = []
        for s in item.sub_items:
            if s.label:
                sub_lines.append(f"· {s.label}  {s.spec}")
            elif sub_lines:
                sub_lines[-1] += f"  {s.spec}"
            else:
                sub_lines.append(f"   {s.spec}")
        for ln in sub_lines:
            p_sub = name_cell.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(0)
            r_sub = p_sub.add_run(ln)
            _apply_font(r_sub, font, size_pt=6.5)

    # 컬럼 3: 기간
    _fill_cell(table.cell(row_idx, 3), item.billing_period or "-",
               font=font, size_pt=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 컬럼 4: 단가
    if item.unit_price is not None:
        up_text = _money(item.unit_price)
        up_align = WD_ALIGN_PARAGRAPH.RIGHT
    elif item.unit_price_text:
        up_text = item.unit_price_text
        up_align = WD_ALIGN_PARAGRAPH.LEFT
    else:
        up_text = "-"
        up_align = WD_ALIGN_PARAGRAPH.CENTER
    _fill_cell(table.cell(row_idx, 4), up_text,
               font=font, size_pt=7.5, align=up_align)

    # 컬럼 5: 할인율
    dr_text = f"{int(item.discount_rate * 100)}%" if item.discount_rate else "-"
    _fill_cell(table.cell(row_idx, 5), dr_text,
               font=font, size_pt=7.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 컬럼 6: 금액
    eff = item.effective_amount()
    if item.amount_text:
        amt_text = item.amount_text
    elif eff is not None:
        amt_text = _money(eff)
    else:
        amt_text = "-"
    _fill_cell(table.cell(row_idx, 6), amt_text,
               font=font, size_pt=8, bold=True,
               align=WD_ALIGN_PARAGRAPH.RIGHT)

    # 컬럼 7: 비고
    _fill_cell(table.cell(row_idx, 7), item.notes or "-",
               font=font, size_pt=7.5, align=WD_ALIGN_PARAGRAPH.LEFT)


def _render_unit_notice(doc, document: MembershipQuoteDocument, brand: Brand) -> None:
    font = brand.branding.font_family
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(f"아래와 같이 견적하오니 참조하시기 바랍니다.    {document.unit_notice}")
    _apply_font(r, font, size_pt=7.5, color=RGBColor(0x77, 0x77, 0x77))


def _render_grand_total(doc, scenario: MembershipScenario,
                        brand: Brand, vat_rate: float = 0.10) -> None:
    """시나리오 전체의 공급가액·부가세·합계 금액. 기간별 breakdown 유지."""
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)
    by_period = scenario_grand_total_by_period(scenario)
    if not by_period:
        return

    def _inline(period_map: dict) -> str:
        return "  /  ".join(f"{p}: {_money(v)}" for p, v in period_map.items())

    vat_by_period = {p: round(v * vat_rate) for p, v in by_period.items()}
    total_by_period = {p: by_period[p] + vat_by_period[p] for p in by_period}

    pct = int(vat_rate * 100)
    rows = [
        ("공급가액", _inline(by_period), False),
        (f"부가세 ({pct}%)", _inline(vat_by_period), False),
        ("합계 금액", _inline(total_by_period), True),
    ]

    for idx, (label, value, highlight) in enumerate(rows):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(4 if idx == 0 else 0)
        p.paragraph_format.space_after = Pt(0)
        r_label = p.add_run(f"{label}  ")
        _apply_font(r_label, font, size_pt=9, bold=True,
                    color=primary if highlight else None)
        r_value = p.add_run(value)
        _apply_font(r_value, font,
                    size_pt=10 if highlight else 9,
                    bold=True,
                    color=primary if highlight else None)


def _render_remarks(doc, document: MembershipQuoteDocument, brand: Brand) -> None:
    if not document.remarks:
        return
    font = brand.branding.font_family
    _add_paragraph(doc, "※ 기타 안내", font=font, size_pt=8, bold=True,
                   space_before_pt=4, space_after_pt=0)
    for i, line in enumerate(document.remarks, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(0.2)
        r = p.add_run(f"{i}. {line}")
        _apply_font(r, font, size_pt=7.5)


def _render_date(doc, document: MembershipQuoteDocument, brand: Brand) -> None:
    font = brand.branding.font_family
    _add_paragraph(doc, document.issued_date.strftime("%Y년 %m월 %d일"),
                   font=font, size_pt=9,
                   alignment=WD_ALIGN_PARAGRAPH.CENTER,
                   space_before_pt=4, space_after_pt=2)


def _add_page_break(doc) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _add_tiny_spacer(doc) -> None:
    """섹션 사이 아주 작은 간격 (2pt)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("")
    r.font.size = Pt(2)


# ─── 메인 진입점 ───────────────────────────────────────────

def render_membership_docx(brand: Brand, document: MembershipQuoteDocument,
                           project_root: Path, output_path: Path) -> Path:
    labels = load_labels(project_root)
    vat_rate = labels.quote.vat_rate

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(0.9)
        section.bottom_margin = Cm(0.9)
        section.left_margin = Cm(1.1)
        section.right_margin = Cm(1.1)

    for s_idx, scenario in enumerate(document.scenarios):
        if s_idx > 0:
            _add_page_break(doc)
        _render_logo(doc, brand, project_root)
        _render_title(doc, document, scenario, brand)
        _render_parties(doc, document, brand)
        _render_unit_notice(doc, document, brand)
        for sec_i, sec in enumerate(scenario.sections):
            _render_section_table(doc, sec, brand)
            if sec_i < len(scenario.sections) - 1:
                _add_tiny_spacer(doc)
        if scenario.show_grand_total:
            _render_grand_total(doc, scenario, brand, vat_rate)
        _render_date(doc, document, brand)
        _render_remarks(doc, document, brand)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
