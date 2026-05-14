from __future__ import annotations

from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

from .labels import DocumentLabels, load_labels
from .loader import load_clause, render_clause_body
from .models import Brand, QuoteDocument, ensure_totals


def _hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_cell_bg(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tc_pr.append(shd)


def _vcenter(cell) -> None:
    """셀의 수직 정렬을 가운데로 (XML 직접 조작 - LibreOffice 변환 호환)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:vAlign")):
        tc_pr.remove(old)
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), "center")
    tc_pr.append(v)


def _zero_cell_lr_margin(cell) -> None:
    """셀의 좌우 padding(margin)을 0으로. 페이지 여백과 셀 내용을 정렬할 때 사용."""
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcMar"))
    if existing is None:
        existing = OxmlElement("w:tcMar")
        tc_pr.append(existing)
    for direction in ("left", "right"):
        for old in existing.findall(qn(f"w:{direction}")):
            existing.remove(old)
        m = OxmlElement(f"w:{direction}")
        m.set(qn("w:w"), "0")
        m.set(qn("w:type"), "dxa")
        existing.append(m)


def _set_table_borders(table, color: str = "BFBFBF", size: int = 4) -> None:
    """표 전체에 격자선을 추가. size는 1/8 pt 단위 (4 = 0.5pt)."""
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


def _add_paragraph(doc, text: str, *, font: str, size_pt: float = 10,
                   bold: bool = False, alignment=None, color: RGBColor | None = None,
                   space_after_pt: float = 6, space_before_pt: float = 0):
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.space_before = Pt(space_before_pt)
    run = p.add_run(text)
    _apply_font(run, font, size_pt=size_pt, bold=bold, color=color)
    return p


def _format_money(amount: float, currency: str = "KRW") -> str:
    if currency == "KRW":
        return f"₩{int(round(amount)):,}"
    return f"{amount:,.2f} {currency}"


def _render_logo(doc, brand: Brand, project_root: Path) -> None:
    """브랜드 로고가 있으면 문서 상단에 삽입. 없으면 조용히 건너뜀."""
    if not brand.branding.logo_path:
        return
    logo_full_path = project_root / "brands" / brand.brand_id / brand.branding.logo_path
    if not logo_full_path.exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(str(logo_full_path), width=Cm(3.2))


def _render_header(doc, brand: Brand, document: QuoteDocument,
                   labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)

    title_text = labels.quote.title if document.document_type == "quote" else labels.contract.title
    _add_paragraph(doc, title_text, font=font, size_pt=20, bold=True,
                   alignment=WD_ALIGN_PARAGRAPH.CENTER, color=primary, space_after_pt=4)

    info_table = doc.add_table(rows=1, cols=2)
    info_table.autofit = False
    info_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    info_table.columns[0].width = Cm(7.5)
    info_table.columns[1].width = Cm(9.5)

    # 우측 inner 표(5행 × 0.6cm = 3.0cm)와 맞춰 외부 행 높이 명시 → 좌측도 vAlign center 동작
    info_table.rows[0].height = Cm(3.0)
    info_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    left, right = info_table.rows[0].cells
    left.width = Cm(7.5)
    right.width = Cm(9.5)
    _vcenter(left)
    _vcenter(right)
    # 좌측 셀 padding 제거 → "(주)소프트먼트" 가 페이지 여백과 정확히 정렬
    _zero_cell_lr_margin(left)

    p = left.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0)
    run = p.add_run(brand.company.name_ko)
    _apply_font(run, font, size_pt=13, bold=True, color=primary)

    info_lines = [
        f"사업자등록번호: {brand.company.registration_number}",
        f"대표자: {brand.company.ceo}",
    ]
    if brand.company.address:
        info_lines.append(brand.company.address)

    for line in info_lines:
        para = left.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.left_indent = Cm(0)
        r = para.add_run(line)
        _apply_font(r, font, size_pt=8.5)

    info_rows = [
        ("발행일", document.issued_date.isoformat()),
    ]
    if document.valid_until:
        info_rows.append(("유효기간", document.valid_until.isoformat() + " 까지"))
    if document.effective_date:
        info_rows.append(("계약시작일", document.effective_date.isoformat()))

    cp = brand.contact_person
    if cp:
        담당자_표기 = cp.name + (f" ({cp.title})" if cp.title else "")
        info_rows.append(("담당자", 담당자_표기))
        info_rows.append(("연락처", cp.phone or brand.company.phone or "-"))
        info_rows.append(("Email", cp.email or brand.company.email or "-"))
    else:
        info_rows.append(("연락처", brand.company.phone or "-"))
        info_rows.append(("Email", brand.company.email or "-"))

    inner = right.add_table(rows=len(info_rows), cols=2)
    inner.autofit = True
    for idx, (label, value) in enumerate(info_rows):
        row_obj = inner.rows[idx]
        # row 높이 명시 (LibreOffice가 vAlign 인식하도록).
        # AT_LEAST → 이메일 등 긴 텍스트는 줄바꿈하면서 행이 자동으로 늘어남 (잘림 방지)
        row_obj.height = Cm(0.6)
        row_obj.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        lc, vc = row_obj.cells
        _vcenter(lc)
        _vcenter(vc)
        _set_cell_bg(lc, "F0F2F8")
        lp = lc.paragraphs[0]
        lp.paragraph_format.space_before = Pt(0)
        lp.paragraph_format.space_after = Pt(0)
        lr = lp.add_run(label)
        _apply_font(lr, font, size_pt=8.5, bold=True)
        vp = vc.paragraphs[0]
        vp.paragraph_format.space_before = Pt(0)
        vp.paragraph_format.space_after = Pt(0)
        vr = vp.add_run(value)
        _apply_font(vr, font, size_pt=8.5)


def _render_counterparty(doc, brand: Brand, document: QuoteDocument,
                         labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)

    _add_paragraph(doc, labels.quote.labels.counterparty_section,
                   font=font, size_pt=10, bold=True,
                   color=primary, space_before_pt=18, space_after_pt=2)

    cp = document.counterparty
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    _vcenter(cell)
    _set_cell_bg(cell, "FAFBFD")

    lines = [(cp.name, True, 11)]
    if cp.registration_number:
        lines.append((f"사업자등록번호: {cp.registration_number}", False, 8.5))
    if cp.address:
        lines.append((cp.address, False, 8.5))
    contact_bits = []
    if cp.contact_name:
        contact_bits.append(f"담당: {cp.contact_name}" + (f" ({cp.contact_title})" if cp.contact_title else ""))
    if cp.email:
        contact_bits.append(f"Email: {cp.email}")
    if contact_bits:
        lines.append(("  |  ".join(contact_bits), False, 8.5))

    first = cell.paragraphs[0]
    first.paragraph_format.space_after = Pt(1)
    fr = first.add_run(lines[0][0])
    _apply_font(fr, font, size_pt=lines[0][2], bold=lines[0][1])

    for text, bold, size in lines[1:]:
        para = cell.add_paragraph()
        para.paragraph_format.space_after = Pt(0)
        r = para.add_run(text)
        _apply_font(r, font, size_pt=size, bold=bold)

    _add_paragraph(doc, f"{labels.quote.labels.subject_prefix}: {document.subject}",
                   font=font, size_pt=10.5, bold=True,
                   space_before_pt=14, space_after_pt=2)


def _render_line_items(doc, brand: Brand, document: QuoteDocument,
                       labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary_hex = brand.branding.colors.primary.lstrip("#")
    ql = labels.quote

    if ql.labels.vat_separate_notice:
        vat_label_p = doc.add_paragraph()
        vat_label_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        vat_label_p.paragraph_format.space_before = Pt(0)
        vat_label_p.paragraph_format.space_after = Pt(0)
        vat_run = vat_label_p.add_run(ql.labels.vat_separate_notice)
        _apply_font(vat_run, font, size_pt=8.5, bold=True,
                    color=RGBColor(0x99, 0x33, 0x33))

    th = ql.table_headers
    headers = [th.name, th.description, th.qty, th.period, th.unit_price, th.amount, th.notes]
    table = doc.add_table(rows=1 + len(document.line_items), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(table, color="9FB0CC", size=4)

    widths = [Cm(3.2), Cm(3.0), Cm(1.0), Cm(1.2), Cm(2.3), Cm(2.7), Cm(4.0)]

    header_row = table.rows[0]
    # 헤더 행 높이 명시 → vAlign center 강제
    header_row.height = Cm(0.8)
    header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for idx, (header, width) in enumerate(zip(headers, widths)):
        cell = header_row.cells[idx]
        cell.width = width
        _vcenter(cell)
        _set_cell_bg(cell, primary_hex)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(header)
        _apply_font(r, font, size_pt=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for r_idx, item in enumerate(document.line_items, start=1):
        row_obj = table.rows[r_idx]
        # 명시적 row 높이 부여 (vAlign center가 LibreOffice에서 동작하도록)
        row_obj.height = Cm(1.0)
        row_obj.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row = row_obj.cells
        qty_text = f"{item.qty:g}" if item.qty is not None else ""
        period_text = f"{item.period:g}" if item.period is not None else ""
        for c_idx, (text, align) in enumerate([
            (item.name, WD_ALIGN_PARAGRAPH.LEFT),
            (item.description or "", WD_ALIGN_PARAGRAPH.LEFT),
            (qty_text, WD_ALIGN_PARAGRAPH.CENTER),
            (period_text, WD_ALIGN_PARAGRAPH.CENTER),
            (_format_money(item.unit_price, item.currency), WD_ALIGN_PARAGRAPH.RIGHT),
            (_format_money(item.amount, item.currency), WD_ALIGN_PARAGRAPH.RIGHT),
            (item.notes or "", WD_ALIGN_PARAGRAPH.LEFT),
        ]):
            cell = row[c_idx]
            cell.width = widths[c_idx]
            _vcenter(cell)
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            _apply_font(run, font, size_pt=9)


def _render_totals(doc, brand: Brand, document: QuoteDocument,
                   labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)
    t = document.totals
    currency = t.currency

    ql = labels.quote.labels
    rows = [
        (ql.subtotal, _format_money(t.subtotal, currency), False),
        (f"{ql.vat} ({int(t.vat_rate * 100)}%)", _format_money(t.vat, currency), False),
        (ql.total, _format_money(t.total, currency), True),
    ]

    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    _set_table_borders(table, color="9FB0CC", size=4)
    for idx, (label, value, highlight) in enumerate(rows):
        row_obj = table.rows[idx]
        row_obj.height = Cm(0.75 if highlight else 0.65)
        row_obj.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        lc, vc = row_obj.cells
        lc.width = Cm(4)
        vc.width = Cm(4)
        _vcenter(lc)
        _vcenter(vc)
        if highlight:
            _set_cell_bg(lc, brand.branding.colors.primary.lstrip("#"))
            _set_cell_bg(vc, brand.branding.colors.primary.lstrip("#"))
        lp = lc.paragraphs[0]
        lp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        lp.paragraph_format.space_before = Pt(0)
        lp.paragraph_format.space_after = Pt(0)
        lr = lp.add_run(label)
        _apply_font(lr, font, size_pt=10, bold=highlight,
                    color=RGBColor(0xFF, 0xFF, 0xFF) if highlight else None)
        vp = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        vp.paragraph_format.space_before = Pt(0)
        vp.paragraph_format.space_after = Pt(0)
        vr = vp.add_run(value)
        _apply_font(vr, font, size_pt=11 if highlight else 10, bold=highlight,
                    color=RGBColor(0xFF, 0xFF, 0xFF) if highlight else None)


def _render_clauses(doc, brand: Brand, document: QuoteDocument, project_root: Path) -> None:
    if not document.clauses:
        return
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)

    _add_paragraph(doc, "약관 및 조건", font=font, size_pt=12, bold=True,
                   color=primary, space_after_pt=6)

    for number, clause_ref in enumerate(document.clauses, start=1):
        clause = load_clause(project_root, document.document_type, clause_ref.id)
        body = render_clause_body(clause, clause_ref, number)

        title_p = doc.add_paragraph()
        title_p.paragraph_format.space_before = Pt(4)
        title_p.paragraph_format.space_after = Pt(2)
        tr = title_p.add_run(f"제 {number}조  {clause.title}")
        _apply_font(tr, font, size_pt=10.5, bold=True)

        body_p = doc.add_paragraph()
        body_p.paragraph_format.space_after = Pt(6)
        br = body_p.add_run(body.strip())
        _apply_font(br, font, size_pt=10)


def _render_etc_notice(doc, brand: Brand, document: QuoteDocument,
                       labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)
    ql = labels.quote

    bullets: list[str] = []
    # 1) 유효기간 자동 문구 (템플릿 비어있으면 생략)
    if (document.document_type == "quote"
            and document.valid_until
            and ql.auto_notices.validity_template):
        valid_str = document.valid_until.strftime("%Y년 %m월 %d일")
        bullets.append(ql.auto_notices.validity_template.format(valid_until=valid_str))
    # 2) 문서별 추가 비고 (data JSON의 notes)
    if document.notes:
        bullets.append(document.notes)
    # 3) 입금 계좌 자동 문구 (템플릿 비어있으면 생략)
    if brand.bank_account and ql.auto_notices.bank_account_template:
        ba = brand.bank_account
        bullets.append(ql.auto_notices.bank_account_template.format(
            bank=ba.bank,
            account_number=ba.account_number,
            account_holder=ba.account_holder or "",
        ))
    # 4) 항상 표시되는 고정 문구
    for static in ql.static_notices:
        if static:
            bullets.append(static)

    if not bullets:
        return

    _add_paragraph(doc, ql.labels.etc_notice_section,
                   font=font, size_pt=10, bold=True,
                   color=primary, space_before_pt=10, space_after_pt=2)
    for text in bullets:
        para = doc.add_paragraph()
        para.paragraph_format.space_after = Pt(1)
        para.paragraph_format.left_indent = Cm(0.3)
        r = para.add_run(f"• {text}")
        _apply_font(r, font, size_pt=9.5)


def _render_signature(doc, brand: Brand, document: QuoteDocument,
                      labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)

    _add_paragraph(doc, document.issued_date.strftime("%Y년 %m월 %d일"),
                   font=font, size_pt=11, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                   space_before_pt=28, space_after_pt=10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(brand.company.name_ko)
    _apply_font(r, font, size_pt=14, bold=True, color=primary)

    if brand.footer_text:
        _add_paragraph(doc, brand.footer_text, font=font, size_pt=8.5,
                       alignment=WD_ALIGN_PARAGRAPH.CENTER,
                       color=RGBColor(0x88, 0x88, 0x88),
                       space_before_pt=8, space_after_pt=0)


def _render_contract_title(doc, brand: Brand, document: QuoteDocument,
                           labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)
    _add_paragraph(doc, labels.contract.title, font=font, size_pt=28, bold=True,
                   alignment=WD_ALIGN_PARAGRAPH.CENTER, color=primary,
                   space_after_pt=6)
    if document.subject:
        _add_paragraph(doc, f"〈 {document.subject} 〉",
                       font=font, size_pt=13, bold=True,
                       alignment=WD_ALIGN_PARAGRAPH.CENTER,
                       space_after_pt=14)


def _render_contract_preamble(doc, brand: Brand, document: QuoteDocument,
                              labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    preamble = labels.contract.preamble_template.format(
        supplier_name=brand.company.name_ko,
        counterparty_name=document.counterparty.name,
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.line_spacing = 1.6
    r = p.add_run(preamble)
    _apply_font(r, font, size_pt=11)


def _render_contract_parties(doc, brand: Brand, document: QuoteDocument,
                             labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary_hex = brand.branding.colors.primary.lstrip("#")
    cl = labels.contract.labels

    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for col, label in enumerate([cl.party_a, cl.party_b]):
        cell = table.rows[0].cells[col]
        _set_cell_bg(cell, primary_hex)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(label)
        _apply_font(r, font, size_pt=11, bold=True,
                    color=RGBColor(0xFF, 0xFF, 0xFF))

    cp = document.counterparty
    parties = [
        [
            ("회사명", brand.company.name_ko),
            ("사업자등록번호", brand.company.registration_number),
            ("대표자", brand.company.ceo),
            ("주소", brand.company.address or "-"),
        ],
        [
            ("회사명", cp.name),
            ("사업자등록번호", cp.registration_number or "-"),
            ("대표자", cp.ceo or "-"),
            ("주소", cp.address or "-"),
        ],
    ]
    for col, lines in enumerate(parties):
        cell = table.rows[1].cells[col]
        _set_cell_bg(cell, "FAFBFD")
        for idx, (label, value) in enumerate(lines):
            p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            r1 = p.add_run(f"{label}: ")
            _apply_font(r1, font, size_pt=9, bold=True)
            r2 = p.add_run(value or "-")
            _apply_font(r2, font, size_pt=9)
    doc.add_paragraph()


def _render_contract_overview(doc, brand: Brand, document: QuoteDocument,
                              labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    cl = labels.contract.labels
    rows = []
    if document.effective_date:
        rows.append((cl.effective_date, document.effective_date.strftime("%Y년 %m월 %d일")))
    if document.contract_term:
        rows.append((cl.contract_term, document.contract_term))
    if document.totals:
        rows.append((cl.amount_pre_vat,
                     _format_money(document.totals.subtotal, document.totals.currency)))
        rows.append((cl.amount_with_vat,
                     _format_money(document.totals.total, document.totals.currency)))
    if not rows:
        return

    primary_hex = brand.branding.colors.primary.lstrip("#")
    _add_paragraph(doc, cl.overview_section, font=font, size_pt=12, bold=True,
                   color=_hex_to_rgb(brand.branding.colors.primary),
                   space_after_pt=4)

    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, (label, value) in enumerate(rows):
        lc, vc = table.rows[idx].cells
        _set_cell_bg(lc, "F0F2F8")
        lp = lc.paragraphs[0]
        lr = lp.add_run(label)
        _apply_font(lr, font, size_pt=10, bold=True)
        vp = vc.paragraphs[0]
        vr = vp.add_run(value)
        _apply_font(vr, font, size_pt=10)
    doc.add_paragraph()


def _render_contract_signature(doc, brand: Brand, document: QuoteDocument,
                               labels: DocumentLabels) -> None:
    font = brand.branding.font_family
    primary = _hex_to_rgb(brand.branding.colors.primary)

    doc.add_paragraph()
    _add_paragraph(doc, document.issued_date.strftime("%Y년 %m월 %d일"),
                   font=font, size_pt=12, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                   space_after_pt=20)

    table = doc.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for col, label in enumerate(["갑", "을"]):
        cell = table.rows[0].cells[col]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(label)
        _apply_font(r, font, size_pt=14, bold=True, color=primary)

    cp = document.counterparty
    for col, name in enumerate([brand.company.name_ko, cp.name]):
        cell = table.rows[1].cells[col]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(name)
        _apply_font(r, font, size_pt=13, bold=True)

    for col, addr in enumerate([brand.company.address or "", cp.address or ""]):
        cell = table.rows[2].cells[col]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(addr)
        _apply_font(r, font, size_pt=9)

    rep_a = f"{brand.signature.signer_title}  {brand.signature.signer_name}  (인)"
    cp_signer_title = "대표이사"
    cp_signer_name = cp.ceo or cp.contact_name or "_______________"
    rep_b = f"{cp_signer_title}  {cp_signer_name}  (인)"
    for col, rep in enumerate([rep_a, rep_b]):
        cell = table.rows[3].cells[col]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        r = p.add_run(rep)
        _apply_font(r, font, size_pt=11, bold=True)


def render_docx(brand: Brand, document: QuoteDocument, project_root: Path,
                output_path: Path) -> Path:
    labels = load_labels(project_root)
    # 누락된 totals 필드를 자동 계산 (공급가액 → 부가세 → 합계)
    ensure_totals(document, labels.quote.vat_rate)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    _render_logo(doc, brand, project_root)

    if document.document_type == "contract":
        _render_contract_title(doc, brand, document, labels)
        _render_contract_preamble(doc, brand, document, labels)
        _render_contract_parties(doc, brand, document, labels)
        _render_contract_overview(doc, brand, document, labels)
        _render_clauses(doc, brand, document, project_root)
        _render_contract_signature(doc, brand, document, labels)
    else:
        _render_header(doc, brand, document, labels)
        _render_counterparty(doc, brand, document, labels)
        _render_line_items(doc, brand, document, labels)
        _render_totals(doc, brand, document, labels)
        _render_etc_notice(doc, brand, document, labels)
        _render_clauses(doc, brand, document, project_root)
        _render_signature(doc, brand, document, labels)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
