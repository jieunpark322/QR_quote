"""견적서 자동 생성 웹 인터페이스 (Streamlit).

`python -m src.cli web` 로 실행하면 브라우저가 자동으로 열립니다.

페이지 구성:
  📋 견적서 작성       — 폼 입력 → DOCX/PDF 다운로드
  📦 카탈로그 관리     — 상품 추가/수정/삭제
  ⚙ 설정              — 브랜드 정보, 양식 라벨/문구 편집
"""
from __future__ import annotations

import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

# Streamlit이 이 파일을 단독 스크립트로 실행하므로,
# 상위 폴더(src의 부모)를 path에 추가해 `src.xxx` 형태로 import 가능하게 함
_PARENT = Path(__file__).resolve().parent.parent
if str(_PARENT) not in sys.path:
    sys.path.insert(0, str(_PARENT))

import pandas as pd
import streamlit as st

from src.labels import DocumentLabels, load_labels
from src.loader import load_brand
from src.models import (
    BankAccount,
    Brand,
    BrandColors,
    BrandingAssets,
    CompanyInfo,
    ContactPerson,
    Counterparty,
    LineItem,
    QuoteDocument,
    SignatureInfo,
    Totals,
)
from src.pdf_converter import convert_docx_to_pdf, find_soffice
from src.renderer import render_docx


def _detect_project_root() -> Path:
    """프로젝트 루트 자동 감지 (CLI/로컬/클라우드 환경 모두 호환)."""
    # 1) CLI가 환경변수로 명시했으면 우선
    if env := os.environ.get("CONTRACT_SYSTEM_ROOT"):
        return Path(env).resolve()
    # 2) 현재 파일이 <root>/src/webapp.py 형태인지 확인
    here = Path(__file__).resolve().parent
    if here.name == "src" and (here.parent / "brands").exists():
        return here.parent
    # 3) 작업 디렉터리에 brands/ 가 있으면 거기
    cwd = Path.cwd().resolve()
    if (cwd / "brands").exists():
        return cwd
    # 4) 최종 fallback
    return here.parent


PROJECT_ROOT = _detect_project_root()


# ═════════════════════════════════════════════════════════════
# 데이터 로더 / 저장 헬퍼
# ═════════════════════════════════════════════════════════════

@st.cache_data
def _load_products() -> list[dict]:
    path = PROJECT_ROOT / "catalog" / "products.json"
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8")).get("products", [])


def _save_products(products: list[dict]) -> None:
    path = PROJECT_ROOT / "catalog" / "products.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"products": products}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    _load_products.clear()


@st.cache_data
def _list_brands() -> list[str]:
    brands_dir = PROJECT_ROOT / "brands"
    if not brands_dir.exists():
        return ["softment"]
    return sorted([d.name for d in brands_dir.iterdir()
                   if d.is_dir() and (d / "brand.json").exists()])


def _save_brand(brand_id: str, brand_dict: dict) -> None:
    path = PROJECT_ROOT / "brands" / brand_id / "brand.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(brand_dict, ensure_ascii=False, indent=2),
                    encoding="utf-8")


def _save_labels(labels_dict: dict) -> None:
    path = PROJECT_ROOT / "config" / "labels.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(labels_dict, ensure_ascii=False, indent=2),
                    encoding="utf-8")


# ═════════════════════════════════════════════════════════════
# 페이지 1: 견적서 작성
# ═════════════════════════════════════════════════════════════

ITEM_COLUMNS = ["품목", "설명", "수량", "기간(회)", "단가", "비고"]


def _empty_items_df() -> pd.DataFrame:
    return pd.DataFrame(columns=ITEM_COLUMNS).astype({
        "품목": "string",
        "설명": "string",
        "수량": "Int64",
        "기간(회)": "Int64",
        "단가": "Int64",
        "비고": "string",
    })


def _ensure_items_state():
    if "items_df" not in st.session_state:
        st.session_state.items_df = _empty_items_df()


def _add_catalog_row(product: dict):
    new_row = {
        "품목": product["name"],
        "설명": product.get("description", ""),
        "수량": 1,
        "기간(회)": 1,
        "단가": int(product.get("unit_price", 0)),
        "비고": "",
    }
    df = st.session_state.items_df
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True
    )


def _add_blank_row():
    df = st.session_state.items_df
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([{c: None for c in ITEM_COLUMNS}])],
        ignore_index=True,
    )


def _reset_items():
    st.session_state.items_df = _empty_items_df()


def _row_amount(row) -> int:
    qty = row.get("수량")
    period = row.get("기간(회)")
    price = row.get("단가")
    q = qty if pd.notna(qty) and qty else 1
    p = period if pd.notna(period) and period else 1
    pr = price if pd.notna(price) and price else 0
    try:
        return int(q) * int(p) * int(pr)
    except (TypeError, ValueError):
        return 0


def render_quote_page():
    _ensure_items_state()
    labels = load_labels(PROJECT_ROOT)
    products = _load_products()
    brand_ids = _list_brands()
    soffice_available = find_soffice() is not None

    # ─── 사이드바: 발행 정보 + 카탈로그 ───
    with st.sidebar:
        st.divider()
        st.header("⚙ 발행 정보")
        brand_id = st.selectbox("브랜드", brand_ids,
                                index=0 if brand_ids else None,
                                disabled=len(brand_ids) <= 1)
        issued_date = st.date_input("발행일", value=date.today())
        valid_days = st.number_input(
            "유효기간 (일)",
            value=labels.quote.default_validity_days,
            min_value=1, max_value=365,
        )
        valid_until = issued_date + timedelta(days=int(valid_days))
        st.caption(f"📅 유효기간: **{valid_until.isoformat()}** 까지")

        if not soffice_available:
            st.divider()
            st.warning("⚠ LibreOffice 미감지 — PDF 생성이 불가합니다.")

    # ─── 메인 ───
    st.title("📋 견적서 자동 생성")
    st.caption("폼을 채우고 '견적서 생성' 버튼을 누르면 DOCX/PDF 가 다운로드됩니다.")

    # 브랜드 기본 담당자 정보 (오버라이드 입력의 기본값)
    try:
        _brand_for_defaults = load_brand(PROJECT_ROOT, brand_id)
        _default_cp = _brand_for_defaults.contact_person
    except Exception:
        _default_cp = None

    st.subheader("0. 발행 담당자 정보")
    st.caption(
        "이번 견적서에 표시될 **소프트먼트 측 담당자** 정보입니다. "
        "브랜드 기본값에서 자동으로 채워졌습니다 — 자주 바뀌면 여기서 매번 수정하세요. "
        "(영구 변경은 ⚙ 설정 페이지에서)"
    )
    ic1, ic2 = st.columns(2)
    with ic1:
        issuer_name = st.text_input(
            "담당자명", key="issuer_name",
            value=(_default_cp.name if _default_cp else ""),
            placeholder="예: 박지은",
        )
        issuer_phone = st.text_input(
            "연락처", key="issuer_phone",
            value=(_default_cp.phone if _default_cp and _default_cp.phone else ""),
            placeholder="예: 010-0000-0000",
        )
    with ic2:
        issuer_title = st.text_input(
            "직책", key="issuer_title",
            value=(_default_cp.title if _default_cp and _default_cp.title else ""),
            placeholder="예: QR사업부 매니저",
        )
        issuer_email = st.text_input(
            "이메일", key="issuer_email",
            value=(_default_cp.email if _default_cp and _default_cp.email else ""),
            placeholder="예: name@softment.co.kr",
        )

    st.subheader("1. 수신처 정보")
    cp_col1, cp_col2 = st.columns(2)
    with cp_col1:
        cp_name = st.text_input("회사명 *", key="cp_name",
                                placeholder="예: 주식회사 ○○")
        cp_reg = st.text_input("사업자등록번호", key="cp_reg",
                               placeholder="000-00-00000")
        cp_address = st.text_input("주소", key="cp_address",
                                   placeholder="시/도 ○○구 ○○로 ...")
    with cp_col2:
        cp_contact_name = st.text_input("담당자", key="cp_contact_name",
                                        placeholder="예: 김담당")
        cp_contact_title = st.text_input("직책", key="cp_contact_title",
                                         placeholder="예: 구매팀장")
        cp_email = st.text_input("Email", key="cp_email",
                                 placeholder="buyer@example.com")

    st.subheader("2. 건명")
    subject = st.text_input("건명 *", key="subject",
                            placeholder="예: 주문 접수 / QR오더 솔루션 도입 견적",
                            label_visibility="collapsed")

    st.subheader("3. 품목 내역")
    st.caption(
        "💡 아래 드롭다운에서 카탈로그 상품을 선택해 추가하거나, "
        "'+ 빈 행' 으로 직접 입력하세요. 표 셀을 클릭해 수정 가능합니다."
    )

    pick_col, add_col, blank_col, reset_col = st.columns([6, 1.2, 1.2, 1.2])
    with pick_col:
        if products:
            options = list(range(len(products)))
            picked_idx = st.selectbox(
                "카탈로그에서 추가",
                options=options,
                index=None,
                format_func=lambda i: (
                    f"{products[i]['name']} · ₩{int(products[i].get('unit_price', 0)):,}"
                ),
                placeholder="상품 선택 (입력해서 검색 가능)...",
                key="catalog_pick",
                label_visibility="collapsed",
            )
        else:
            st.info("카탈로그가 비어있습니다. '카탈로그 관리' 페이지에서 상품을 추가하세요.")
            picked_idx = None
    with add_col:
        if st.button("+ 추가", use_container_width=True,
                     disabled=picked_idx is None,
                     help="선택한 카탈로그 상품을 품목 표에 추가합니다."):
            _add_catalog_row(products[picked_idx])
            st.rerun()
    with blank_col:
        if st.button("+ 빈 행", use_container_width=True):
            _add_blank_row()
            st.rerun()
    with reset_col:
        if st.button("전체 초기화", use_container_width=True):
            _reset_items()
            st.rerun()

    if st.session_state.items_df.empty:
        st.info("아직 품목이 없습니다. 위 드롭다운에서 카탈로그 상품을 추가하거나 '+ 빈 행' 을 누르세요.")
        edited_df = st.session_state.items_df
    else:
        edited_df = st.data_editor(
            st.session_state.items_df,
            column_config={
                "품목": st.column_config.TextColumn("품목", required=True),
                "설명": st.column_config.TextColumn("설명"),
                "수량": st.column_config.NumberColumn("수량", min_value=0, step=1),
                "기간(회)": st.column_config.NumberColumn("기간(회)", min_value=0, step=1),
                "단가": st.column_config.NumberColumn("단가", min_value=0, step=1000,
                                                      format="₩%d"),
                "비고": st.column_config.TextColumn("비고"),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="items_editor",
        )
        st.session_state.items_df = edited_df

    if not edited_df.empty:
        amounts = edited_df.apply(_row_amount, axis=1)
        subtotal = int(amounts.sum())
        vat_rate = labels.quote.vat_rate
        vat = int(round(subtotal * vat_rate))
        total = subtotal + vat

        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("공급가액", f"₩{subtotal:,}")
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", f"₩{vat:,}")
        m3.metric("합계 금액", f"₩{total:,}")

    st.subheader("4. 기타 안내 (선택)")
    notes = st.text_area(
        "기타 안내", key="notes",
        placeholder="이 견적서에만 들어갈 추가 문구가 있다면 적어주세요.",
        label_visibility="collapsed", height=80,
    )

    # 자동으로 추가되는 문구 미리보기 (config + 브랜드 입금계좌 기준)
    auto_lines = _compute_auto_notice_lines(brand_id, valid_until, labels)
    if auto_lines:
        with st.expander("ℹ 견적서에 자동으로 추가되는 문구 (펼쳐서 확인)", expanded=True):
            for line in auto_lines:
                st.markdown(f"- {line}")
            st.caption("위 문구는 위에 입력한 내용과 함께 '기타 안내' 영역에 자동 포함됩니다.")

    st.divider()
    common_args = dict(
        brand_id=brand_id, issued_date=issued_date, valid_until=valid_until,
        counterparty_data=dict(
            name=cp_name.strip(),
            registration_number=cp_reg.strip() or None,
            address=cp_address.strip() or None,
            contact_name=cp_contact_name.strip() or None,
            contact_title=cp_contact_title.strip() or None,
            email=cp_email.strip() or None,
        ),
        issuer_contact_data=dict(
            name=issuer_name.strip(),
            title=issuer_title.strip() or None,
            phone=issuer_phone.strip() or None,
            email=issuer_email.strip() or None,
        ),
        subject=subject.strip(),
        items_df=edited_df,
        notes=notes.strip() or None,
        soffice_available=soffice_available,
    )

    btn_prev, btn_gen = st.columns([1, 1])
    with btn_prev:
        preview_clicked = st.button(
            "👁 미리보기 (PDF)", use_container_width=True,
            disabled=not soffice_available,
            help=("작성한 내용을 PDF로 렌더링해서 이 페이지 안에서 바로 확인합니다. "
                  "다운로드는 아래 '견적서 생성' 버튼을 누르세요.")
            if soffice_available else "PDF 변환기(LibreOffice) 미감지 — 미리보기 불가",
        )
    with btn_gen:
        generate_clicked = st.button(
            "📝 견적서 생성 (다운로드)", type="primary", use_container_width=True,
        )

    if preview_clicked:
        _preview_quote(**common_args)
    if generate_clicked:
        _generate_quote(**common_args)


def _compute_auto_notice_lines(brand_id: str, valid_until: date,
                               labels: DocumentLabels) -> list[str]:
    """'기타 안내' 영역에 자동으로 들어가는 고정/계산 문구 목록을 반환."""
    ql = labels.quote
    lines: list[str] = []
    if valid_until and ql.auto_notices.validity_template:
        valid_str = valid_until.strftime("%Y년 %m월 %d일")
        lines.append(ql.auto_notices.validity_template.format(valid_until=valid_str))
    try:
        brand = load_brand(PROJECT_ROOT, brand_id)
    except Exception:
        brand = None
    if brand and brand.bank_account and ql.auto_notices.bank_account_template:
        ba = brand.bank_account
        lines.append(ql.auto_notices.bank_account_template.format(
            bank=ba.bank,
            account_number=ba.account_number,
            account_holder=ba.account_holder or "",
        ))
    for static in ql.static_notices:
        if static:
            lines.append(static)
    return lines


def _build_quote_artifacts(*, brand_id, issued_date, valid_until,
                           counterparty_data, issuer_contact_data,
                           subject, items_df, notes,
                           soffice_available, status_label="문서 생성 중..."):
    """입력값을 검증·렌더링하여 (document_id, docx_bytes, pdf_bytes) 를 반환.
    실패 시 None 반환 (사용자에게 에러는 이미 표시됨)."""
    if not counterparty_data["name"]:
        st.error("❌ 회사명은 필수입니다.")
        return None
    if not subject:
        st.error("❌ 건명은 필수입니다.")
        return None

    items: list[LineItem] = []
    for _, row in items_df.iterrows():
        name = row.get("품목")
        name = name.strip() if isinstance(name, str) else None
        if not name:
            continue
        qty = row.get("수량")
        period = row.get("기간(회)")
        unit_price = row.get("단가") or 0
        desc_val = row.get("설명")
        notes_val = row.get("비고")
        items.append(LineItem(
            name=name,
            description=desc_val if isinstance(desc_val, str) and desc_val else None,
            qty=float(qty) if pd.notna(qty) and qty else None,
            period=float(period) if pd.notna(period) and period else None,
            unit_price=float(unit_price),
            notes=notes_val if isinstance(notes_val, str) and notes_val else None,
        ))

    if not items:
        st.error("❌ 최소 1개 이상의 품목을 입력해주세요.")
        return None

    cp_name_short = "".join(c for c in counterparty_data["name"] if c.isalnum())[:8]
    document_id = f"Q-{issued_date.strftime('%Y%m%d')}-{cp_name_short or '고객'}"

    document = QuoteDocument(
        document_id=document_id, document_type="quote",
        brand_id=brand_id,
        issued_date=issued_date, valid_until=valid_until,
        counterparty=Counterparty(**counterparty_data),
        subject=subject, line_items=items,
        totals=Totals(), clauses=[], notes=notes,
    )

    try:
        brand = load_brand(PROJECT_ROOT, brand_id)
    except FileNotFoundError as e:
        st.error(f"❌ 브랜드 로드 실패: {e}")
        return None

    # 발행 담당자 정보 — 폼에서 입력한 값으로 일회용 오버라이드
    issuer_name = (issuer_contact_data.get("name") or "").strip()
    if issuer_name:
        brand = brand.model_copy(update={
            "contact_person": ContactPerson(
                name=issuer_name,
                title=issuer_contact_data.get("title"),
                phone=issuer_contact_data.get("phone"),
                email=issuer_contact_data.get("email"),
            )
        })
    else:
        # 담당자명 비워두면 견적서에서 담당자/연락처/이메일 줄 자체를 숨김
        brand = brand.model_copy(update={"contact_person": None})

    output_dir = PROJECT_ROOT / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_path = output_dir / f"{document_id}.docx"

    with st.status(status_label, expanded=True) as status:
        st.write("📝 DOCX 렌더링 중...")
        try:
            render_docx(brand, document, PROJECT_ROOT, docx_path)
        except Exception as e:
            status.update(label="❌ DOCX 생성 실패", state="error")
            st.exception(e)
            return None
        st.write(f"   ✓ {docx_path.name}")

        pdf_bytes = None
        if soffice_available:
            st.write("📑 PDF 변환 중...")
            try:
                pdf_path = convert_docx_to_pdf(docx_path, output_dir)
                pdf_bytes = pdf_path.read_bytes()
                st.write(f"   ✓ {pdf_path.name}")
            except Exception as e:
                st.warning(f"PDF 변환 실패 (DOCX만 사용 가능): {e}")

        status.update(label="✅ 완료", state="complete")

    docx_bytes = docx_path.read_bytes()
    return document_id, docx_bytes, pdf_bytes


def _generate_quote(**kwargs):
    result = _build_quote_artifacts(status_label="문서 생성 중...", **kwargs)
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📝 DOCX 다운로드", data=docx_bytes,
            file_name=f"{document_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl2:
        if pdf_bytes:
            st.download_button(
                "📑 PDF 다운로드", data=pdf_bytes,
                file_name=f"{document_id}.pdf",
                mime="application/pdf", use_container_width=True,
            )
        else:
            st.button("📑 PDF 사용 불가", disabled=True, use_container_width=True)


def _preview_quote(**kwargs):
    result = _build_quote_artifacts(status_label="미리보기 생성 중...", **kwargs)
    if result is None:
        return
    document_id, _docx_bytes, pdf_bytes = result

    if not pdf_bytes:
        st.error("❌ PDF 변환기(LibreOffice)를 사용할 수 없어 미리보기를 표시할 수 없습니다.")
        return

    st.success(f"미리보기: **{document_id}.pdf**")

    # PDF를 페이지별 이미지로 렌더링 (브라우저 PDF 차단 회피)
    rendered = False
    try:
        import pypdfium2 as pdfium  # noqa: WPS433  (지연 임포트로 실패 시 graceful fallback)
        pdf = pdfium.PdfDocument(pdf_bytes)
        try:
            n_pages = len(pdf)
            for i in range(n_pages):
                page = pdf[i]
                bitmap = page.render(scale=2)  # 2배 해상도 (선명)
                img = bitmap.to_pil()
                st.image(img, caption=f"페이지 {i + 1} / {n_pages}",
                         use_container_width=True)
            rendered = True
        finally:
            pdf.close()
    except Exception as e:
        st.warning(f"이미지 미리보기를 사용할 수 없습니다 ({e}). 아래 다운로드로 확인하세요.")

    # 항상 미리보기 PDF 다운로드 제공 (이미지 렌더 실패 시 fallback)
    st.download_button(
        ("📑 미리보기 PDF 다운로드" if rendered else "📑 미리보기 PDF 다운로드 (필수)"),
        data=pdf_bytes,
        file_name=f"preview_{document_id}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    st.caption("💡 미리보기는 화면 표시용입니다. 정식 파일은 '📝 견적서 생성' 버튼을 누르세요.")


# ═════════════════════════════════════════════════════════════
# 페이지 2: 카탈로그 관리
# ═════════════════════════════════════════════════════════════

def render_catalog_page():
    st.title("📦 카탈로그 관리")
    st.caption(
        "견적서 작성 화면의 '카탈로그 빠른 추가' 와 엑셀 템플릿의 드롭다운에 사용되는 상품 목록입니다. "
        "행을 추가/수정/삭제 후 **저장** 버튼을 누르세요."
    )

    products = _load_products()
    if not products:
        # 빈 카탈로그면 예시 행 한 줄 미리 채워두기
        df = pd.DataFrame([{
            "code": "NEW-CODE",
            "name": "(여기를 클릭해서 수정)",
            "description": "",
            "unit_price": 0,
            "currency": "KRW",
        }])
    else:
        df = pd.DataFrame(products)
        # 누락 가능한 컬럼 보완
        for col, default in [("description", ""), ("currency", "KRW")]:
            if col not in df.columns:
                df[col] = default

    edited = st.data_editor(
        df[["code", "name", "description", "unit_price", "currency"]],
        column_config={
            "code": st.column_config.TextColumn(
                "코드", help="고유 식별자 (영문/숫자/하이픈)", required=True,
            ),
            "name": st.column_config.TextColumn(
                "품목명", help="견적서에 표시되는 이름", required=True,
            ),
            "description": st.column_config.TextColumn(
                "설명/청구기준", help="예: '매장 / 월', '매장 월 정액제'",
            ),
            "unit_price": st.column_config.NumberColumn(
                "단가 (원)", min_value=0, step=1000, format="₩%d",
            ),
            "currency": st.column_config.SelectboxColumn(
                "통화", options=["KRW", "USD", "EUR", "JPY"],
            ),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="catalog_editor",
    )

    st.divider()
    col_save, col_info = st.columns([1, 3])
    with col_save:
        if st.button("💾 카탈로그 저장", type="primary", use_container_width=True):
            new_products = []
            for _, row in edited.iterrows():
                code = (row.get("code") or "").strip()
                name = (row.get("name") or "").strip()
                if not code or not name:
                    continue
                new_products.append({
                    "code": code,
                    "name": name,
                    "description": (row.get("description") or "").strip(),
                    "unit_price": int(row.get("unit_price") or 0),
                    "currency": row.get("currency") or "KRW",
                })
            _save_products(new_products)
            st.success(f"✅ {len(new_products)}개 상품 저장 완료. "
                       "'견적서 작성' 페이지로 가면 즉시 반영됩니다.")
    with col_info:
        st.caption(
            "📝 **저장 후**: 견적서 작성 페이지의 사이드바 카탈로그가 즉시 갱신됩니다. "
            "엑셀 템플릿의 드롭다운은 `template` 명령으로 재생성해야 반영됩니다."
        )


# ═════════════════════════════════════════════════════════════
# 페이지 3: 설정 (브랜드 정보 + 양식 라벨)
# ═════════════════════════════════════════════════════════════

def render_settings_page():
    st.title("⚙ 설정")
    tab_brand, tab_labels = st.tabs(["🏢 브랜드 정보", "📐 양식 라벨/문구"])

    with tab_brand:
        _render_brand_settings()
    with tab_labels:
        _render_label_settings()


def _render_brand_settings():
    brand_ids = _list_brands()
    if not brand_ids:
        st.warning("등록된 브랜드가 없습니다.")
        return

    target_brand_id = st.selectbox(
        "편집할 브랜드", brand_ids,
        disabled=len(brand_ids) <= 1,
    )
    try:
        brand = load_brand(PROJECT_ROOT, target_brand_id)
    except FileNotFoundError as e:
        st.error(f"브랜드 파일을 찾을 수 없습니다: {e}")
        return

    st.subheader("회사 기본 정보")
    c1, c2 = st.columns(2)
    with c1:
        name_ko = st.text_input("회사명 (한글)", brand.company.name_ko)
        registration_number = st.text_input("사업자등록번호", brand.company.registration_number)
        ceo = st.text_input("대표자", brand.company.ceo)
    with c2:
        name_en = st.text_input("회사명 (영문, 선택)", brand.company.name_en or "")
        phone = st.text_input("대표 전화 (선택)", brand.company.phone or "")
        email = st.text_input("대표 이메일 (선택)", brand.company.email or "")
    address = st.text_input("주소", brand.company.address or "")

    st.subheader("담당자 정보")
    cp_c1, cp_c2 = st.columns(2)
    cp = brand.contact_person
    with cp_c1:
        cp_name = st.text_input("담당자 이름", (cp.name if cp else ""))
        cp_title = st.text_input("직책", (cp.title if cp else "") or "")
    with cp_c2:
        cp_phone = st.text_input("담당자 전화", (cp.phone if cp else "") or "")
        cp_email = st.text_input("담당자 이메일", (cp.email if cp else "") or "")

    st.subheader("서명자 (견적서/계약서 서명란)")
    sig_c1, sig_c2 = st.columns(2)
    with sig_c1:
        signer_name = st.text_input("서명자명", brand.signature.signer_name)
    with sig_c2:
        signer_title = st.text_input("서명자 직책", brand.signature.signer_title)

    st.subheader("입금 계좌 (선택)")
    ba = brand.bank_account
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        bank = st.text_input("은행", (ba.bank if ba else ""))
    with bc2:
        account_number = st.text_input("계좌번호", (ba.account_number if ba else ""))
    with bc3:
        account_holder = st.text_input("예금주", (ba.account_holder if ba else "") or "")

    st.subheader("색상 (CI)")
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        primary = st.color_picker("Primary (주 색상)", brand.branding.colors.primary)
    with cc2:
        accent = st.color_picker("Accent (강조)", brand.branding.colors.accent)
    with cc3:
        text_color = st.color_picker("Text", brand.branding.colors.text)
    font_family = st.text_input("폰트 패밀리", brand.branding.font_family)

    st.divider()
    if st.button("💾 브랜드 정보 저장", type="primary", use_container_width=True):
        new_dict = {
            "brand_id": brand.brand_id,
            "company": {
                "name_ko": name_ko,
                "name_en": name_en or None,
                "registration_number": registration_number,
                "ceo": ceo,
                "address": address or None,
                "phone": phone or None,
                "email": email or None,
            },
            "branding": {
                "logo_path": brand.branding.logo_path,
                "colors": {
                    "primary": primary,
                    "accent": accent,
                    "text": text_color,
                },
                "font_family": font_family,
            },
            "signature": {
                "signer_name": signer_name,
                "signer_title": signer_title,
            },
            "contact_person": ({
                "name": cp_name,
                "title": cp_title or None,
                "phone": cp_phone or None,
                "email": cp_email or None,
            } if cp_name else None),
            "bank_account": ({
                "bank": bank,
                "account_number": account_number,
                "account_holder": account_holder or None,
            } if bank and account_number else None),
            "footer_text": brand.footer_text or "",
        }
        _save_brand(brand.brand_id, new_dict)
        st.success("✅ 브랜드 정보 저장 완료")


def _render_label_settings():
    labels = load_labels(PROJECT_ROOT)
    ql = labels.quote

    st.subheader("기본 값")
    c1, c2 = st.columns(2)
    with c1:
        vat_rate_pct = st.number_input(
            "VAT율 (%)", value=float(ql.vat_rate * 100),
            min_value=0.0, max_value=100.0, step=0.5,
            help="예: 10 = 10%",
        )
    with c2:
        default_validity_days = st.number_input(
            "유효기간 기본 일수", value=ql.default_validity_days,
            min_value=1, max_value=365,
        )

    st.subheader("자동 안내 문구")
    st.caption("`{valid_until}`, `{bank}`, `{account_number}`, `{account_holder}` 같은 자리표시자가 자동으로 채워집니다. "
               "**빈 칸으로 두면** 해당 자동 안내가 표시되지 않습니다.")
    validity_template = st.text_input(
        "유효기간 자동 문구",
        value=ql.auto_notices.validity_template,
        placeholder="본 견적의 유효기간은 {valid_until}까지입니다.",
    )
    bank_template = st.text_input(
        "입금 계좌 자동 문구",
        value=ql.auto_notices.bank_account_template,
        placeholder="입금 계좌: {bank} {account_number} (예금주: {account_holder})",
    )

    st.subheader("고정 안내 문구")
    st.caption("모든 견적서의 '기타 안내' 섹션에 항상 표시되는 문구. 한 줄에 하나씩 입력.")
    static_notices_text = st.text_area(
        "고정 안내",
        value="\n".join(ql.static_notices),
        height=100,
        label_visibility="collapsed",
    )
    static_notices = [
        line.strip() for line in static_notices_text.splitlines() if line.strip()
    ]

    st.subheader("표시 라벨")
    l1, l2, l3 = st.columns(3)
    with l1:
        subtotal_label = st.text_input("공급가액 라벨", ql.labels.subtotal)
        counterparty_section = st.text_input("수신처 섹션", ql.labels.counterparty_section)
    with l2:
        vat_label = st.text_input("부가세 라벨", ql.labels.vat)
        subject_prefix = st.text_input("건명 prefix", ql.labels.subject_prefix)
    with l3:
        total_label = st.text_input("합계 라벨", ql.labels.total)
        etc_notice_section = st.text_input("기타 안내 섹션", ql.labels.etc_notice_section)

    vat_separate = st.text_input(
        "VAT 별도 표시 (비우면 미표시)",
        value=ql.labels.vat_separate_notice,
        placeholder="* VAT 별도",
    )

    st.subheader("표 컬럼 헤더")
    th = ql.table_headers
    th1, th2, th3, th4 = st.columns(4)
    with th1:
        h_name = st.text_input("품목 컬럼명", th.name)
        h_unit_price = st.text_input("단가 컬럼명", th.unit_price)
    with th2:
        h_description = st.text_input("설명 컬럼명", th.description)
        h_amount = st.text_input("금액 컬럼명", th.amount)
    with th3:
        h_qty = st.text_input("수량 컬럼명", th.qty)
        h_notes = st.text_input("비고 컬럼명", th.notes)
    with th4:
        h_period = st.text_input("기간 컬럼명", th.period)

    st.subheader("문서 제목")
    t1, t2 = st.columns(2)
    with t1:
        quote_title = st.text_input("견적서 제목", ql.title)
    with t2:
        contract_title = st.text_input("계약서 제목", labels.contract.title)

    st.divider()
    if st.button("💾 양식 설정 저장", type="primary", use_container_width=True):
        new_labels = {
            "_doc": "양식의 라벨·문구·기본값을 관리하는 파일입니다.",
            "quote": {
                "title": quote_title,
                "vat_rate": round(vat_rate_pct / 100, 4),
                "default_validity_days": int(default_validity_days),
                "labels": {
                    "counterparty_section": counterparty_section,
                    "subject_prefix": subject_prefix,
                    "etc_notice_section": etc_notice_section,
                    "vat_separate_notice": vat_separate,
                    "subtotal": subtotal_label,
                    "vat": vat_label,
                    "total": total_label,
                },
                "table_headers": {
                    "name": h_name,
                    "description": h_description,
                    "qty": h_qty,
                    "period": h_period,
                    "unit_price": h_unit_price,
                    "amount": h_amount,
                    "notes": h_notes,
                },
                "auto_notices": {
                    "validity_template": validity_template,
                    "bank_account_template": bank_template,
                },
                "static_notices": static_notices,
            },
            "contract": {
                "title": contract_title,
                "preamble_template": labels.contract.preamble_template,
                "labels": labels.contract.labels.model_dump(),
            },
        }
        _save_labels(new_labels)
        st.success("✅ 양식 설정 저장 완료. 새로 생성하는 견적서/계약서에 반영됩니다.")


# ═════════════════════════════════════════════════════════════
# 메인 라우터
# ═════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="견적서 자동 생성", page_icon="📋", layout="wide")

    with st.sidebar:
        st.markdown("### 메뉴")
        page = st.radio(
            "페이지",
            ["📋 견적서 작성", "📦 카탈로그 관리", "⚙ 설정"],
            label_visibility="collapsed",
        )

    if page == "📋 견적서 작성":
        render_quote_page()
    elif page == "📦 카탈로그 관리":
        render_catalog_page()
    elif page == "⚙ 설정":
        render_settings_page()


if __name__ == "__main__":
    main()
