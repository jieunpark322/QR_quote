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
from src.membership_models import (
    MembershipCategory,
    MembershipLineItem,
    MembershipParty,
    MembershipQuoteDocument,
    MembershipScenario,
    MembershipSection,
    MembershipSubItem,
    category_subtotal,
    scenario_grand_total_by_period,
    scenario_vat_and_total,
    section_subtotals_by_period,
)
from src.membership_renderer import render_membership_docx
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
# 자동저장 (새로고침/이탈 보호)
# ═════════════════════════════════════════════════════════════

AUTOSAVE_DIR = PROJECT_ROOT / "output" / "_autosave"
QR_AUTOSAVE_PATH = AUTOSAVE_DIR / "qr_quote.json"
MC_AUTOSAVE_PATH = AUTOSAVE_DIR / "membership_quote.json"

QR_HISTORY_DIR = PROJECT_ROOT / "output" / "_history" / "qr"
QR_TEMPLATE_DIR = PROJECT_ROOT / "output" / "_templates" / "qr"
HISTORY_LIMIT = 10

# 자동 저장/복원할 위젯 키들
QR_FORM_KEYS = [
    "issuer_name", "issuer_phone", "issuer_title", "issuer_email",
    "cp_name", "cp_reg", "cp_address", "cp_contact_name",
    "cp_contact_title", "cp_email", "subject", "notes",
    "total_discount_pct",
]


def _read_json_safe(path: Path) -> dict | None:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _write_json_safe(path: Path, payload: dict) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, default=str),
                        encoding="utf-8")
    except OSError:
        pass


def _qr_autosave_load_once() -> None:
    """페이지 진입 시 1회만 호출 — 디스크의 autosave를 session_state에 채워넣음."""
    if st.session_state.get("_qr_autosave_loaded"):
        return
    st.session_state["_qr_autosave_loaded"] = True
    payload = _read_json_safe(QR_AUTOSAVE_PATH)
    if not payload:
        return
    items = payload.get("items")
    if items:
        try:
            st.session_state["items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in QR_FORM_KEYS:
        v = payload.get(k)
        if v not in (None, "") and k not in st.session_state:
            st.session_state[k] = v


def _qr_autosave_write() -> None:
    df = st.session_state.get("items_df")
    payload = {k: st.session_state.get(k) for k in QR_FORM_KEYS}
    if df is not None and not df.empty:
        payload["items"] = df.to_dict(orient="records")
    has_any = (df is not None and not df.empty) or any(payload.get(k) for k in QR_FORM_KEYS)
    if has_any:
        _write_json_safe(QR_AUTOSAVE_PATH, payload)
    else:
        # 비어있으면 autosave 파일 제거
        try:
            QR_AUTOSAVE_PATH.unlink(missing_ok=True)
        except OSError:
            pass


def _qr_autosave_clear() -> None:
    try:
        QR_AUTOSAVE_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def _mc_autosave_load_once() -> None:
    if st.session_state.get("_mc_autosave_loaded"):
        return
    st.session_state["_mc_autosave_loaded"] = True
    payload = _read_json_safe(MC_AUTOSAVE_PATH)
    if not payload:
        return
    if "mc_doc" not in st.session_state and payload.get("mc_doc"):
        st.session_state["mc_doc"] = payload["mc_doc"]
    for k in ("mc_issuer_name", "mc_issuer_title",
              "mc_issuer_phone", "mc_issuer_email"):
        v = payload.get(k)
        if v not in (None, "") and k not in st.session_state:
            st.session_state[k] = v


def _mc_autosave_write() -> None:
    doc = st.session_state.get("mc_doc")
    payload = {
        "mc_doc": doc,
        "mc_issuer_name": st.session_state.get("mc_issuer_name"),
        "mc_issuer_title": st.session_state.get("mc_issuer_title"),
        "mc_issuer_phone": st.session_state.get("mc_issuer_phone"),
        "mc_issuer_email": st.session_state.get("mc_issuer_email"),
    }
    has_any = bool(doc) or any(payload.get(k) for k in payload if k != "mc_doc")
    if has_any:
        _write_json_safe(MC_AUTOSAVE_PATH, payload)
    else:
        try:
            MC_AUTOSAVE_PATH.unlink(missing_ok=True)
        except OSError:
            pass


def _mc_autosave_clear() -> None:
    try:
        MC_AUTOSAVE_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def _qr_snapshot_payload() -> dict:
    """현재 견적서 입력 상태 전체를 직렬화 가능한 dict로."""
    df = st.session_state.get("items_df")
    payload = {k: st.session_state.get(k) for k in QR_FORM_KEYS}
    if df is not None and not df.empty:
        payload["items"] = df.to_dict(orient="records")
    return payload


def _qr_apply_snapshot(payload: dict) -> None:
    """저장된 견적서 payload를 session_state에 복원."""
    items = payload.get("items")
    if items:
        try:
            st.session_state["items_df"] = pd.DataFrame(items)
        except (ValueError, KeyError):
            pass
    for k in QR_FORM_KEYS:
        v = payload.get(k)
        if v not in (None, ""):
            st.session_state[k] = v


def _qr_save_history(snapshot: dict, document_id: str) -> None:
    """견적서 생성 직후 히스토리에 저장. 오래된 것은 정리."""
    QR_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_id = "".join(c for c in document_id if c.isalnum() or c in "-_")[:40]
    payload = {
        **snapshot,
        "_document_id": document_id,
        "_saved_at": datetime.now().isoformat(timespec="seconds"),
        "_subject": st.session_state.get("subject", ""),
        "_cp_name": st.session_state.get("cp_name", ""),
    }
    path = QR_HISTORY_DIR / f"{ts}_{safe_id}.json"
    _write_json_safe(path, payload)
    # 오래된 히스토리 정리
    files = sorted(QR_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    for old in files[HISTORY_LIMIT:]:
        try:
            old.unlink()
        except OSError:
            pass


def _qr_list_history() -> list[tuple[Path, dict]]:
    if not QR_HISTORY_DIR.exists():
        return []
    files = sorted(QR_HISTORY_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)[:HISTORY_LIMIT]
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _qr_save_template(name: str, snapshot: dict) -> bool:
    name = (name or "").strip()
    if not name:
        return False
    QR_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in name if c.isalnum() or c in " -_가-힣")[:60].strip()
    if not safe:
        return False
    from datetime import datetime
    payload = {
        **snapshot,
        "_template_name": name,
        "_saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    path = QR_TEMPLATE_DIR / f"{safe}.json"
    _write_json_safe(path, payload)
    return True


def _qr_list_templates() -> list[tuple[Path, dict]]:
    if not QR_TEMPLATE_DIR.exists():
        return []
    files = sorted(QR_TEMPLATE_DIR.glob("*.json"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    result = []
    for p in files:
        data = _read_json_safe(p)
        if data:
            result.append((p, data))
    return result


def _qr_delete_template(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _inject_beforeunload(active: bool) -> None:
    """작성 중 내용이 있으면 페이지 이탈 시 브라우저 경고를 띄움."""
    from streamlit.components.v1 import html
    flag = "1" if active else "0"
    html(f"""
<script>
(function() {{
  try {{
    var w = window.parent;
    if (!w._qrUnloadHandler) {{
      w._qrUnloadHandler = function(e) {{
        if (w._qrUnloadActive) {{
          e.preventDefault();
          e.returnValue = '작성 중인 내용이 있어요. 정말 이 페이지를 떠나시겠어요?';
          return e.returnValue;
        }}
      }};
      w.addEventListener('beforeunload', w._qrUnloadHandler);
    }}
    w._qrUnloadActive = ({flag} === 1);
  }} catch (err) {{ /* iframe cross-origin 등 무시 */ }}
}})();
</script>
""", height=0)


def _render_qr_history_template_panel() -> None:
    """견적서 작성 페이지 상단의 '최근/표본 불러오기' UI."""
    history = _qr_list_history()
    templates = _qr_list_templates()
    if not history and not templates:
        with st.expander("📂 최근 견적서 · 📋 표본", expanded=False):
            st.caption(
                "💡 견적서를 한 번 생성하면 여기에 **최근 10개** 가 자동 보관되어 "
                "다시 불러올 수 있고, 자주 쓰는 형태는 **표본** 으로 저장해 두면 "
                "다음 견적서 작성 시 끌어와서 쓸 수 있어요."
            )
            tpl_name = st.text_input(
                "표본 이름", placeholder="예: 단가 표준안",
                key="qr_tpl_name_input",
                label_visibility="collapsed",
            )
            if st.button("💾 현재 입력을 표본으로 저장",
                         use_container_width=True, key="qr_tpl_save_empty"):
                if _qr_save_template(tpl_name, _qr_snapshot_payload()):
                    st.success(f"✅ 표본 저장: {tpl_name}")
                    st.rerun()
                else:
                    st.warning("표본 이름을 입력해 주세요.")
        return

    with st.expander(
        f"📂 최근 견적서 ({len(history)}) · 📋 표본 ({len(templates)})",
        expanded=False,
    ):
        tab_hist, tab_tpl = st.tabs(["📂 최근 견적서", "📋 표본 (수기 저장)"])

        with tab_hist:
            if not history:
                st.caption("아직 생성한 견적서가 없습니다. 한 번 생성하면 여기에 자동 저장됩니다.")
            else:
                st.caption("견적서를 생성하면 자동으로 최근 10개까지 보관됩니다. 클릭해서 이어서 편집할 수 있어요.")
                options = list(range(len(history)))

                def _fmt_hist(i: int) -> str:
                    _, data = history[i]
                    saved = (data.get("_saved_at") or "")[:16].replace("T", " ")
                    subj = data.get("_subject") or "(건명 없음)"
                    cp = data.get("_cp_name") or "(수신처 없음)"
                    return f"🕐 {saved}  ·  {cp}  ·  {subj}"

                sel_h = st.selectbox(
                    "불러올 견적서", options=options,
                    format_func=_fmt_hist, index=None,
                    placeholder="선택하세요...", key="qr_hist_select",
                )
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("📥 선택한 견적서 불러오기",
                                 use_container_width=True,
                                 disabled=sel_h is None,
                                 key="qr_hist_load"):
                        _, data = history[sel_h]
                        _qr_apply_snapshot(data)
                        st.success("✅ 견적서를 불러왔습니다.")
                        st.rerun()
                with col_b:
                    if st.button("🗑 선택한 견적서 히스토리에서 삭제",
                                 use_container_width=True,
                                 disabled=sel_h is None,
                                 key="qr_hist_del"):
                        path, _ = history[sel_h]
                        try:
                            path.unlink(missing_ok=True)
                        except OSError:
                            pass
                        st.rerun()

        with tab_tpl:
            st.caption(
                "**표본** 은 자주 쓰는 견적서 형태를 수기로 저장해 두는 공간입니다. "
                "현재 입력값을 그대로 표본으로 저장하거나, 저장된 표본을 끌어와서 시작할 수 있어요."
            )
            save_col1, save_col2 = st.columns([3, 2])
            with save_col1:
                tpl_name = st.text_input(
                    "표본 이름", placeholder="예: 단가 표준안 / 가맹점 A 견적",
                    key="qr_tpl_name", label_visibility="collapsed",
                )
            with save_col2:
                if st.button("💾 현재 입력을 표본으로 저장",
                             use_container_width=True, key="qr_tpl_save"):
                    if _qr_save_template(tpl_name, _qr_snapshot_payload()):
                        st.success(f"✅ 표본 저장: {tpl_name}")
                        st.rerun()
                    else:
                        st.warning("표본 이름을 입력해 주세요.")
            st.divider()
            if templates:
                options = list(range(len(templates)))

                def _fmt_tpl(i: int) -> str:
                    _, data = templates[i]
                    name = data.get("_template_name") or templates[i][0].stem
                    saved = (data.get("_saved_at") or "")[:10]
                    return f"📋 {name}  ·  {saved}"

                sel_t = st.selectbox(
                    "저장된 표본", options=options,
                    format_func=_fmt_tpl, index=None,
                    placeholder="선택하세요...", key="qr_tpl_select",
                )
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("📥 선택한 표본으로 시작",
                                 use_container_width=True,
                                 disabled=sel_t is None,
                                 key="qr_tpl_load"):
                        _, data = templates[sel_t]
                        _qr_apply_snapshot(data)
                        st.success("✅ 표본을 불러왔습니다.")
                        st.rerun()
                with col_b:
                    if st.button("🗑 선택한 표본 삭제",
                                 use_container_width=True,
                                 disabled=sel_t is None,
                                 key="qr_tpl_del"):
                        path, _ = templates[sel_t]
                        _qr_delete_template(path)
                        st.rerun()
            else:
                st.caption("저장된 표본이 없습니다.")


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

# 편집 가능 컬럼 (공급가는 계산 결과라 별도)
ITEM_COLUMNS = ["분류", "항목", "설명", "단가", "기간(횟수)", "수량",
                "할인율(%)", "할인금액", "비고"]
# 화면 표시 순서 (공급가 포함)
DISPLAY_COLUMNS = ["분류", "항목", "설명", "단가", "기간(횟수)", "수량",
                   "할인율(%)", "할인금액", "공급가", "비고"]

# 분류 선택지
ITEM_KIND_NORMAL = "📋 품목"
ITEM_KIND_DISCOUNT = "💰 할인"
ITEM_KINDS = [ITEM_KIND_NORMAL, ITEM_KIND_DISCOUNT]


def _empty_items_df() -> pd.DataFrame:
    return pd.DataFrame(columns=ITEM_COLUMNS).astype({
        "분류": "string",
        "항목": "string",
        "설명": "string",
        "단가": "Int64",
        "기간(횟수)": "Int64",
        "수량": "Int64",
        "할인율(%)": "Int64",
        "할인금액": "Int64",
        "비고": "string",
    })


def _ensure_items_state():
    if "items_df" not in st.session_state:
        st.session_state.items_df = _empty_items_df()


def _add_catalog_row(product: dict):
    new_row = {
        "분류": ITEM_KIND_NORMAL,
        "항목": product["name"],
        "설명": product.get("description", ""),
        "단가": int(product.get("unit_price", 0)),
        "기간(횟수)": 1,
        "수량": 1,
        "할인율(%)": None,
        "할인금액": None,
        "비고": "",
    }
    df = st.session_state.items_df
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True
    )


def _add_blank_row():
    df = st.session_state.items_df
    blank = {c: None for c in ITEM_COLUMNS}
    blank["분류"] = ITEM_KIND_NORMAL
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([blank])],
        ignore_index=True,
    )


def _add_discount_row():
    """할인 행 추가 — 분류 '💰 할인'. 단가는 양수 입력 시 자동 차감."""
    df = st.session_state.items_df
    new_row = {
        "분류": ITEM_KIND_DISCOUNT,
        "항목": "할인",
        "설명": "협상가",
        "단가": 500000,
        "기간(횟수)": None,
        "수량": 1,
        "할인율(%)": None,
        "할인금액": None,
        "비고": "협상에 따라 단가 조정",
    }
    st.session_state.items_df = pd.concat(
        [df, pd.DataFrame([new_row])], ignore_index=True,
    )


def _reset_items():
    st.session_state.items_df = _empty_items_df()
    for k in QR_FORM_KEYS:
        if k in st.session_state:
            del st.session_state[k]
    _qr_autosave_clear()


def _row_amount(row):
    """행의 공급가 계산 (할인 반영). 입력이 전혀 없으면 None."""
    qty = row.get("수량")
    period = row.get("기간(횟수)")
    price = row.get("단가")
    if not (pd.notna(price) and price):
        return None
    q = qty if pd.notna(qty) and qty else 1
    p = period if pd.notna(period) and period else 1
    try:
        gross = int(q) * int(p) * int(price)
    except (TypeError, ValueError):
        return None
    # 분류='💰 할인' 이면 무조건 차감 (양수 입력도 음수 처리)
    if row.get("분류") == ITEM_KIND_DISCOUNT:
        return -abs(gross)
    # 일반 항목별 할인: 할인금액 > 할인율
    disc_amt = row.get("할인금액")
    disc_rate = row.get("할인율(%)")
    discount = 0
    if pd.notna(disc_amt) and disc_amt:
        try:
            discount = int(disc_amt)
        except (TypeError, ValueError):
            discount = 0
    elif pd.notna(disc_rate) and disc_rate:
        try:
            discount = int(round(gross * float(disc_rate) / 100))
        except (TypeError, ValueError):
            discount = 0
    return gross - discount


def render_quote_page():
    _qr_autosave_load_once()
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

    # ─── 최근 견적서 / 표본 불러오기 ───
    _render_qr_history_template_panel()

    st.subheader("0. 발행 담당자 정보")
    st.caption(
        "이번 견적서에 표시될 **소프트먼트 측 담당자** 정보입니다. "
        "비워두면 견적서에서 담당자 정보가 표시되지 않습니다."
    )
    # Tab 키가 좌→우→다음 줄 순으로 이동하도록 행 단위로 columns 를 생성
    ic_r1_l, ic_r1_r = st.columns(2)
    with ic_r1_l:
        issuer_name = st.text_input(
            "담당자명", key="issuer_name",
            placeholder="예: 박지은",
        )
    with ic_r1_r:
        issuer_title = st.text_input(
            "직책", key="issuer_title",
            placeholder="예: QR사업부 매니저",
        )
    ic_r2_l, ic_r2_r = st.columns(2)
    with ic_r2_l:
        issuer_phone = st.text_input(
            "연락처", key="issuer_phone",
            placeholder="예: 010-0000-0000",
        )
    with ic_r2_r:
        issuer_email = st.text_input(
            "이메일", key="issuer_email",
            placeholder="예: name@softment.co.kr",
        )

    st.subheader("1. 수신처 정보")
    # Tab 이 좌→우→다음 줄 순으로 이동하도록 행 단위 columns
    cp_r1_l, cp_r1_r = st.columns(2)
    with cp_r1_l:
        cp_name = st.text_input("회사명 *", key="cp_name",
                                placeholder="예: 주식회사 ○○")
    with cp_r1_r:
        cp_contact_name = st.text_input("담당자", key="cp_contact_name",
                                        placeholder="예: 김담당")
    cp_r2_l, cp_r2_r = st.columns(2)
    with cp_r2_l:
        cp_reg = st.text_input("사업자등록번호", key="cp_reg",
                               placeholder="000-00-00000")
    with cp_r2_r:
        cp_contact_title = st.text_input("직책", key="cp_contact_title",
                                         placeholder="예: 구매팀장")
    cp_r3_l, cp_r3_r = st.columns(2)
    with cp_r3_l:
        cp_address = st.text_input("주소", key="cp_address",
                                   placeholder="시/도 ○○구 ○○로 ...")
    with cp_r3_r:
        cp_email = st.text_input("Email", key="cp_email",
                                 placeholder="buyer@example.com")

    st.subheader("2. 건명")
    subject = st.text_input("건명 *", key="subject",
                            placeholder="예: 주문 접수 / QR오더 솔루션 도입 견적",
                            label_visibility="collapsed")

    st.subheader("3. 품목 내역")
    st.caption(
        "💡 카탈로그 선택 또는 '+ 빈 행' 으로 추가. 할인은 **'+ 할인 행'** 으로 "
        "음수 단가 행이 자동 추가됩니다. 표 셀을 클릭해 수정 가능합니다."
    )

    pick_col, add_col, blank_col, disc_col, reset_col = st.columns([5, 1.2, 1.2, 1.2, 1.2])
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
    with disc_col:
        if st.button("💰 + 할인 행", use_container_width=True,
                     help="음수 단가 행이 자동 추가됩니다. 협상 금액으로 단가 조정 가능."):
            _add_discount_row()
            st.rerun()
    with reset_col:
        if st.button("전체 초기화", use_container_width=True):
            _reset_items()
            st.rerun()

    if st.session_state.items_df.empty:
        st.info("아직 품목이 없습니다. 위 드롭다운에서 카탈로그 상품을 추가하거나 '+ 빈 행' 을 누르세요.")
        edited_df = st.session_state.items_df
    else:
        # 행 순서 변경 — 행 선택 후 ↑/↓ 버튼으로 이동
        df_for_order = st.session_state.items_df.reset_index(drop=True)
        if len(df_for_order) > 1:
            order_label, order_up, order_down = st.columns([5, 0.7, 0.7])
            with order_label:
                row_options = list(range(len(df_for_order)))
                sel_row = st.selectbox(
                    "🔃 순서 변경할 행",
                    options=row_options,
                    format_func=lambda i: (
                        f"{i + 1}. {df_for_order.at[i, '항목'] or '(이름 없음)'}"
                    ),
                    index=None, placeholder="행을 선택해 위/오른쪽 버튼으로 이동...",
                    key="row_reorder_sel",
                    label_visibility="collapsed",
                )
            with order_up:
                if st.button("⬆ 위로", use_container_width=True,
                             disabled=sel_row is None or sel_row == 0,
                             key="row_move_up"):
                    df = st.session_state.items_df.reset_index(drop=True)
                    df.iloc[[sel_row - 1, sel_row]] = df.iloc[[sel_row, sel_row - 1]].values
                    st.session_state.items_df = df
                    st.session_state["row_reorder_sel"] = sel_row - 1
                    st.rerun()
            with order_down:
                if st.button("⬇ 아래로", use_container_width=True,
                             disabled=sel_row is None or sel_row == len(df_for_order) - 1,
                             key="row_move_down"):
                    df = st.session_state.items_df.reset_index(drop=True)
                    df.iloc[[sel_row + 1, sel_row]] = df.iloc[[sel_row, sel_row + 1]].values
                    st.session_state.items_df = df
                    st.session_state["row_reorder_sel"] = sel_row + 1
                    st.rerun()

        # 공급가 컬럼을 계산해서 디스플레이용 DataFrame 생성
        display_df = st.session_state.items_df.copy()
        # 분류 값이 비었으면 기본 '품목'으로 채움
        display_df["분류"] = display_df["분류"].fillna(ITEM_KIND_NORMAL).replace("", ITEM_KIND_NORMAL)
        display_df["공급가"] = display_df.apply(_row_amount, axis=1).astype("Int64")
        display_df = display_df[DISPLAY_COLUMNS]

        edited_df = st.data_editor(
            display_df,
            column_config={
                "분류": st.column_config.SelectboxColumn(
                    "분류", options=ITEM_KINDS, required=True, width="small",
                    help="'💰 할인' 선택 시 단가는 음수로 입력. 할인 행은 아래 빨강 박스에 음영 처리됩니다.",
                ),
                "항목": st.column_config.TextColumn("항목", required=True, width="medium"),
                "설명": st.column_config.TextColumn("설명", width="large"),
                "단가": st.column_config.NumberColumn(
                    "단가", step=1000, format="₩%,d", width="small",
                    help="단가 (양수). '💰 할인' 분류는 자동 차감됩니다.",
                ),
                "기간(횟수)": st.column_config.NumberColumn(
                    "기간(횟수)", min_value=0, step=1, width="small",
                ),
                "수량": st.column_config.NumberColumn(
                    "수량", min_value=0, step=1, width="small",
                ),
                "할인율(%)": st.column_config.NumberColumn(
                    "할인율(%)", min_value=0, max_value=100, step=1, format="%d%%",
                    width="small",
                    help="항목별 할인율 (0~100). 할인금액이 입력되면 할인금액이 우선됩니다.",
                ),
                "할인금액": st.column_config.NumberColumn(
                    "할인금액", min_value=0, step=1000, format="₩%,d", width="small",
                    help="항목별 할인 금액 (양수). 비워두면 할인율이 적용됩니다.",
                ),
                "공급가": st.column_config.NumberColumn(
                    "공급가", disabled=True, format="₩%,d", width="small",
                    help="(수량 × 기간 × 단가) − 항목별 할인",
                ),
                "비고": st.column_config.TextColumn("비고", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="items_editor",
        )
        # 편집 가능 컬럼만 세션에 저장
        edited_core = edited_df[ITEM_COLUMNS]
        # 공급가가 입력 변경에 따라 즉시 갱신되도록 — 변경 감지 시 rerun
        old = st.session_state.items_df.reset_index(drop=True)
        new = edited_core.reset_index(drop=True)
        changed = (
            len(old) != len(new) or
            not all((old[c].astype(object).fillna("__").tolist()
                     == new[c].astype(object).fillna("__").tolist())
                    for c in ITEM_COLUMNS)
        )
        st.session_state.items_df = edited_core
        if changed:
            st.rerun()

        # ── 할인 행 음영 미리보기 (분류=할인인 행만 빨강 박스에 다시 표시) ──
        disc_mask = edited_df["분류"].astype(str) == ITEM_KIND_DISCOUNT
        disc_rows = edited_df[disc_mask]
        if not disc_rows.empty:
            disc_amounts = disc_rows.apply(_row_amount, axis=1)
            disc_total = int(pd.to_numeric(disc_amounts, errors="coerce").fillna(0).sum())
            st.markdown(
                f"""
<div style="background:#FDECEA; border-left:4px solid #C0392B;
            border-radius:6px; padding:10px 14px; margin:6px 0 4px;">
  <div style="color:#C0392B; font-weight:700; font-size:0.95rem;">
    💰 할인 적용 내역 · {len(disc_rows)}건 · 차감 합계 <span style="font-size:1.05rem">₩{abs(disc_total):,}</span>
  </div>
  <div style="color:#7B241C; font-size:0.82rem; margin-top:3px;">
    아래 행은 할인 분류로 분류된 행입니다. PDF에서도 빨강으로 음영·강조 표시됩니다.
  </div>
</div>
                """,
                unsafe_allow_html=True,
            )
            disc_preview = disc_rows[["항목", "설명", "단가", "공급가", "비고"]].copy()
            st.dataframe(
                disc_preview,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "단가": st.column_config.NumberColumn("단가", format="₩%,d"),
                    "공급가": st.column_config.NumberColumn("공급가", format="₩%,d"),
                },
            )

    if not edited_df.empty:
        amounts = edited_df.apply(_row_amount, axis=1)
        items_sum = int(pd.to_numeric(amounts, errors="coerce").fillna(0).sum())

        st.divider()
        st.caption("📊 **공급가액 · 부가세 · 합계 금액**")
        td_col, _ = st.columns([1, 3])
        with td_col:
            total_discount_pct = st.number_input(
                "전체 일괄 할인율 (%)",
                min_value=0, max_value=100, step=1,
                value=int(st.session_state.get("total_discount_pct", 0)),
                key="total_discount_pct",
                help="공급가액에 일괄 적용되는 할인율 (항목별 할인은 별도 행에서 적용).",
            )
        total_discount_value = int(round(items_sum * total_discount_pct / 100))
        subtotal = items_sum - total_discount_value

        vat_rate = labels.quote.vat_rate
        vat = int(round(subtotal * vat_rate))
        total = subtotal + vat

        if total_discount_pct > 0:
            m0, m1, m2, m3 = st.columns(4)
            m0.metric("일괄 할인", f"-₩{total_discount_value:,}",
                      delta=f"{total_discount_pct}%", delta_color="inverse")
            m1.metric("공급가액", f"₩{subtotal:,}")
            m2.metric(f"부가세 ({int(vat_rate * 100)}%)", f"₩{vat:,}")
            m3.metric("합계 금액", f"₩{total:,}")
        else:
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
        total_discount_rate=(
            float(st.session_state.get("total_discount_pct", 0)) / 100
            if int(st.session_state.get("total_discount_pct", 0) or 0) > 0
            else None
        ),
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

    # 작성 내용 자동저장 + 페이지 이탈 시 브라우저 경고
    _qr_autosave_write()
    df = st.session_state.get("items_df")
    has_data = (df is not None and not df.empty) or any(
        st.session_state.get(k) for k in QR_FORM_KEYS
    )
    _inject_beforeunload(has_data)


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
                           soffice_available, total_discount_rate=None,
                           status_label="문서 생성 중..."):
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
        name = row.get("항목")
        name = name.strip() if isinstance(name, str) else None
        if not name:
            continue
        qty = row.get("수량")
        period = row.get("기간(횟수)")
        unit_price = float(row.get("단가") or 0)
        is_discount_row = row.get("분류") == ITEM_KIND_DISCOUNT
        # 분류='💰 할인' 이면 단가 양수 입력도 음수로 강제 (자동 차감)
        if is_discount_row:
            unit_price = -abs(unit_price)
        desc_val = row.get("설명")
        notes_val = row.get("비고")
        disc_rate = row.get("할인율(%)")
        disc_amt = row.get("할인금액")
        items.append(LineItem(
            name=name,
            description=desc_val if isinstance(desc_val, str) and desc_val else None,
            qty=float(qty) if pd.notna(qty) and qty else None,
            period=float(period) if pd.notna(period) and period else None,
            unit_price=unit_price,
            discount_rate=(float(disc_rate) / 100
                           if pd.notna(disc_rate) and disc_rate
                           and not is_discount_row else None),
            discount_amount=(float(disc_amt)
                             if pd.notna(disc_amt) and disc_amt
                             and not is_discount_row else None),
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
        total_discount_rate=total_discount_rate,
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

    # 최근 견적서 히스토리에 저장 (입력 데이터 그대로 복원 가능)
    _qr_save_history(_qr_snapshot_payload(), document_id)

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
        "견적서 작성 화면의 '카탈로그 빠른 추가' 드롭다운에 사용되는 상품 목록입니다. "
        "탭으로 견적서 종류별로 관리할 수 있어요."
    )

    tab_qr, tab_mc = st.tabs(["📋 QR 견적서 카탈로그", "🏢 멤버십 견적서 카탈로그"])
    with tab_qr:
        _render_qr_catalog_editor()
    with tab_mc:
        _render_membership_catalog_editor()


def _render_qr_catalog_editor():
    """QR 견적서용 평면 카탈로그 편집기."""
    products = _load_products()
    if not products:
        df = pd.DataFrame([{
            "code": "NEW-CODE",
            "name": "(여기를 클릭해서 수정)",
            "description": "",
            "unit_price": 0,
            "currency": "KRW",
        }])
    else:
        df = pd.DataFrame(products)
        for col, default in [("description", ""), ("currency", "KRW")]:
            if col not in df.columns:
                df[col] = default

    st.caption(
        "💡 설명/청구기준은 셀 안에서 Alt+Enter로 줄바꿈 입력 가능합니다. "
        "행 높이를 크게 잡아 여러 줄이 보이게 표시하며, 견적서 PDF에도 "
        "줄바꿈이 그대로 반영됩니다."
    )

    edited = st.data_editor(
        df[["code", "name", "description", "unit_price", "currency"]],
        column_config={
            "code": st.column_config.TextColumn(
                "코드", help="고유 식별자 (영문/숫자/하이픈)", required=True,
                width="small",
            ),
            "name": st.column_config.TextColumn(
                "품목명", help="견적서에 표시되는 이름", required=True,
                width="medium",
            ),
            "description": st.column_config.TextColumn(
                "설명/청구기준",
                help=("Alt+Enter 로 줄바꿈 입력. "
                      "여러 줄 입력하면 표·PDF 모두 줄바꿈 그대로 반영됩니다."),
                width="large",
            ),
            "unit_price": st.column_config.NumberColumn(
                "단가 (원)", min_value=0, step=1000, format="₩%,d",
                width="small",
            ),
            "currency": st.column_config.SelectboxColumn(
                "통화", options=["KRW", "USD", "EUR", "JPY"],
                width="small",
            ),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=720,
        key="catalog_editor_qr",
    )

    st.divider()
    col_save, col_info = st.columns([1, 3])
    with col_save:
        if st.button("💾 QR 카탈로그 저장", type="primary",
                     use_container_width=True, key="save_qr_catalog"):
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
                       "'QR 견적서 작성' 페이지로 가면 즉시 반영됩니다.")
    with col_info:
        st.caption(
            "📝 **저장 후**: 견적서 작성 페이지의 카탈로그가 즉시 갱신됩니다. "
            "엑셀 템플릿의 드롭다운은 `template` 명령으로 재생성해야 반영됩니다."
        )


def _render_membership_catalog_editor():
    """멤버십 견적서용 계층 카탈로그 편집기 (구분/분류 컬럼 포함)."""
    st.caption(
        "구분(예: 멤버십 클라우드)과 분류(초기구축비/사용료/옵션)로 묶인 상품 목록입니다. "
        "멤버십 견적서 작성 화면에서 '구분' 이 일치하는 항목만 드롭다운에 나타납니다."
    )

    products = _load_membership_products()
    if not products:
        df = pd.DataFrame([{
            "code": "NEW-CODE",
            "section": "멤버십 클라우드",
            "subcategory": "초기구축비",
            "name": "(여기를 클릭해서 수정)",
            "billing_period": "1회성",
            "unit_price": 0,
            "unit_price_text": "",
            "default_amount_text": "",
            "notes": "",
        }])
        original_extra = {}
    else:
        df = pd.DataFrame(products)
        # 사용자 표에서 다루지 않을 필드(sub_items, name_detail 등)는 보존
        original_extra = {p["code"]: {k: v for k, v in p.items()
                                       if k not in {"code", "section", "subcategory",
                                                    "name", "billing_period", "unit_price",
                                                    "unit_price_text", "default_amount_text",
                                                    "notes"}}
                          for p in products if p.get("code")}
        # 누락 가능한 컬럼 보완
        for col, default in [
            ("section", ""), ("subcategory", ""), ("billing_period", ""),
            ("unit_price", None), ("unit_price_text", ""),
            ("default_amount_text", ""), ("notes", ""),
        ]:
            if col not in df.columns:
                df[col] = default

    # 기존 카탈로그에 등록된 구분들 + 기본값
    existing_sections = sorted({p.get("section", "") for p in products if p.get("section")})
    section_options = existing_sections or ["멤버십 클라우드", "오더 솔루션"]

    edited = st.data_editor(
        df[["code", "section", "subcategory", "name", "billing_period",
            "unit_price", "unit_price_text", "default_amount_text", "notes"]],
        column_config={
            "code": st.column_config.TextColumn(
                "코드", help="고유 식별자 (예: MC-INIT-SERVER)", required=True,
            ),
            "section": st.column_config.SelectboxColumn(
                "구분", help="대분류 (시나리오 내 그룹)",
                options=section_options + ["기타"],
            ),
            "subcategory": st.column_config.SelectboxColumn(
                "분류", help="중분류",
                options=_DEFAULT_SUBCATEGORIES + ["기타"],
            ),
            "name": st.column_config.TextColumn(
                "상세 구분", help="견적서에 표시되는 항목명", required=True,
            ),
            "billing_period": st.column_config.SelectboxColumn(
                "기간", options=["", "1회성", "매월", "발생시", "발생월", "1개당"],
                help="청구 주기",
            ),
            "unit_price": st.column_config.NumberColumn(
                "단가 (숫자)", step=100000, format="₩%,d",
                help="숫자 단가. 할인은 음수로 (예: -1000000). 텍스트는 옆 칸 사용.",
            ),
            "unit_price_text": st.column_config.TextColumn(
                "단가(텍스트)",
                help="예: '투입기간 X SW개발자 임금', '무상 제공', '7.9원 / 건'",
            ),
            "default_amount_text": st.column_config.TextColumn(
                "기본 금액 텍스트",
                help="비우면 자동 계산. '후청구', '협의 금액', '무상제공' 등 텍스트 입력 가능",
            ),
            "notes": st.column_config.TextColumn(
                "비고", width="large",
                help="Alt+Enter 로 줄바꿈 입력 가능. PDF에도 줄바꿈 반영.",
            ),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=720,
        key="catalog_editor_mc",
    )

    st.divider()
    col_save, col_info = st.columns([1, 3])
    with col_save:
        if st.button("💾 멤버십 카탈로그 저장", type="primary",
                     use_container_width=True, key="save_mc_catalog"):
            new_products = []
            for _, row in edited.iterrows():
                code = (row.get("code") or "").strip()
                name = (row.get("name") or "").strip()
                if not code or not name:
                    continue
                item: dict = {
                    "code": code,
                    "section": (row.get("section") or "").strip(),
                    "subcategory": (row.get("subcategory") or "").strip(),
                    "name": name,
                }
                bp = row.get("billing_period")
                if isinstance(bp, str) and bp.strip():
                    item["billing_period"] = bp.strip()
                up = row.get("unit_price")
                if pd.notna(up) and up not in ("", None):
                    try:
                        item["unit_price"] = float(up)
                    except (TypeError, ValueError):
                        item["unit_price"] = None
                else:
                    item["unit_price"] = None
                upt = row.get("unit_price_text")
                if isinstance(upt, str) and upt.strip():
                    item["unit_price_text"] = upt.strip()
                amt_text = row.get("default_amount_text")
                if isinstance(amt_text, str) and amt_text.strip():
                    item["default_amount_text"] = amt_text.strip()
                notes = row.get("notes")
                if isinstance(notes, str) and notes.strip():
                    item["notes"] = notes.strip()
                # 기존 sub_items / name_detail 등 추가 필드는 보존
                extra = original_extra.get(code, {})
                item.update(extra)
                new_products.append(item)
            _save_membership_products(new_products)
            st.success(
                f"✅ {len(new_products)}개 멤버십 상품 저장 완료. "
                "'멤버십 견적서 작성' 페이지로 가면 즉시 반영됩니다."
            )
    with col_info:
        st.caption("")


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
# 페이지 4: 멤버십 클라우드 견적서 작성
# ═════════════════════════════════════════════════════════════

# 분류 미리보기 옵션 (사용자가 직접 입력도 가능)
# '할인' 은 음수 단가로 등록 → 총비용에서 자동 차감됨
_DEFAULT_SUBCATEGORIES = ["초기구축비", "사용료", "옵션", "할인"]


@st.cache_data
def _load_membership_products() -> list[dict]:
    path = PROJECT_ROOT / "catalog" / "membership_products.json"
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8")).get("products", [])


def _save_membership_products(products: list[dict]) -> None:
    path = PROJECT_ROOT / "catalog" / "membership_products.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"products": products}, ensure_ascii=False, indent=2),
                    encoding="utf-8")
    _load_membership_products.clear()


def _load_membership_sample() -> dict:
    """샘플 데이터 (사용자가 '샘플 불러오기' 버튼을 눌렀을 때 사용)."""
    path = PROJECT_ROOT / "data" / "membership_quotes" / "sample.json"
    if path.exists():
        sample = json.loads(path.read_text(encoding="utf-8"))
        sample.setdefault("issued_date", date.today().isoformat())
        return sample
    return _empty_membership_state()


def _empty_membership_state() -> dict:
    """빈 초기 상태 — QR 견적서와 동일하게 placeholder만 보이는 상태로 시작.
    supplier(회사 정보)는 brand.json에서 자동 로드되므로 None."""
    return {
        "document_id": "",
        "title": "멤버십 클라우드 견적서",
        "issued_date": date.today().isoformat(),
        "counterparty": {
            "label": "제휴사", "name": "",
            "address": None, "ceo": None, "contact": None,
        },
        "supplier": None,
        "scenarios": [],
        "remarks": [
            "견적유효 : 견적일로부터 15일",
            "결제조건 : 현금결제 (귀사 결제조건)",
        ],
    }


def _ensure_membership_state():
    if "mc_doc" not in st.session_state:
        st.session_state.mc_doc = _empty_membership_state()


def _mc_items_to_df(items: list[dict]) -> pd.DataFrame:
    """[{name, billing_period, unit_price, ...}, ...] → DataFrame."""
    if not items:
        return pd.DataFrame(columns=[
            "분류", "상세 구분", "기간", "단가", "단가(텍스트)",
            "할인율(%)", "할인금액", "금액 텍스트", "비고",
        ])
    rows = []
    for it in items:
        dr = it.get("discount_rate")
        # 분류='할인' 행은 표 표시 시 단가를 양수로 노출 (자동 차감 처리는 저장 단계에서)
        up = it.get("unit_price")
        sub = (it.get("_subcategory") or "").strip()
        if sub == "할인" and up is not None:
            try:
                up = abs(float(up))
            except (TypeError, ValueError):
                pass
        rows.append({
            "분류": sub,
            "상세 구분": it.get("name", ""),
            "기간": it.get("billing_period", ""),
            "단가": up,
            "단가(텍스트)": it.get("unit_price_text", ""),
            "할인율(%)": int(dr * 100) if dr else None,
            "할인금액": it.get("discount_amount"),
            "금액 텍스트": it.get("amount_text", ""),
            "비고": it.get("notes", ""),
        })
    return pd.DataFrame(rows)


def _df_to_section_categories(df: pd.DataFrame) -> list[dict]:
    """편집된 DataFrame → categories 리스트 (분류별 그룹).
    설정된 분류가 없는 행은 '기타' 분류로 묶음.
    """
    if df.empty:
        return []
    # 분류별 그룹 유지 순서
    seen_order = []
    groups: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        name = (row.get("상세 구분") or "").strip() if isinstance(row.get("상세 구분"), str) else ""
        if not name:
            continue
        sub = (row.get("분류") or "").strip() if isinstance(row.get("분류"), str) else ""
        if not sub:
            sub = "기타"
        if sub not in groups:
            groups[sub] = []
            seen_order.append(sub)
        item: dict = {"name": name}
        billing = row.get("기간")
        if isinstance(billing, str) and billing.strip():
            item["billing_period"] = billing.strip()
        up = row.get("단가")
        if pd.notna(up) and up not in ("", None):
            try:
                up_val = float(up)
                # 분류='할인' 이면 단가 양수 입력도 음수로 강제 (자동 차감)
                if sub == "할인":
                    up_val = -abs(up_val)
                item["unit_price"] = up_val
            except (TypeError, ValueError):
                pass
        upt = row.get("단가(텍스트)")
        if isinstance(upt, str) and upt.strip():
            item["unit_price_text"] = upt.strip()
        dr = row.get("할인율(%)")
        if pd.notna(dr) and dr not in ("", None):
            try:
                drv = float(dr)
                if drv > 0:
                    item["discount_rate"] = drv / 100
            except (TypeError, ValueError):
                pass
        da = row.get("할인금액")
        if pd.notna(da) and da not in ("", None):
            try:
                dav = float(da)
                if dav > 0:
                    item["discount_amount"] = dav
            except (TypeError, ValueError):
                pass
        amt_text = row.get("금액 텍스트")
        if isinstance(amt_text, str) and amt_text.strip():
            item["amount_text"] = amt_text.strip()
        notes = row.get("비고")
        if isinstance(notes, str) and notes.strip():
            item["notes"] = notes.strip()
        groups[sub].append(item)

    cats = []
    for sub in seen_order:
        cats.append({
            "name": sub,
            "items": groups[sub],
            "show_subtotal": sub != "옵션",  # 옵션은 합계 미표시 기본
        })
    return cats


def _section_to_flat_items(section: dict) -> list[dict]:
    """section의 categories→items 를 평면화. _subcategory 필드 추가."""
    flat = []
    for cat in section.get("categories", []):
        for it in cat.get("items", []):
            flat.append({**it, "_subcategory": cat.get("name", "")})
    return flat


def _add_blank_row_to_section(section: dict,
                              default_subcategory: str = "초기구축비") -> None:
    """빈 항목 한 줄 추가. 사용자가 수정해서 진짜 항목으로 만듦."""
    section.setdefault("categories", [])
    target_cat = None
    for cat in section["categories"]:
        if cat.get("name") == default_subcategory:
            target_cat = cat
            break
    if target_cat is None:
        target_cat = {
            "name": default_subcategory,
            "items": [],
            "show_subtotal": default_subcategory != "옵션",
        }
        section["categories"].append(target_cat)
    target_cat.setdefault("items", []).append({
        "name": "(새 항목 — 클릭해서 수정)",
    })


def _add_discount_row_to_section(section: dict) -> None:
    """'할인' 분류에 할인 행 추가. 단가는 양수 입력 시 자동 차감."""
    section.setdefault("categories", [])
    target_cat = None
    for cat in section["categories"]:
        if cat.get("name") == "할인":
            target_cat = cat
            break
    if target_cat is None:
        target_cat = {
            "name": "할인",
            "items": [],
            "show_subtotal": True,
        }
        section["categories"].append(target_cat)
    target_cat.setdefault("items", []).append({
        "name": "협상 할인 (이름 수정 가능)",
        "billing_period": "1회성",
        "unit_price": 1000000,
        "notes": "협상에 따라 단가 조정",
    })


def _build_mc_document(state: dict) -> MembershipQuoteDocument:
    """session_state.mc_doc 를 Pydantic 객체로 변환."""
    return MembershipQuoteDocument.model_validate(state)


def render_membership_quote_page():
    _mc_autosave_load_once()
    _ensure_membership_state()
    state = st.session_state.mc_doc
    products = _load_membership_products()
    soffice_available = find_soffice() is not None

    # 사이드바: 발행 정보
    with st.sidebar:
        st.divider()
        st.header("⚙ 발행 정보")
        try:
            cur_issued = date.fromisoformat(state.get("issued_date", date.today().isoformat()))
        except (TypeError, ValueError):
            cur_issued = date.today()
        new_issued = st.date_input("발행일", value=cur_issued, key="mc_issued_date")
        state["issued_date"] = new_issued.isoformat()
        if not soffice_available:
            st.warning("⚠ LibreOffice 미감지 — PDF 변환 건너뜀")

    st.title("🏢 멤버십 클라우드 견적서 작성")

    # ─── 0. 발행 담당자 정보 ───
    try:
        _brand_for_default = load_brand(PROJECT_ROOT, "softment")
        _default_cp = _brand_for_default.contact_person
    except FileNotFoundError:
        _default_cp = None

    st.subheader("0. 발행 담당자 정보")
    st.caption("이 견적서를 발행하는 우리 쪽 담당자. 회사(우리)의 '담당자/연락처/이메일' 칸에 표시됩니다.")
    # Tab 좌→우→다음 줄 순으로 이동하도록 행 단위 columns
    ic_r1_l, ic_r1_r = st.columns(2)
    with ic_r1_l:
        issuer_name = st.text_input(
            "담당자명", key="mc_issuer_name",
            value=(_default_cp.name if _default_cp else ""),
            placeholder="예: 박지은",
        )
    with ic_r1_r:
        issuer_title = st.text_input(
            "직책", key="mc_issuer_title",
            value=(_default_cp.title if _default_cp and _default_cp.title else ""),
            placeholder="예: QR사업부 매니저",
        )
    ic_r2_l, ic_r2_r = st.columns(2)
    with ic_r2_l:
        issuer_phone = st.text_input(
            "연락처", key="mc_issuer_phone",
            value=(_default_cp.phone if _default_cp and _default_cp.phone else ""),
            placeholder="예: 010-0000-0000",
        )
    with ic_r2_r:
        issuer_email = st.text_input(
            "이메일", key="mc_issuer_email",
            value=(_default_cp.email if _default_cp and _default_cp.email else ""),
            placeholder="예: name@softment.co.kr",
        )

    # ─── 1. 수신처 정보 ───
    st.subheader("1. 수신처 정보")
    st.caption("회사(우리) 기본 정보는 '⚙ 설정' 의 브랜드 정보에서 자동으로 가져옵니다.")

    cp = state.setdefault("counterparty", {"label": "제휴사", "name": ""})
    # supplier 는 PDF 렌더 시 brand.json 에서 자동 채워지므로 폼에서 제거
    state["supplier"] = None

    # Tab 좌→우→다음 줄 순으로
    pc_r1_l, pc_r1_r = st.columns(2)
    with pc_r1_l:
        cp["name"] = st.text_input(
            "회사명 *", value=cp.get("name") or "", key="mc_cp_name",
            placeholder="예: 주식회사 ○○",
        )
    with pc_r1_r:
        cp["address"] = st.text_input(
            "주소", value=cp.get("address") or "", key="mc_cp_addr",
            placeholder="예: 서울특별시 ○○구 ○○로 ○○",
        )
    pc_r2_l, pc_r2_r = st.columns(2)
    with pc_r2_l:
        cp["ceo"] = st.text_input(
            "대표이사", value=cp.get("ceo") or "", key="mc_cp_ceo",
            placeholder="예: 홍길동",
        )
    with pc_r2_r:
        cp["contact"] = st.text_input(
            "담당자", value=cp.get("contact") or "", key="mc_cp_contact",
            placeholder="예: 김담당 (구매팀장)",
        )

    # ─── 2. 건명 ───
    st.subheader("2. 건명")
    state["title"] = st.text_input(
        "건명",
        value=state.get("title", "") or "",
        key="mc_title",
        placeholder="예: 멤버십 클라우드 견적서",
        label_visibility="collapsed",
    )

    # ─── 3. 품목 내역 (시나리오 탭) ───
    st.subheader("3. 품목 내역")
    st.caption(
        "💡 한 견적서 안에 여러 시나리오를 넣어 비교 견적을 제공할 수 있어요. "
        "예: '앱+POS 연동' 시나리오와 'POS만' 시나리오를 한 문서에 묶어 발행."
    )
    scenarios = state.setdefault("scenarios", [])

    sc_btn_add, sc_btn_sample, sc_btn_reset, _ = st.columns([1.2, 1.4, 1.4, 3])
    with sc_btn_add:
        if st.button("+ 시나리오 추가", use_container_width=True):
            scenarios.append({
                "name": f"시나리오 {len(scenarios) + 1}",
                "subject": "",
                "sections": [],
                "show_grand_total": True,
            })
            st.rerun()
    with sc_btn_sample:
        if st.button("📋 (간편)샘플 데이터로 채우기", use_container_width=True,
                     help="예시 데이터(시나리오 2개)로 폼을 채워봅니다. 현재 입력은 덮어써져요."):
            st.session_state.mc_doc = _load_membership_sample()
            st.rerun()
    with sc_btn_reset:
        if st.button("🗑 전체 초기화", use_container_width=True,
                     help="제휴사·회사·시나리오까지 모든 입력을 비웁니다."):
            st.session_state.mc_doc = _empty_membership_state()
            for k in ("mc_issuer_name", "mc_issuer_title",
                      "mc_issuer_phone", "mc_issuer_email"):
                if k in st.session_state:
                    del st.session_state[k]
            _mc_autosave_clear()
            st.rerun()

    if not scenarios:
        st.info(
            "아직 시나리오가 없어요. **+ 시나리오 추가** 를 눌러 시작하거나, "
            "**📋 (간편)샘플 데이터로 채우기** 로 예시 구조를 먼저 확인할 수 있어요."
        )
    else:
        tab_labels = [(sc.get("name") or f"시나리오 {i+1}") for i, sc in enumerate(scenarios)]
        tabs = st.tabs(tab_labels)
        for s_idx, (tab, scenario) in enumerate(zip(tabs, scenarios)):
            with tab:
                _render_scenario_editor(s_idx, scenario, products)

    # ─── 3. Remarks ───
    st.subheader("4. 기타 안내")
    remarks = state.setdefault("remarks", [])
    remarks_text = st.text_area(
        "한 줄에 하나씩",
        value="\n".join(remarks),
        height=80,
        key="mc_remarks",
        label_visibility="collapsed",
    )
    state["remarks"] = [ln.strip() for ln in remarks_text.splitlines() if ln.strip()]

    # ─── 4. 생성 / 미리보기 ───
    st.divider()
    can_generate = bool(cp.get("name")) and bool(scenarios)
    if not can_generate:
        st.info("제휴사 회사명과 시나리오 최소 1개를 입력해주세요.")

    issuer_contact = dict(
        name=issuer_name.strip(),
        title=issuer_title.strip() or None,
        phone=issuer_phone.strip() or None,
        email=issuer_email.strip() or None,
    )

    btn_preview, btn_generate = st.columns(2)
    with btn_preview:
        preview_disabled = (not can_generate) or (not soffice_available)
        preview_clicked = st.button(
            "👁 미리보기 (PDF)", use_container_width=True,
            disabled=preview_disabled,
            help=(
                "현재 입력값으로 PDF 미리보기를 표시합니다."
                if soffice_available
                else "PDF 변환기(LibreOffice) 미감지 — 미리보기 불가"
            ),
            key="mc_btn_preview",
        )
    with btn_generate:
        generate_clicked = st.button(
            "📝 멤버십 견적서 생성", type="primary",
            use_container_width=True, disabled=not can_generate,
            key="mc_btn_generate",
        )

    if preview_clicked:
        _preview_membership_quote(state, soffice_available, issuer_contact)
    if generate_clicked:
        _generate_membership_quote(state, soffice_available, issuer_contact)

    # 자동저장 + 페이지 이탈 시 브라우저 경고
    _mc_autosave_write()
    doc_has_data = bool(state.get("scenarios")) or bool(state.get("counterparty", {}).get("name"))
    issuer_has_data = any(
        st.session_state.get(k)
        for k in ("mc_issuer_name", "mc_issuer_title", "mc_issuer_phone", "mc_issuer_email")
    )
    _inject_beforeunload(doc_has_data or issuer_has_data)


def _render_scenario_editor(s_idx: int, scenario: dict, products: list[dict]) -> None:
    """한 시나리오 탭 안의 편집기."""
    sc1, sc2 = st.columns([5, 1])
    with sc1:
        scenario["name"] = st.text_input(
            "시나리오 이름", value=scenario.get("name", ""),
            key=f"sc_name_{s_idx}",
            placeholder="예: 앱&POS 연동 (멤버십+오더)",
        )
    with sc2:
        if st.button("❌ 시나리오 삭제", key=f"sc_del_{s_idx}",
                     use_container_width=True):
            st.session_state.mc_doc["scenarios"].pop(s_idx)
            st.rerun()

    scenario["subject"] = st.text_input(
        "항목 부제 (선택)", value=scenario.get("subject") or "",
        key=f"sc_subject_{s_idx}",
        placeholder="예: 멤버십 클라우드 솔루션 + 오더 솔루션",
    )

    # 구분(섹션) 목록
    sections = scenario.setdefault("sections", [])

    sec_btn1, sec_btn2 = st.columns([1, 5])
    with sec_btn1:
        if st.button(f"+ 구분 추가", key=f"sec_add_{s_idx}",
                     use_container_width=True):
            sections.append({
                "name": "새 구분",
                "categories": [],
                "show_section_total": True,
            })
            st.rerun()

    if not sections:
        st.info("이 시나리오에 구분이 없습니다. '+ 구분 추가' 를 눌러 추가하세요.")
        return

    for sec_idx, section in enumerate(sections):
        with st.expander(f"📂 {section.get('name', '(이름 없음)')}",
                         expanded=True):
            _render_section_editor(s_idx, sec_idx, section, products)

    # 시나리오 합계 — 공급가액 / 부가세 / 합계 금액 (샘플 로드 시에도 항상 정확 반영)
    try:
        sc_obj = MembershipScenario.model_validate(scenario)
    except Exception as e:  # noqa: BLE001
        st.warning(f"⚠ 시나리오 데이터 검증 실패: {e}")
        return
    try:
        vat_rate = load_labels(PROJECT_ROOT).quote.vat_rate
    except Exception:  # noqa: BLE001
        vat_rate = 0.10

    # 전체 일괄 할인율 입력 (문서 단위; 모든 시나리오에 공통 적용)
    st.divider()
    st.caption("📊 **이 시나리오의 합계** (모든 기간 합산, 할인 차감 반영)")
    td_col, _ = st.columns([1, 3])
    with td_col:
        td_pct = st.number_input(
            "전체 일괄 할인율 (%)",
            min_value=0, max_value=100, step=1,
            value=int((st.session_state.mc_doc.get("total_discount_rate") or 0) * 100),
            key=f"mc_total_discount_pct_{s_idx}",
            help="공급가액 합계에 일괄 적용. 항목별 할인은 별도 행에서 입력합니다.",
        )
    st.session_state.mc_doc["total_discount_rate"] = (td_pct / 100) if td_pct > 0 else None
    td_rate = td_pct / 100 if td_pct > 0 else 0
    items_sum = sum(scenario_vat_and_total(sc_obj, vat_rate, None)[0:1])  # 공급가액 (할인 전)
    items_sum = scenario_vat_and_total(sc_obj, vat_rate, None)[0]
    subtotal, vat, total = scenario_vat_and_total(sc_obj, vat_rate, td_rate)
    td_value = items_sum - subtotal

    def _money_label(v: float) -> str:
        val = int(round(v))
        return f"-₩{abs(val):,}" if val < 0 else f"₩{val:,}"

    if td_pct > 0:
        m0, m1, m2, m3 = st.columns(4)
        m0.metric("일괄 할인", f"-₩{int(round(td_value)):,}",
                  delta=f"{td_pct}%", delta_color="inverse")
        m1.metric("공급가액", _money_label(subtotal))
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", _money_label(vat))
        m3.metric("합계 금액", _money_label(total))
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("공급가액", _money_label(subtotal))
        m2.metric(f"부가세 ({int(vat_rate * 100)}%)", _money_label(vat))
        m3.metric("합계 금액", _money_label(total))


def _render_section_editor(s_idx: int, sec_idx: int, section: dict,
                           products: list[dict]) -> None:
    sec1, sec2 = st.columns([5, 1])
    with sec1:
        section["name"] = st.text_input(
            "구분 이름", value=section.get("name", ""),
            key=f"sec_name_{s_idx}_{sec_idx}",
            placeholder="예: 멤버십 클라우드",
        )
    with sec2:
        if st.button("❌ 구분 삭제", key=f"sec_del_{s_idx}_{sec_idx}",
                     use_container_width=True):
            st.session_state.mc_doc["scenarios"][s_idx]["sections"].pop(sec_idx)
            st.rerun()

    # 이 구분의 항목들을 평면 표로 (분류는 컬럼)
    flat_items = _section_to_flat_items(section)
    df = _mc_items_to_df(flat_items)

    # 카탈로그 빠른 추가 (드롭다운) + 빈 행 + 할인 행
    section_name = section.get("name", "")
    matching = [p for p in products if p.get("section") == section_name]
    quick_col1, quick_col2, quick_col3, quick_col4 = st.columns([4, 1.3, 1.0, 1.3])
    with quick_col1:
        if matching:
            options = list(range(len(matching)))
            picked = st.selectbox(
                "카탈로그에서 항목 추가",
                options=options,
                index=None,
                format_func=lambda i: (
                    f"[{matching[i].get('subcategory', '-')}] {matching[i]['name']}"
                ),
                placeholder=f"'{section_name}' 카탈로그에서 항목 선택...",
                key=f"sec_pick_{s_idx}_{sec_idx}",
                label_visibility="collapsed",
            )
        else:
            st.caption(
                f"'{section_name}' 와 매칭되는 카탈로그 항목이 없습니다. "
                "오른쪽 '+ 빈 행' 으로 직접 추가할 수 있어요."
            )
            picked = None
    with quick_col2:
        if st.button("+ 카탈로그 추가", key=f"sec_addrow_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="좌측 드롭다운에서 항목을 선택한 뒤 누르세요"):
            if picked is None or not matching:
                st.warning("⚠ 좌측 드롭다운에서 카탈로그 항목을 먼저 선택해주세요.")
            else:
                p = matching[picked]
                new_row = {
                    "분류": p.get("subcategory", "기타"),
                    "상세 구분": p["name"],
                    "기간": p.get("billing_period", ""),
                    "단가": p.get("unit_price"),
                    "단가(텍스트)": p.get("unit_price_text", ""),
                    "할인율(%)": None,
                    "할인금액": None,
                    "금액 텍스트": p.get("default_amount_text", ""),
                    "비고": p.get("notes", ""),
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                section["categories"] = _df_to_section_categories(df)
                st.rerun()
    with quick_col3:
        if st.button("+ 빈 행", key=f"sec_addblank_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="빈 행을 추가하고 직접 입력합니다"):
            _add_blank_row_to_section(section)
            st.rerun()
    with quick_col4:
        if st.button("💰 + 할인", key=f"sec_adddiscount_{s_idx}_{sec_idx}",
                     use_container_width=True,
                     help="'할인' 분류에 음수 단가 행이 추가됩니다. 협상가로 조정 가능."):
            _add_discount_row_to_section(section)
            st.rerun()

    # 표 편집
    edited = st.data_editor(
        df,
        column_config={
            "분류": st.column_config.SelectboxColumn(
                "분류", options=_DEFAULT_SUBCATEGORIES + ["기타"],
                help="초기구축비/사용료/옵션 등",
            ),
            "상세 구분": st.column_config.TextColumn("상세 구분", required=True),
            "기간": st.column_config.TextColumn("기간",
                help="1회성 / 매월 / 발생시 / 발생월 / 1개당"),
            "단가": st.column_config.NumberColumn(
                "단가 (숫자)", format="₩%,d",
                help="단가 (양수). 분류='할인'은 자동 차감됩니다.",
            ),
            "단가(텍스트)": st.column_config.TextColumn(
                "단가(텍스트)",
                help="숫자로 표현 불가한 단가 (예: '투입기간 X SW개발자 임금')",
            ),
            "할인율(%)": st.column_config.NumberColumn(
                "할인율(%)", min_value=0, max_value=100, step=1, format="%d%%",
                help="항목별 할인율 (0~100). 할인금액 입력 시 할인금액이 우선.",
            ),
            "할인금액": st.column_config.NumberColumn(
                "할인금액", min_value=0, step=1000, format="₩%,d",
                help="항목별 할인 금액 (양수).",
            ),
            "금액 텍스트": st.column_config.TextColumn(
                "금액 텍스트",
                help="비워두면 자동 계산. '후청구', '협의 금액' 등 텍스트도 입력 가능.",
            ),
            "비고": st.column_config.TextColumn("비고"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"sec_editor_{s_idx}_{sec_idx}",
    )

    # 편집 결과를 section에 반영
    section["categories"] = _df_to_section_categories(edited)

    # 소계 미리보기
    try:
        sec_obj = MembershipSection.model_validate(section)
        by_period = section_subtotals_by_period(sec_obj)
        if by_period:
            badge_cols = st.columns(len(by_period) + 1)
            badge_cols[0].caption("**예상 합계**")
            for col, (period, amt) in zip(badge_cols[1:], by_period.items()):
                col.metric(period, f"₩{int(amt):,}")
    except Exception:
        pass


def _build_membership_artifacts(state: dict, soffice_available: bool,
                                issuer_contact: dict | None = None,
                                status_label: str = "멤버십 견적서 생성 중..."):
    """state를 검증·렌더링하여 (document_id, docx_bytes, pdf_bytes) 반환.
    실패 시 None 반환 (에러는 이미 화면에 표시됨)."""
    # 문서번호 비어있으면 자동 생성 (QR 견적서와 동일 패턴)
    if not (state.get("document_id") or "").strip():
        try:
            issued = date.fromisoformat(state.get("issued_date", date.today().isoformat()))
        except (TypeError, ValueError):
            issued = date.today()
        cp_name = ((state.get("counterparty") or {}).get("name") or "고객").strip()
        cp_short = "".join(c for c in cp_name if c.isalnum())[:8] or "고객"
        state["document_id"] = f"MC-{issued.strftime('%Y%m%d')}-{cp_short}"
    try:
        document = MembershipQuoteDocument.model_validate(state)
    except Exception as e:
        st.error(f"❌ 데이터 검증 실패: {e}")
        return None

    try:
        brand = load_brand(PROJECT_ROOT, document.brand_id)
    except FileNotFoundError as e:
        st.error(f"❌ 브랜드 로드 실패: {e}")
        return None

    # 발행 담당자 정보 — 폼에서 입력한 값으로 brand.contact_person 일회용 override
    if issuer_contact and (issuer_contact.get("name") or "").strip():
        brand = brand.model_copy(update={
            "contact_person": ContactPerson(
                name=issuer_contact["name"].strip(),
                title=issuer_contact.get("title"),
                phone=issuer_contact.get("phone"),
                email=issuer_contact.get("email"),
            )
        })

    output_dir = PROJECT_ROOT / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    docx_path = output_dir / f"{document.document_id}.docx"

    with st.status(status_label, expanded=True) as status:
        st.write("📝 DOCX 생성 중...")
        try:
            render_membership_docx(brand, document, PROJECT_ROOT, docx_path)
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
                st.warning(f"PDF 변환 실패 (DOCX만 다운로드 가능): {e}")
        status.update(label="✅ 생성 완료", state="complete")

    docx_bytes = docx_path.read_bytes()
    return document.document_id, docx_bytes, pdf_bytes


def _generate_membership_quote(state: dict, soffice_available: bool,
                               issuer_contact: dict | None = None):
    result = _build_membership_artifacts(
        state, soffice_available, issuer_contact,
        status_label="멤버십 견적서 생성 중...",
    )
    if result is None:
        return
    document_id, docx_bytes, pdf_bytes = result

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📝 DOCX 다운로드",
            data=docx_bytes,
            file_name=f"{document_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl2:
        if pdf_bytes:
            st.download_button(
                "📑 PDF 다운로드",
                data=pdf_bytes,
                file_name=f"{document_id}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.button("📑 PDF 사용 불가", disabled=True, use_container_width=True)


def _preview_membership_quote(state: dict, soffice_available: bool,
                              issuer_contact: dict | None = None):
    result = _build_membership_artifacts(
        state, soffice_available, issuer_contact,
        status_label="미리보기 생성 중...",
    )
    if result is None:
        return
    document_id, _docx_bytes, pdf_bytes = result

    if not pdf_bytes:
        st.error("❌ PDF 변환기(LibreOffice)를 사용할 수 없어 미리보기를 표시할 수 없습니다.")
        return

    st.success(f"미리보기: **{document_id}.pdf**")

    rendered = False
    try:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(pdf_bytes)
        try:
            n_pages = len(pdf)
            for i in range(n_pages):
                page = pdf[i]
                bitmap = page.render(scale=2)
                img = bitmap.to_pil()
                st.image(img, caption=f"페이지 {i + 1} / {n_pages}",
                         use_container_width=True)
            rendered = True
        finally:
            pdf.close()
    except Exception as e:
        st.warning(f"이미지 미리보기를 사용할 수 없습니다 ({e}). 아래 다운로드로 확인하세요.")

    st.download_button(
        ("📑 미리보기 PDF 다운로드" if rendered else "📑 미리보기 PDF 다운로드 (필수)"),
        data=pdf_bytes,
        file_name=f"preview_{document_id}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    st.caption("💡 미리보기는 화면 표시용입니다. 정식 파일은 '📝 멤버십 견적서 생성' 버튼을 누르세요.")


# ═════════════════════════════════════════════════════════════
# 메인 라우터
# ═════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="견적서 자동 생성", page_icon="📋", layout="wide")

    with st.sidebar:
        st.markdown("### 메뉴")
        page = st.radio(
            "페이지",
            [
                "📋 QR 견적서 작성",
                "🏢 멤버십 견적서 작성",
                "📦 카탈로그 관리",
                "⚙ 설정",
            ],
            label_visibility="collapsed",
        )

    if page == "📋 QR 견적서 작성":
        render_quote_page()
    elif page == "🏢 멤버십 견적서 작성":
        render_membership_quote_page()
    elif page == "📦 카탈로그 관리":
        render_catalog_page()
    elif page == "⚙ 설정":
        render_settings_page()


if __name__ == "__main__":
    main()
