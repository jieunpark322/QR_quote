from __future__ import annotations

import json
from pathlib import Path

import frontmatter
from jinja2 import Environment, StrictUndefined

from .models import Brand, Clause, ClauseRef, QuoteDocument


def load_brand(project_root: Path, brand_id: str) -> Brand:
    path = project_root / "brands" / brand_id / "brand.json"
    if not path.exists():
        raise FileNotFoundError(f"브랜드 파일을 찾을 수 없습니다: {path}")
    return Brand.model_validate_json(path.read_text(encoding="utf-8"))


def load_document(data_path: Path) -> QuoteDocument:
    if not data_path.exists():
        raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_path}")
    return QuoteDocument.model_validate_json(data_path.read_text(encoding="utf-8"))


def load_clause(project_root: Path, category: str, clause_id: str) -> Clause:
    candidates = [
        project_root / "clauses" / category / f"{clause_id}.md",
        project_root / "clauses" / "common" / f"{clause_id}.md",
    ]
    for path in candidates:
        if path.exists():
            post = frontmatter.load(path)
            meta = post.metadata
            return Clause(
                clause_id=meta.get("clause_id", clause_id),
                title=meta.get("title", clause_id),
                category=meta.get("category", category),
                required=meta.get("required", False),
                variables=meta.get("variables", []),
                body_template=post.content,
            )
    raise FileNotFoundError(
        f"조항 파일을 찾을 수 없습니다: {clause_id} (category={category})"
    )


def render_clause_body(clause: Clause, clause_ref: ClauseRef, clause_number: int) -> str:
    env = Environment(undefined=StrictUndefined, autoescape=False)
    template = env.from_string(clause.body_template)
    return template.render(clause_number=clause_number, **clause_ref.vars)
