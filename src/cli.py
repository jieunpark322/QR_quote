from __future__ import annotations

import os
import subprocess
import sys
from datetime import date
from pathlib import Path

import typer

from .excel_reader import read_quote_from_excel
from .excel_template import build_template
from .loader import load_brand, load_document
from .pdf_converter import convert_docx_to_pdf, find_soffice
from .renderer import render_docx


app = typer.Typer(help="B2B 견적서/계약서 자동 생성기", no_args_is_help=True)


@app.command(help="JSON 또는 엑셀 입력 파일에서 DOCX/PDF 생성")
def render(
    input: Path = typer.Option(..., "--input", "-i",
                               help="입력 파일 (.json 또는 .xlsx)"),
    brand: str | None = typer.Option(None, "--brand", "-b",
                                     help="브랜드 ID (지정하지 않으면 입력 파일의 brand_id 사용)"),
    out: Path = typer.Option(Path("output"), "--out", "-o", help="출력 디렉터리"),
    pdf: bool = typer.Option(True, "--pdf/--no-pdf", help="PDF도 생성할지 여부"),
    project_root: Path = typer.Option(Path("."), "--project-root", help="프로젝트 루트"),
):
    project_root = project_root.resolve()
    input_path = input.resolve()
    out = out.resolve()

    typer.echo(f"📂 입력 파일: {input_path}")
    ext = input_path.suffix.lower()
    if ext == ".xlsx":
        typer.echo("📊 엑셀에서 견적 정보 읽는 중...")
        document = read_quote_from_excel(input_path, project_root)
    elif ext == ".json":
        document = load_document(input_path)
    else:
        raise typer.BadParameter(f"지원하지 않는 입력 형식입니다: {ext} (.json 또는 .xlsx)")

    brand_id = brand or document.brand_id
    typer.echo(f"🎨 브랜드 로드: {brand_id}")
    brand_obj = load_brand(project_root, brand_id)

    docx_path = out / f"{document.document_id}.docx"
    typer.echo(f"📝 DOCX 생성 중: {docx_path}")
    render_docx(brand_obj, document, project_root, docx_path)
    typer.secho("   ✓ DOCX 생성 완료", fg=typer.colors.GREEN)

    if pdf:
        if find_soffice() is None:
            typer.secho("   ⚠ LibreOffice 미설치 — PDF 변환을 건너뜁니다.",
                        fg=typer.colors.YELLOW)
        else:
            typer.echo("📑 PDF 변환 중...")
            pdf_path = convert_docx_to_pdf(docx_path, out)
            typer.secho(f"   ✓ PDF 생성 완료: {pdf_path}", fg=typer.colors.GREEN)

    typer.secho("\n완료되었습니다 ✨", fg=typer.colors.BRIGHT_GREEN, bold=True)


@app.command(help="작성용 엑셀 입력 템플릿을 새로 만듭니다")
def template(
    out: Path = typer.Option(Path("input/견적서_작성.xlsx"), "--out", "-o",
                             help="생성할 엑셀 경로"),
    brand: str = typer.Option("softment", "--brand", "-b", help="기본 브랜드 ID"),
    valid_days: int | None = typer.Option(None, "--valid-days",
                                          help="유효기간 (발행일로부터, 일수). 미지정 시 config/labels.json 기본값 사용."),
    project_root: Path = typer.Option(Path("."), "--project-root", help="프로젝트 루트"),
):
    project_root = project_root.resolve()
    out = out.resolve()
    typer.echo(f"📋 엑셀 템플릿 생성 중: {out}")
    path = build_template(project_root, out,
                          brand_id=brand,
                          issued_date=date.today(),
                          valid_days=valid_days)
    typer.secho(f"   ✓ 생성 완료: {path}", fg=typer.colors.GREEN)
    typer.echo("\n다음 단계:")
    typer.echo("  1. 위 엑셀을 열어 내용을 채우고 저장하세요.")
    typer.echo(f"  2. python -m src.cli render --input \"{path}\" --out output")


@app.command(help="웹 인터페이스 실행 (브라우저에서 폼으로 견적서 작성)")
def web(
    port: int = typer.Option(8501, "--port", help="웹 포트 (기본 8501)"),
    project_root: Path = typer.Option(Path("."), "--project-root", help="프로젝트 루트"),
):
    project_root = project_root.resolve()
    webapp_path = Path(__file__).parent / "webapp.py"
    typer.secho(
        f"🌐 웹 앱을 시작합니다.  http://localhost:{port}",
        fg=typer.colors.BRIGHT_GREEN, bold=True,
    )
    typer.echo("   브라우저가 자동으로 열립니다. 종료하려면 Ctrl+C.")
    env = os.environ.copy()
    env["CONTRACT_SYSTEM_ROOT"] = str(project_root)
    subprocess.run(
        [sys.executable, "-m", "streamlit", "run", str(webapp_path),
         "--server.port", str(port),
         "--server.headless", "false"],
        env=env,
    )


if __name__ == "__main__":
    app()
