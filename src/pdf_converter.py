from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice() -> str | None:
    on_path = shutil.which("soffice")
    if on_path:
        return on_path
    for candidate in _SOFFICE_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return None


def convert_docx_to_pdf(docx_path: Path, output_dir: Path | None = None) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice(soffice)를 찾을 수 없습니다. PDF 변환을 위해 LibreOffice 설치가 필요합니다."
        )
    out_dir = output_dir or docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(docx_path)],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"PDF 변환 실패 (exit={result.returncode}):\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}"
        )

    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"PDF 파일이 생성되지 않았습니다: {pdf_path}")
    return pdf_path
