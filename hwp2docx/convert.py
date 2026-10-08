"""한 파일을 변환하는 진입점."""

from __future__ import annotations

import os
from dataclasses import dataclass, field

from .bodytext import parse_document
from .docinfo import parse_doc_info
from .reader import HwpError, HwpFile
from .writer import write_docx


@dataclass
class ConvertResult:
    source: str
    output: str | None
    ok: bool
    warnings: list[str] = field(default_factory=list)
    error: str | None = None
    verify_message: str | None = None


def convert(path: str, output: str | None = None, safe_fonts: bool = False,
            use_layout: bool = True, font_map: dict[str, str] | None = None,
            check: bool = False) -> ConvertResult:
    """.hwp 하나를 .docx로. 실패해도 예외 대신 결과 객체를 돌려준다."""
    if output is None:
        output = os.path.splitext(path)[0] + ".docx"

    try:
        with HwpFile(path) as hwp:
            info = parse_doc_info(hwp.doc_info())
            doc = parse_document(hwp, info)
            preview = hwp.preview_image() if check else None
            page = doc.sections[0].page if doc.sections else None
            write_docx(doc, info, output, safe_fonts=safe_fonts,
                       font_map=font_map, use_layout=use_layout)
    except HwpError as exc:
        return ConvertResult(path, None, False, error=str(exc))
    except Exception as exc:  # 형식이 어긋난 파일
        return ConvertResult(path, None, False,
                             error=f"{type(exc).__name__}: {exc}")

    result = ConvertResult(path, output, True, warnings=list(doc.warnings))

    if check:
        from .verify import verify
        width_inch = (page.width / 7200) if page else 8.268
        if page and page.landscape:
            width_inch = page.height / 7200
        outcome = verify(path, output, preview, page_width_inch=width_inch)
        result.verify_message = outcome.message
    return result
