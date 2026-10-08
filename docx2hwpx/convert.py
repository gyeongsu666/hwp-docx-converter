"""한 파일을 변환하는 진입점."""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass, field

from .reader import read_docx
from .writer import write_hwpx


@dataclass
class ConvertResult:
    source: str
    output: str | None
    ok: bool
    warnings: list[str] = field(default_factory=list)
    error: str | None = None


def convert(path: str, output: str | None = None) -> ConvertResult:
    if output is None:
        output = os.path.splitext(path)[0] + ".hwpx"
    if not zipfile.is_zipfile(path):
        return ConvertResult(path, None, False, error=(
            "Word 2007 이후 형식(.docx)이 아닙니다. 옛 .doc 파일은 Word에서 .docx로 다시 저장해 주세요."))
    try:
        doc = read_docx(path)
        write_hwpx(doc, output)
    except Exception as exc:
        return ConvertResult(path, None, False, error=f"{type(exc).__name__}: {exc}")
    return ConvertResult(path, output, True, warnings=list(doc.warnings))
