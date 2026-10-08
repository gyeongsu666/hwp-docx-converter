"""변환 결과를 원본과 대조한다.

.hwp 안에는 한글이 렌더링해 둔 첫 쪽 미리보기 그림(PrvImage)이 들어 있다.
변환한 .docx를 PDF로 뽑아 이미지로 만든 뒤, 두 그림에서 '글자나 선이 있는
가로 띠'를 찾아 세로 위치를 비교한다. 눈으로 보지 않고도 밀림·눌림을
수치로 확인할 수 있다.

LibreOffice(soffice)와 pdftoppm, Pillow, numpy가 있어야 동작한다.
없으면 조용히 건너뛴다.
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass

A4_WIDTH_INCH = 8.268


@dataclass
class Band:
    top: float
    bottom: float


@dataclass
class VerifyResult:
    ok: bool
    message: str
    max_error: float = 0.0
    mean_error: float = 0.0
    original_bands: int = 0
    converted_bands: int = 0


def _tool_missing() -> str | None:
    for tool in ("soffice", "pdftoppm"):
        if shutil.which(tool) is None:
            return tool
    try:
        import importlib
        importlib.import_module("numpy")
        importlib.import_module("PIL.Image")
    except ImportError:
        return "Pillow/numpy"
    return None


def _bands(path: str, page_width_inch: float, gap_inch: float = 0.02,
           threshold: int = 170) -> list[Band]:
    import numpy as np
    from PIL import Image

    image = Image.open(path).convert("L")
    array = np.array(image)
    ppi = image.size[0] / page_width_inch
    dark = (array < threshold).sum(axis=1)

    bands: list[Band] = []
    start = None
    last = None
    for y, value in enumerate(dark):
        if value > 0:
            if start is None:
                start = y
            last = y
        elif start is not None and (y - last) / ppi > gap_inch:
            bands.append(Band(start / ppi, last / ppi))
            start = None
    if start is not None:
        bands.append(Band(start / ppi, last / ppi))
    return bands


def _render_first_page(docx_path: str, out_dir: str) -> str | None:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=False, capture_output=True, timeout=180,
    )
    pdfs = glob.glob(os.path.join(out_dir, "*.pdf"))
    if not pdfs:
        return None
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", "100", "-f", "1", "-l", "1", pdfs[0],
         os.path.join(out_dir, "page")],
        check=False, capture_output=True, timeout=120,
    )
    pages = sorted(glob.glob(os.path.join(out_dir, "page*.jpg")))
    return pages[0] if pages else None


def verify(hwp_path: str, docx_path: str, preview_png: bytes | None,
           page_width_inch: float = A4_WIDTH_INCH) -> VerifyResult:
    missing = _tool_missing()
    if missing:
        return VerifyResult(False, f"검증을 건너뜁니다 ({missing}이(가) 없습니다).")
    if not preview_png:
        return VerifyResult(False, "원본에 미리보기 그림이 없어 검증할 수 없습니다.")

    with tempfile.TemporaryDirectory() as tmp:
        original_path = os.path.join(tmp, "original.png")
        with open(original_path, "wb") as fh:
            fh.write(preview_png)
        rendered = _render_first_page(docx_path, tmp)
        if rendered is None:
            return VerifyResult(False, "변환 결과를 렌더링하지 못해 검증할 수 없습니다.")

        original = _bands(original_path, page_width_inch)
        converted = _bands(rendered, page_width_inch)

    if not original:
        return VerifyResult(False, "원본 미리보기에서 내용을 찾지 못했습니다.")

    pairs = min(len(original), len(converted))
    if pairs == 0:
        return VerifyResult(False, "변환 결과에서 내용을 찾지 못했습니다.",
                            original_bands=len(original), converted_bands=len(converted))

    errors = [abs(converted[i].top - original[i].top) for i in range(pairs)]
    max_error = max(errors)
    mean_error = sum(errors) / len(errors)
    ok = max_error < 0.08 and len(original) == len(converted)

    if len(original) != len(converted):
        message = (f"첫 쪽 구성이 다릅니다 (원본 {len(original)}덩어리, "
                   f"변환본 {len(converted)}덩어리). 직접 확인해 보세요.")
    else:
        message = (f"첫 쪽 세로 위치 오차 최대 {max_error * 25.4:.1f}mm, "
                   f"평균 {mean_error * 25.4:.1f}mm")
    return VerifyResult(ok, message, max_error, mean_error, len(original), len(converted))
