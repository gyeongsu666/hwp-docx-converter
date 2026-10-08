"""hwp2docx — 한글 문서(.hwp)를 Word 문서(.docx)로 변환한다."""

from .convert import ConvertResult, convert
from .reader import HwpError

__version__ = "0.1.2"
__all__ = ["convert", "ConvertResult", "HwpError", "__version__"]
