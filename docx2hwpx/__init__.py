"""docx2hwpx — Word 문서(.docx)를 한글 문서(.hwpx)로 변환한다."""

from .convert import ConvertResult, convert

__version__ = "0.1.0"
__all__ = ["convert", "ConvertResult", "__version__"]
