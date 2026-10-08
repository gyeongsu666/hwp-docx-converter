"""Word와 한글 사이에 두는 중간 문서 모델.

hwp2docx의 모델과 달리 서식을 ID가 아니라 '값'으로 들고 있다. Word 쪽은
스타일 상속을 풀어서 최종 값을 채우고, 한글 쪽 작성기는 같은 값끼리 묶어
서식표(header.xml) ID를 새로 매긴다. 그래서 값 객체는 전부 frozen이다
(해시 가능해야 중복 제거를 dict 한 번으로 할 수 있다).

길이는 전부 HWPUNIT(1/7200 inch), 글자 크기는 1/100 pt.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union


# ---------------------------------------------------------------------------
# 서식 값
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class CharProps:
    font_hangul: str = "맑은 고딕"
    font_latin: str = "맑은 고딕"
    size: int = 1000                  # 1/100 pt
    bold: bool = False
    italic: bool = False
    underline: str | None = None      # SOLID, DOUBLE, DOT, DASH, WAVE …
    underline_color: str = "#000000"
    strike: str | None = None         # SOLID, DOUBLE
    color: str = "#000000"
    shade: str | None = None          # 형광펜·글자 음영
    superscript: bool = False
    subscript: bool = False
    ratio: int = 100                  # 장평 %
    spacing: int = 0                  # 자간 %


@dataclass(frozen=True)
class ParaProps:
    align: str = "JUSTIFY"            # JUSTIFY LEFT RIGHT CENTER DISTRIBUTE
    left: int = 0
    right: int = 0
    indent: int = 0                   # 첫 줄. 음수면 내어쓰기
    before: int = 0
    after: int = 0
    line_type: str = "PERCENT"        # PERCENT FIXED AT_LEAST
    line_value: int = 160             # PERCENT면 %, 아니면 HWPUNIT
    keep_with_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    widow_orphan: bool = False
    #: 번호/글머리표 문단: 내어쓰기 자동 탭(한글 tabPr id=1)을 쓴다
    auto_tab: bool = False


@dataclass(frozen=True)
class BorderLine:
    type: str = "NONE"                # NONE SOLID DOT DASH DOUBLE_SLIM …
    width: str = "0.12 mm"
    color: str = "#000000"


NO_LINE = BorderLine()


@dataclass(frozen=True)
class BorderFillProps:
    left: BorderLine = NO_LINE
    right: BorderLine = NO_LINE
    top: BorderLine = NO_LINE
    bottom: BorderLine = NO_LINE
    fill: str | None = None


NO_BORDER = BorderFillProps()


# ---------------------------------------------------------------------------
# 내용
# ---------------------------------------------------------------------------

@dataclass
class Image:
    data: bytes
    ext: str                          # png jpg gif bmp wmf emf …
    width: int                        # HWPUNIT
    height: int


@dataclass
class Run:
    kind: str                         # text | tab | break | image
    props: CharProps
    text: str = ""
    image: Image | None = None


@dataclass
class Paragraph:
    props: ParaProps = field(default_factory=ParaProps)
    runs: list[Run] = field(default_factory=list)
    #: 글자가 없을 때 줄 높이를 정하는 글자 모양 (Word의 문단 기호 서식)
    mark: CharProps = field(default_factory=CharProps)
    page_break: bool = False
    column_break: bool = False

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs if r.kind == "text")


@dataclass
class Cell:
    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    border: BorderFillProps = NO_BORDER
    margins: tuple[int, int, int, int] = (510, 510, 141, 141)   # 좌 우 상 하
    valign: str = "CENTER"            # TOP CENTER BOTTOM
    blocks: list["Block"] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    col_widths: list[int] = field(default_factory=list)          # HWPUNIT
    row_heights: list[int] = field(default_factory=list)         # 최소 높이
    cells: list[Cell] = field(default_factory=list)
    align: str = "LEFT"               # 표 자체의 가로 위치
    indent: int = 0
    margins: tuple[int, int, int, int] = (510, 510, 141, 141)
    repeat_header: bool = False

    @property
    def rows(self) -> int:
        return len(self.row_heights)

    @property
    def cols(self) -> int:
        return len(self.col_widths)


Block = Union[Paragraph, Table]


@dataclass
class PageSetup:
    width: int = 59528                # A4
    height: int = 84186
    landscape: bool = False
    left: int = 8504
    right: int = 8504
    top: int = 5668
    bottom: int = 4252
    header: int = 4252
    footer: int = 4252
    gutter: int = 0


@dataclass
class Section:
    page: PageSetup = field(default_factory=PageSetup)
    blocks: list[Block] = field(default_factory=list)


@dataclass
class Document:
    sections: list[Section] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)
