"""HWP와 DOCX 사이에 두는 중간 문서 모델.

파서는 이 모델만 만들고, 작성기는 이 모델만 읽는다.
그래야 HWP 쪽 바이트 오프셋 문제와 OOXML 쪽 표현 문제가 섞이지 않는다.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Run:
    """같은 글자 모양을 공유하는 텍스트 조각."""
    text: str
    char_shape_id: int = 0
    #: 텍스트 대신 탭/줄바꿈을 뜻하는 경우
    kind: str = "text"   # 'text' | 'tab' | 'break'


@dataclass
class Image:
    """문단 안에 놓이는 그림."""
    data: bytes
    ext: str
    width: int          # HWPUNIT
    height: int
    #: 글자처럼 취급이 아니면 Word에서도 대충 제자리에 두는 수밖에 없다
    inline: bool = True


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    para_shape_id: int = 0
    style_id: int = 0
    #: 글자가 하나도 없어도 줄 높이를 결정하는 글자 모양.
    #: 빈 표 칸이 많은 양식 문서에서 행 높이가 이 값에 좌우된다.
    default_char_shape_id: int = 0
    images: list[Image] = field(default_factory=list)
    #: 이 문단 자리에 끼어드는 표들 (HWP는 표를 문단에 매단다)
    tables: list["Table"] = field(default_factory=list)
    #: 글머리표/번호 매기기가 걸린 문단인지
    list_level: int | None = None
    page_break_before: bool = False

    # --- 한글이 직접 계산해 둔 배치 정보 (PARA_LINE_SEG) ---------------------
    # 한글의 줄 간격 계산 방식을 역으로 추측하는 대신, 한글이 저장해 둔
    # 실제 배치 결과를 그대로 쓴다. 세로 위치가 원본과 맞는 이유.
    #: 한 줄의 높이 (본문 높이 + 줄 간격), HWPUNIT
    line_height: int = 0
    #: 구역 본문 상단 기준 이 문단의 시작/끝 y좌표, HWPUNIT
    layout_top: int | None = None
    layout_bottom: int | None = None

    @property
    def has_layout(self) -> bool:
        return self.layout_top is not None and self.line_height > 0

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs if r.kind == "text")

    @property
    def is_empty(self) -> bool:
        return not self.runs and not self.images and not self.tables


@dataclass
class Cell:
    col: int
    row: int
    col_span: int = 1
    row_span: int = 1
    width: int = 0              # HWPUNIT
    height: int = 0
    border_fill_id: int = 0
    margins: tuple[int, int, int, int] = (0, 0, 0, 0)   # 좌 우 상 하, HWPUNIT
    vertical_align: int = 0     # 0 위, 1 가운데, 2 아래
    paragraphs: list[Paragraph] = field(default_factory=list)


@dataclass
class Table:
    rows: int = 0
    cols: int = 0
    cells: list[Cell] = field(default_factory=list)
    border_fill_id: int = 0
    cell_spacing: int = 0
    #: 표 전체 바깥 여백 (좌 우 상 하)
    outer_margins: tuple[int, int, int, int] = (0, 0, 0, 0)
    #: 표를 담은 개체의 폭. 셀 폭 합계가 0일 때 대비용.
    total_width: int = 0
    #: 개체가 들고 있는 표 전체 높이(HWPUNIT). 한글이 계산해 둔 값이라
    #: 행 높이를 여기에 맞춰 보정하면 원본과 같은 높이가 나온다.
    total_height: int = 0

    def row_heights(self) -> list[int]:
        """행별 높이(HWPUNIT). 표 전체 높이에 맞춰 비례 보정한다."""
        heights: list[int] = []
        for row in range(self.rows):
            cells = [c for c in self.cells if c.row == row and c.row_span == 1]
            if not cells:
                cells = [c for c in self.cells if c.row == row]
            declared = max((c.height for c in cells), default=0)
            if declared:
                heights.append(declared)
                continue
            # 셀 높이가 없으면 내용(문단 줄 높이 + 셀 위아래 여백)으로 구한다.
            best = 0
            for cell in cells:
                content = sum(p.line_height or 0 for p in cell.paragraphs)
                best = max(best, content + cell.margins[2] + cell.margins[3])
            heights.append(best)
        total = sum(heights)
        if self.total_height and total and abs(self.total_height - total) > 4:
            scale = self.total_height / total
            heights = [int(round(h * scale)) for h in heights]
        return heights

    def grid(self) -> list[int]:
        """열 폭 목록(HWPUNIT). 병합을 고려해 첫 행부터 채워 나간다."""
        widths: dict[int, int] = {}
        for cell in sorted(self.cells, key=lambda c: (c.col_span, c.row)):
            if cell.col_span == 1:
                widths.setdefault(cell.col, cell.width)
        if len(widths) < self.cols:
            # 병합 때문에 폭을 못 구한 열은 나머지를 고르게 나눈다.
            known = sum(widths.values())
            missing = [c for c in range(self.cols) if c not in widths]
            rest = max(self.total_width - known, 0)
            share = rest // len(missing) if missing else 0
            for c in missing:
                widths[c] = share or 1000
        return [widths.get(c, 0) for c in range(self.cols)]


@dataclass
class PageDef:
    width: int = 59528          # A4 가로, HWPUNIT
    height: int = 84188
    margin_left: int = 8504
    margin_right: int = 8504
    margin_top: int = 5668
    margin_bottom: int = 4252
    header: int = 4252
    footer: int = 4252
    gutter: int = 0
    landscape: bool = False


@dataclass
class Section:
    page: PageDef = field(default_factory=PageDef)
    paragraphs: list[Paragraph] = field(default_factory=list)


@dataclass
class Document:
    sections: list[Section] = field(default_factory=list)
    #: 변환하면서 버리거나 근사한 것들 — 사용자에게 알려 줄 목록
    warnings: list[str] = field(default_factory=list)

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)
