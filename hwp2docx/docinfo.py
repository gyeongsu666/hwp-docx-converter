"""DocInfo 스트림 파싱: 문서 전체가 공유하는 서식 테이블.

본문 레코드는 서식을 값이 아니라 ID로만 갖고 있다. 여기서 만든
표(글꼴, 글자 모양, 문단 모양, 테두리/배경)를 본문이 ID로 찾아 쓴다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import tags as T
from .reader import Record, colorref, iter_records, read_wstring

# ---------------------------------------------------------------------------
# 값 객체
# ---------------------------------------------------------------------------

ALIGN_JUSTIFY, ALIGN_LEFT, ALIGN_RIGHT, ALIGN_CENTER, ALIGN_DISTRIBUTE, ALIGN_DIVIDE = range(6)

#: HWP 테두리 종류 -> OOXML w:val
BORDER_STYLE = {
    0: None,            # 없음
    1: "single",        # 실선
    2: "dashed",
    3: "dotted",
    4: "dashSmallGap",
    5: "dotDash",
    6: "dotDotDash",
    7: "single",        # 긴 점선
    8: "single",        # 원형 점선
    9: "single",        # 2중 원형 점선
    10: "double",
    11: "double",
    12: "triple",
    13: "thinThickSmallGap",
    14: "thickThinSmallGap",
    15: "thinThickThinSmallGap",
    16: "wave",
    17: "doubleWave",
    18: "thick",
}

#: HWP 테두리 두께 코드 -> mm
BORDER_WIDTH_MM = {
    0: 0.1, 1: 0.12, 2: 0.15, 3: 0.2, 4: 0.25, 5: 0.3, 6: 0.4,
    7: 0.5, 8: 0.6, 9: 0.7, 10: 1.0, 11: 1.5, 12: 2.0,
}


def _mm_to_eighth_pt(mm: float) -> int:
    """mm -> 1/8 pt. OOXML의 w:sz 단위. Word가 받는 범위로 자른다."""
    return max(2, min(96, int(round(mm / 25.4 * 72 * 8))))


@dataclass(frozen=True)
class Border:
    style: str | None
    size: int          # 1/8 pt
    color: str | None  # 'RRGGBB'

    @property
    def visible(self) -> bool:
        return self.style is not None


NO_BORDER = Border(None, 0, None)


@dataclass
class BorderFill:
    left: Border = NO_BORDER
    right: Border = NO_BORDER
    top: Border = NO_BORDER
    bottom: Border = NO_BORDER
    diagonal: Border = NO_BORDER
    fill_color: str | None = None


@dataclass
class FaceName:
    name: str
    substitute: str | None = None


@dataclass
class CharShape:
    face_ids: tuple[int, ...] = (0,) * 7
    base_size: int = 1000      # 1/100 pt
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    superscript: bool = False
    subscript: bool = False
    color: str | None = None
    underline_color: str | None = None
    shade_color: str | None = None
    #: 장평(%) — 100이 기본
    ratio: int = 100
    #: 자간(%)
    spacing: int = 0
    #: 상대 크기(%)
    rel_size: int = 100

    @property
    def half_points(self) -> int:
        """OOXML w:sz 단위(1/2 pt). 상대 크기까지 반영한다."""
        pt = self.base_size / 100.0 * (self.rel_size / 100.0)
        return max(2, int(round(pt * 2)))


@dataclass
class ParaShape:
    align: int = ALIGN_JUSTIFY
    margin_left: int = 0       # HWPUNIT
    margin_right: int = 0
    indent: int = 0            # 음수면 내어쓰기
    space_before: int = 0
    space_after: int = 0
    line_spacing: int = 160
    line_spacing_type: int = 0  # 0 = %, 1 = 고정값, 2 = 여백만
    border_fill_id: int = 0
    keep_with_next: bool = False
    page_break_before: bool = False
    widow_control: bool = False


@dataclass
class BinDataItem:
    kind: str          # 'embedding' | 'storage' | 'link'
    bin_id: int = 0
    ext: str = ""
    path: str = ""


@dataclass
class Style:
    name: str = ""
    eng_name: str = ""
    para_shape_id: int = 0
    char_shape_id: int = 0
    kind: int = 0      # 0 = 문단, 1 = 글자


@dataclass
class DocInfo:
    face_names: list[FaceName] = field(default_factory=list)
    char_shapes: list[CharShape] = field(default_factory=list)
    para_shapes: list[ParaShape] = field(default_factory=list)
    border_fills: list[BorderFill] = field(default_factory=list)
    bin_items: list[BinDataItem] = field(default_factory=list)
    styles: list[Style] = field(default_factory=list)
    #: 글꼴 ID 테이블은 언어별로 나뉘어 있어서 언어별 시작 인덱스가 필요하다
    face_lang_offsets: list[int] = field(default_factory=list)

    # -- 조회 도우미 --------------------------------------------------------
    def char_shape(self, idx: int) -> CharShape:
        if 0 <= idx < len(self.char_shapes):
            return self.char_shapes[idx]
        return CharShape()

    def para_shape(self, idx: int) -> ParaShape:
        if 0 <= idx < len(self.para_shapes):
            return self.para_shapes[idx]
        return ParaShape()

    def border_fill(self, idx: int) -> BorderFill:
        """테두리/배경 ID는 1부터 시작한다."""
        i = idx - 1
        if 0 <= i < len(self.border_fills):
            return self.border_fills[i]
        return BorderFill()

    def font_name(self, face_id: int) -> str:
        if 0 <= face_id < len(self.face_names):
            fn = self.face_names[face_id]
            return fn.name or fn.substitute or ""
        return ""


# ---------------------------------------------------------------------------
# 파서
# ---------------------------------------------------------------------------

def _parse_face_name(rec: Record) -> FaceName:
    data = rec.payload
    prop = data[0]
    name, off = read_wstring(data, 1)
    substitute = None
    if prop & 0x80:  # 대체 글꼴
        off += 1     # 대체 글꼴 유형
        substitute, off = read_wstring(data, off)
    return FaceName(name=name, substitute=substitute)


def _parse_border(data: bytes, off: int) -> tuple[Border, int]:
    kind = data[off]
    width = data[off + 1]
    color = colorref(int.from_bytes(data[off + 2 : off + 6], "little"))
    style = BORDER_STYLE.get(kind)
    size = _mm_to_eighth_pt(BORDER_WIDTH_MM.get(width, 0.12))
    return Border(style, size, color), off + 6


def _parse_border_fill(rec: Record) -> BorderFill:
    data = rec.payload
    off = 2  # 속성
    left, off = _parse_border(data, off)
    right, off = _parse_border(data, off)
    top, off = _parse_border(data, off)
    bottom, off = _parse_border(data, off)
    diagonal, off = _parse_border(data, off)

    fill_color = None
    if off + 4 <= len(data):
        fill_type = int.from_bytes(data[off : off + 4], "little")
        off += 4
        if fill_type & 0x01 and off + 4 <= len(data):
            back = int.from_bytes(data[off : off + 4], "little")
            pattern_color = int.from_bytes(data[off + 4 : off + 8], "little") if off + 8 <= len(data) else None
            pattern_type = int.from_bytes(data[off + 8 : off + 12], "little", signed=True) if off + 12 <= len(data) else -1
            fill_color = colorref(back)
            # 무늬가 있으면 무늬 색이 실제로 더 눈에 띄는 경우가 있으나,
            # Word 쪽 표현이 달라 배경색만 살린다.
            if fill_color is None and pattern_type is not None and pattern_type >= 0:
                fill_color = colorref(pattern_color)
    return BorderFill(left, right, top, bottom, diagonal, fill_color)


def _parse_char_shape(rec: Record) -> CharShape:
    face_ids = tuple(rec.u16(i * 2) for i in range(7))
    ratio = rec.u8(14)
    spacing = rec.i8(21)
    rel_size = rec.u8(28)
    base_size = rec.i32(42)
    prop = rec.u32(46)

    underline_kind = (prop >> 2) & 0x03
    strike_kind = (prop >> 18) & 0x07

    cs = CharShape(
        face_ids=face_ids,
        base_size=base_size,
        italic=bool(prop & 0x01),
        bold=bool((prop >> 1) & 0x01),
        underline=underline_kind == 1,
        strike=strike_kind != 0,
        superscript=bool((prop >> 15) & 0x01),
        subscript=bool((prop >> 16) & 0x01),
        ratio=ratio or 100,
        spacing=spacing,
        rel_size=rel_size or 100,
    )
    if rec.has(52, 4):
        cs.color = colorref(rec.u32(52))
    if rec.has(56, 4):
        cs.underline_color = colorref(rec.u32(56))
    if rec.has(60, 4):
        cs.shade_color = colorref(rec.u32(60))
    return cs


#: 바이너리 .hwp의 문단 모양은 여백·들여쓰기·문단 간격을 실제 HWPUNIT의
#: 2배로 저장한다. 한글이 만든 .hwpx를 보면 같은 값이 <hp:case>(실제 HWPUNIT)와
#: <hp:default>(2배) 두 벌로 들어 있고, .hwp의 값은 default 쪽과 일치한다.
#: 예) 개요 1 왼쪽 여백 10pt: case=1000, default=2000, .hwp=2000
PARA_SHAPE_LENGTH_SCALE = 2


def _halve(value: int) -> int:
    return int(round(value / PARA_SHAPE_LENGTH_SCALE))


def _parse_para_shape(rec: Record) -> ParaShape:
    p1 = rec.u32(0)
    ps = ParaShape(
        align=(p1 >> 2) & 0x07,
        margin_left=_halve(rec.i32_or(4)),
        margin_right=_halve(rec.i32_or(8)),
        indent=_halve(rec.i32_or(12)),
        space_before=_halve(rec.i32_or(16)),
        space_after=_halve(rec.i32_or(20)),
        line_spacing=rec.i32_or(24, 160),
        line_spacing_type=p1 & 0x03,
        border_fill_id=rec.u16_or(32),
        widow_control=bool((p1 >> 16) & 0x01),
        keep_with_next=bool((p1 >> 17) & 0x01),
        page_break_before=bool((p1 >> 19) & 0x01),
    )
    # 5.0.2.5 이후는 줄 간격이 뒤쪽으로 옮겨졌다.
    if rec.has(50, 4):
        later = rec.u32(50)
        if later:
            ps.line_spacing = later
        if rec.has(46, 4):
            ps.line_spacing_type = rec.u32(46) & 0x03
    return ps


def _parse_bin_data(rec: Record) -> BinDataItem:
    data = rec.payload
    prop = int.from_bytes(data[0:2], "little")
    kind_bits = prop & 0x0F
    off = 2
    if kind_bits == 0:  # LINK
        abs_path, off = read_wstring(data, off)
        rel_path, off = read_wstring(data, off)
        return BinDataItem(kind="link", path=abs_path or rel_path)
    bin_id = int.from_bytes(data[off : off + 2], "little")
    off += 2
    ext = ""
    if kind_bits == 1:  # EMBEDDING
        ext, off = read_wstring(data, off)
        return BinDataItem(kind="embedding", bin_id=bin_id, ext=ext)
    return BinDataItem(kind="storage", bin_id=bin_id, ext=ext)


def _parse_style(rec: Record) -> Style:
    data = rec.payload
    name, off = read_wstring(data, 0)
    eng_name, off = read_wstring(data, off)
    kind = data[off] & 0x07 if off < len(data) else 0
    off += 1
    off += 1  # 다음 문단 스타일 — Word에 대응이 없어 쓰지 않는다
    off += 2  # 언어 ID
    para_id = int.from_bytes(data[off : off + 2], "little") if off + 2 <= len(data) else 0
    off += 2
    char_id = int.from_bytes(data[off : off + 2], "little") if off + 2 <= len(data) else 0
    return Style(name=name, eng_name=eng_name, para_shape_id=para_id, char_shape_id=char_id, kind=kind)


def parse_doc_info(data: bytes) -> DocInfo:
    info = DocInfo()
    for rec in iter_records(data):
        try:
            if rec.tag == T.FACE_NAME:
                info.face_names.append(_parse_face_name(rec))
            elif rec.tag == T.BORDER_FILL:
                info.border_fills.append(_parse_border_fill(rec))
            elif rec.tag == T.CHAR_SHAPE:
                info.char_shapes.append(_parse_char_shape(rec))
            elif rec.tag == T.PARA_SHAPE:
                info.para_shapes.append(_parse_para_shape(rec))
            elif rec.tag == T.BIN_DATA:
                info.bin_items.append(_parse_bin_data(rec))
            elif rec.tag == T.STYLE:
                info.styles.append(_parse_style(rec))
            elif rec.tag == T.ID_MAPPINGS:
                # 글꼴은 언어(한글/영문/한자/일어/기타/기호/사용자)별로 나뉜다.
                counts = [rec.u32(i * 4) for i in range(min(18, len(rec.payload) // 4))]
                # 인덱스 1~7이 언어별 글꼴 개수
                offsets, acc = [], 0
                for n in counts[1:8]:
                    offsets.append(acc)
                    acc += n
                info.face_lang_offsets = offsets
        except Exception:
            # 레코드 하나가 깨져도 문서 전체를 포기하지는 않는다.
            continue
    return info
