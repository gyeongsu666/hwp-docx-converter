"""BodyText 스트림 파싱: 문단, 표, 그림을 문서 모델로 옮긴다.

HWP의 구조에서 까다로운 지점 두 가지:

1. 표는 독립된 블록이 아니라 문단 안의 제어 문자(11번)에 매달린다.
   같은 문단에 제어 문자가 여러 개면 표도 여러 개다.
2. 중첩(표 안의 셀 안의 문단 안의 또 다른 표)은 오직 레코드 level로만
   표현된다. 그래서 평평한 목록이 아니라 트리로 읽어야 한다.
"""

from __future__ import annotations

import struct

from . import tags as T
from .docinfo import DocInfo
from .model import Cell, Document, Image, PageDef, Paragraph, Run, Section, Table
from .reader import HwpFile, Record, iter_records, parse_tree
from .text import clean as clean_text

#: 무시해도 되는(경고할 필요 없는) 제어 문자
_SILENT_CONTROLS = {
    T.CTRL_SECTION_DEF, T.CTRL_COLUMN_DEF, T.CTRL_BOOKMARK,
    T.CTRL_INDEX_MARK, T.CTRL_PAGE_NUM_POS, T.CTRL_AUTO_NUM,
    T.CTRL_NEW_NUM, T.CTRL_PAGE_HIDE, T.CTRL_PAGE_ODD_EVEN,
}

_CTRL_LABELS = {
    T.CTRL_FOOTNOTE: "각주",
    T.CTRL_ENDNOTE: "미주",
    T.CTRL_EQUATION: "수식",
    T.CTRL_HEADER_AREA: "머리말",
    T.CTRL_FOOTER_AREA: "꼬리말",
    T.CTRL_HIDDEN_COMMENT: "숨은 설명",
    T.CTRL_DUTMAL: "덧말",
}


class BodyTextParser:
    def __init__(self, hwp: HwpFile, info: DocInfo, doc: Document):
        self.hwp = hwp
        self.info = info
        self.doc = doc

    # -- 공개 --------------------------------------------------------------
    def parse_section(self, data: bytes) -> Section:
        section = Section()
        roots = parse_tree(list(iter_records(data)))
        section.paragraphs = self._parse_paragraph_list(roots, section)
        return section

    # -- 문단 --------------------------------------------------------------
    def _parse_paragraph_list(self, records: list[Record], section: Section | None) -> list[Paragraph]:
        """PARA_HEADER로 시작하는 레코드들을 문단 목록으로."""
        paragraphs: list[Paragraph] = []
        for rec in records:
            if rec.tag == T.PARA_HEADER:
                paragraphs.append(self._parse_paragraph(rec, section))
        return paragraphs

    @staticmethod
    def _group_lists(children: list[Record]) -> list[tuple[Record, list[Record]]]:
        """LIST_HEADER와 그 뒤에 이어지는 PARA_HEADER들을 묶는다.

        HWP는 셀의 문단을 LIST_HEADER의 '자식'이 아니라 같은 level의
        '형제'로 늘어놓는다. 그래서 트리만 보면 셀이 비어 보인다.
        """
        groups: list[tuple[Record, list[Record]]] = []
        current: list[Record] | None = None
        for child in children:
            if child.tag == T.LIST_HEADER:
                current = []
                groups.append((child, current))
            elif child.tag == T.PARA_HEADER and current is not None:
                current.append(child)
        return groups

    def _parse_paragraph(self, rec: Record, section: Section | None) -> Paragraph:
        para = Paragraph(
            para_shape_id=rec.u16_or(8),
            style_id=rec.u8(10) if rec.has(10) else 0,
        )

        text_rec = None
        shape_rec = None
        seg_rec = None
        ctrl_recs: list[Record] = []
        for child in rec.children:
            if child.tag == T.PARA_TEXT:
                text_rec = child
            elif child.tag == T.PARA_CHAR_SHAPE:
                shape_rec = child
            elif child.tag == T.PARA_LINE_SEG:
                seg_rec = child
            elif child.tag == T.CTRL_HEADER:
                ctrl_recs.append(child)

        self._apply_line_segments(para, seg_rec)
        controls = self._parse_controls(ctrl_recs, section, para)
        self._build_runs(para, text_rec, shape_rec, controls)
        return para

    @staticmethod
    def _apply_line_segments(para: Paragraph, rec: Record | None) -> None:
        """PARA_LINE_SEG: 한글이 계산해 둔 줄 배치 결과.

        한 항목은 36바이트:
            UINT32 텍스트 시작 위치
            INT32  줄의 세로 위치        (구역 본문 상단 기준)
            INT32  줄의 높이
            INT32  텍스트 부분의 높이
            INT32  베이스라인까지 거리
            INT32  줄 간격
            INT32  컬럼에서의 시작 위치
            INT32  세그먼트 폭
            UINT32 태그
        """
        if rec is None or len(rec.payload) < 36:
            return
        count = len(rec.payload) // 36
        segments = [struct.unpack_from("<IiiiiiiiI", rec.payload, i * 36) for i in range(count)]
        first, last = segments[0], segments[-1]
        # 줄 상자 = 줄 높이 + 줄 간격
        para.line_height = max(0, first[2] + first[5])
        para.layout_top = first[1]
        para.layout_bottom = last[1] + last[2] + last[5]

    def _build_runs(self, para, text_rec, shape_rec, controls) -> None:
        """PARA_TEXT를 글자 모양 경계에 맞춰 잘라 Run 목록으로 만든다.

        controls는 제어 문자가 나온 순서대로 꺼내 쓴다.
        """
        # 글자 모양 경계: [(시작 위치(문자 단위), char_shape_id), ...]
        boundaries: list[tuple[int, int]] = []
        if shape_rec is not None:
            n = len(shape_rec.payload) // 8
            boundaries = [struct.unpack_from("<II", shape_rec.payload, i * 8) for i in range(n)]
        if not boundaries:
            boundaries = [(0, 0)]
        # 글자가 없어도 줄 높이는 이 글자 모양이 정한다.
        para.default_char_shape_id = boundaries[0][1]

        if text_rec is None:
            return

        def shape_at(char_pos: int) -> int:
            current = boundaries[0][1]
            for start, shape_id in boundaries:
                if start <= char_pos:
                    current = shape_id
                else:
                    break
            return current

        data = text_rec.payload
        i = 0            # 바이트 위치
        char_pos = 0     # 문자 위치 (글자 모양 경계는 이 단위)
        ctrl_index = 0
        buf: list[str] = []
        buf_shape = shape_at(0)

        def flush() -> None:
            nonlocal buf
            if buf:
                cleaned, leftover = clean_text("".join(buf))
                if leftover:
                    self.doc.warn(
                        "한컴 기호 글꼴의 일부 문자를 표준 문자로 바꾸지 못했습니다. "
                        "Word에서 네모로 보일 수 있습니다."
                    )
                para.runs.append(Run(cleaned, buf_shape))
                buf = []

        while i + 1 < len(data):
            (code,) = struct.unpack_from("<H", data, i)

            if code >= 32:
                shape = shape_at(char_pos)
                if shape != buf_shape:
                    flush()
                    buf_shape = shape
                buf.append(chr(code))
                i += 2
                char_pos += 1
                continue

            # --- 제어 문자 ---
            if code in T.CHAR_CONTROLS:
                if code == T.LINE_BREAK:
                    flush()
                    para.runs.append(Run("", shape_at(char_pos), kind="break"))
                elif code == T.NBSP or code == T.FIXED_SPACE:
                    buf.append(" ")
                elif code == T.HYPHEN:
                    buf.append("-")
                i += 2
                char_pos += 1
                continue

            # 인라인/확장 제어는 8 WCHAR(16바이트)를 차지한다.
            if code == T.TAB:
                flush()
                para.runs.append(Run("", shape_at(char_pos), kind="tab"))
            elif code in T.EXTENDED_CONTROLS:
                flush()
                if ctrl_index < len(controls):
                    obj = controls[ctrl_index]
                    ctrl_index += 1
                    if isinstance(obj, Table):
                        para.tables.append(obj)
                    elif isinstance(obj, Image):
                        para.images.append(obj)
                    elif isinstance(obj, list):  # 글상자 안의 문단들
                        for extra in obj:
                            para.runs.extend(extra.runs)
            i += 16
            char_pos += 8

        flush()

        # 제어 문자보다 개체가 더 많이 잡힌 경우(드묾) 남은 것도 붙인다.
        for obj in controls[ctrl_index:]:
            if isinstance(obj, Table):
                para.tables.append(obj)
            elif isinstance(obj, Image):
                para.images.append(obj)

    # -- 제어(표 · 그림 · 구역 정의) ----------------------------------------
    def _parse_controls(self, ctrl_recs: list[Record], section: Section | None, para: Paragraph):
        """CTRL_HEADER 목록을 표/그림/None으로 바꾼다. 순서를 유지한다."""
        result = []
        for rec in ctrl_recs:
            if len(rec.payload) < 4:
                result.append(None)
                continue
            ctrl_id = T.ctrl_id_to_str(rec.u32(0))

            if ctrl_id == T.CTRL_TABLE:
                result.append(self._parse_table(rec))
            elif ctrl_id == T.CTRL_SHAPE:
                result.append(self._parse_shape(rec))
            elif ctrl_id == T.CTRL_SECTION_DEF:
                if section is not None:
                    self._apply_section_def(rec, section)
                result.append(None)
            else:
                label = _CTRL_LABELS.get(ctrl_id)
                if label:
                    self.doc.warn(f"{label}은(는) 옮기지 못했습니다.")
                elif ctrl_id not in _SILENT_CONTROLS and ctrl_id.strip():
                    self.doc.warn(f"알 수 없는 개체 '{ctrl_id.strip()}'을(를) 건너뛰었습니다.")
                result.append(None)
        return result

    def _apply_section_def(self, rec: Record, section: Section) -> None:
        for child in rec.children:
            if child.tag == T.PAGE_DEF:
                section.page = self._parse_page_def(child)
                return

    @staticmethod
    def _parse_page_def(rec: Record) -> PageDef:
        page = PageDef(
            width=rec.u32_or(0, 59528),
            height=rec.u32_or(4, 84188),
            margin_left=rec.u32_or(8, 8504),
            margin_right=rec.u32_or(12, 8504),
            margin_top=rec.u32_or(16, 5668),
            margin_bottom=rec.u32_or(20, 4252),
            header=rec.u32_or(24, 4252),
            footer=rec.u32_or(28, 4252),
            gutter=rec.u32_or(32, 0),
        )
        page.landscape = bool(rec.u32_or(36, 0) & 0x01)
        return page

    # -- 표 ----------------------------------------------------------------
    def _parse_table(self, ctrl: Record) -> Table | None:
        table_rec = None
        for child in ctrl.children:
            if child.tag == T.TABLE:
                table_rec = child
                break
        if table_rec is None:
            return None
        cell_groups = self._group_lists(ctrl.children)

        table = Table(
            rows=table_rec.u16(4),
            cols=table_rec.u16(6),
            cell_spacing=table_rec.i16(8),
        )
        table.outer_margins = tuple(table_rec.u16_or(10 + i * 2) for i in range(4))
        # 셀 개수 배열 뒤에 표 전체 테두리/배경 ID가 온다.
        off = 18 + table.rows * 2
        table.border_fill_id = table_rec.u16_or(off)
        # 개체 공통 속성: id(0) 속성(4) 세로오프셋(8) 가로오프셋(12) 폭(16) 높이(20)
        table.total_width = ctrl.u32_or(16, 0)
        table.total_height = ctrl.u32_or(20, 0)

        for lh, para_recs in cell_groups:
            cell = self._parse_cell(lh, para_recs)
            if cell is not None:
                table.cells.append(cell)

        if not table.cells:
            return None
        # 실제 셀에서 행/열 수를 다시 확인한다(선언값이 어긋나는 파일이 있다).
        table.rows = max(table.rows, max(c.row + c.row_span for c in table.cells))
        table.cols = max(table.cols, max(c.col + c.col_span for c in table.cells))
        if not table.total_width:
            table.total_width = sum(table.grid())
        return table

    def _parse_cell(self, rec: Record, para_recs: list[Record]) -> Cell | None:
        if len(rec.payload) < 24:
            return None
        prop = rec.u32(4)
        cell = Cell(
            col=rec.u16(8),
            row=rec.u16(10),
            col_span=max(1, rec.u16(12)),
            row_span=max(1, rec.u16(14)),
            width=rec.u32(16),
            height=rec.u32(20),
            vertical_align=(prop >> 5) & 0x03,
        )
        if rec.has(24, 8):
            cell.margins = tuple(rec.u16(24 + i * 2) for i in range(4))
        cell.border_fill_id = rec.u16_or(32)
        cell.paragraphs = self._parse_paragraph_list(para_recs, None)
        return cell

    # -- 그리기 개체 / 그림 --------------------------------------------------
    def _parse_shape(self, ctrl: Record):
        """gso 제어: 그림이면 Image, 글상자면 그 안의 문단 목록."""
        picture = self._find(ctrl, T.SHAPE_COMPONENT_PICTURE)
        component = self._find(ctrl, T.SHAPE_COMPONENT)

        # 개체 공통 속성: 폭이 16, 높이가 20 오프셋
        width = ctrl.u32_or(16, 0)
        height = ctrl.u32_or(20, 0)
        if component is not None:
            w = component.u32_or(44, 0)
            h = component.u32_or(48, 0)
            if w and h:
                width, height = w, h

        if picture is not None:
            img = self._picture_to_image(picture, width, height)
            if img is not None:
                return img
            self.doc.warn("그림 하나를 읽지 못해 건너뛰었습니다.")
            return None

        # 글상자: 안쪽 문단을 본문으로 끌어낸다.
        for _lh, para_recs in self._group_lists(ctrl.children):
            paras = self._parse_paragraph_list(para_recs, None)
            if paras:
                self.doc.warn("글상자는 본문 문단으로 풀어서 옮겼습니다. 위치가 달라질 수 있습니다.")
                return paras
        if self._find(ctrl, T.EQEDIT) is not None:
            self.doc.warn("수식은 옮기지 못했습니다. 해당 자리가 비어 있습니다.")
        return None

    def _picture_to_image(self, rec: Record, width: int, height: int) -> Image | None:
        """그림 개체 속성에서 BinItem ID를 찾아 실제 이미지 바이트를 가져온다.

        레코드 배치가 버전마다 조금씩 달라, 그럴듯한 위치를 모두 시도하고
        BinData에 실제로 존재하는 ID를 채택한다.
        """
        candidates = [68, 70, 71, 72, 66]
        for off in candidates:
            if not rec.has(off, 2):
                continue
            bin_id = rec.u16(off)
            if bin_id == 0 or bin_id > len(self.info.bin_items):
                continue
            item = self.info.bin_items[bin_id - 1]
            if item.kind != "embedding":
                continue
            data = self.hwp.bin_data(item.bin_id, item.ext)
            if data and len(data) > 16:
                return Image(data=data, ext=(item.ext or "png").lower(), width=width, height=height)
        return None

    @staticmethod
    def _find(rec: Record, tag: int) -> Record | None:
        """자손 중 해당 태그를 너비 우선으로 찾는다."""
        queue = list(rec.children)
        while queue:
            node = queue.pop(0)
            if node.tag == tag:
                return node
            queue.extend(node.children)
        return None


def parse_document(hwp: HwpFile, info: DocInfo) -> Document:
    doc = Document()
    parser = BodyTextParser(hwp, info, doc)
    for name in hwp.section_names():
        doc.sections.append(parser.parse_section(hwp.section(name)))
    return doc
