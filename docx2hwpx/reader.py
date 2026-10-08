"""DOCX를 읽어 중간 모델로 옮긴다."""

from __future__ import annotations

import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from . import units
from .model import (
    NO_LINE, BorderFillProps, BorderLine, Cell, CharProps, Document, Image,
    PageSetup, Paragraph, ParaProps, Run, Section, Table,
)
from .numbering import Numbering
from .styles import (
    W, StyleSheet, ThemeFonts, merge_ppr, merge_rpr, onoff, parse_ppr,
    parse_rpr, to_int, w, wattr,
)

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
V = "{urn:schemas-microsoft-com:vml}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

DEFAULT_FONT = "맑은 고딕"

HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF", "magenta": "#FF00FF",
    "blue": "#0000FF", "red": "#FF0000", "darkBlue": "#000080", "darkCyan": "#008080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#C0C0C0",
    "black": "#000000", "white": "#FFFFFF",
}
UNDERLINE = {
    "single": "SOLID", "words": "SOLID", "thick": "SOLID", "double": "DOUBLE_SLIM",
    "dotted": "DOT", "dottedHeavy": "DOT", "dash": "DASH", "dashedHeavy": "DASH",
    "dashLong": "LONG_DASH", "dashLongHeavy": "LONG_DASH", "dotDash": "DASH_DOT",
    "dashDotHeavy": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT", "dashDotDotHeavy": "DASH_DOT_DOT",
    "wave": "WAVE", "wavyHeavy": "WAVE", "wavyDouble": "DOUBLE_WAVE",
}
BORDER = {
    "single": "SOLID", "thick": "SOLID", "double": "DOUBLE_SLIM", "dotted": "DOT",
    "dashed": "DASH", "dashSmallGap": "DASH", "dotDash": "DASH_DOT",
    "dotDotDash": "DASH_DOT_DOT", "triple": "SLIM_THICK_SLIM",
    "thinThickSmallGap": "SLIM_THICK", "thickThinSmallGap": "THICK_SLIM",
    "thinThickMediumGap": "SLIM_THICK", "thickThinMediumGap": "THICK_SLIM",
    "thinThickThinSmallGap": "SLIM_THICK_SLIM", "wave": "WAVE", "doubleWave": "DOUBLE_WAVE",
}
ALIGN = {
    "left": "LEFT", "start": "LEFT", "center": "CENTER", "right": "RIGHT", "end": "RIGHT",
    "both": "JUSTIFY", "distribute": "DISTRIBUTE", "thaiDistribute": "DISTRIBUTE",
    "lowKashida": "JUSTIFY", "mediumKashida": "JUSTIFY", "highKashida": "JUSTIFY",
}
IMAGE_EXT = {
    "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/bmp": "bmp",
    "image/x-wmf": "wmf", "image/x-emf": "emf", "image/tiff": "tif",
}


def _hex(value: str | None) -> str | None:
    if not value or value == "auto":
        return None
    v = value.strip().lstrip("#")
    return "#" + v.upper() if len(v) == 6 else None


class DocxReader:
    def __init__(self, path: str):
        self.docx = docx.Document(path)
        self.doc = Document()
        part = self.docx.part

        theme_el = None
        for rel in part.rels.values():
            if rel.reltype == RT.THEME:
                from lxml import etree
                theme_el = etree.fromstring(rel.target_part.blob)
        styles_el = None
        for rel in part.rels.values():
            if rel.reltype == RT.STYLES:
                styles_el = rel.target_part.element
        self.styles = StyleSheet(styles_el, ThemeFonts(theme_el))
        try:
            self.numbering = Numbering(part.numbering_part.element)
        except (KeyError, NotImplementedError, AttributeError):
            self.numbering = Numbering(None)

        self.field_stack: list[str] = []      # 'instr' | 'result'
        self.footnotes = 0

    # ------------------------------------------------------------------ 입구
    def read(self) -> Document:
        body = self.docx.element.body
        current = Section()
        for kind, obj in self._walk_blocks(list(body), table_style=None):
            if kind == "block":
                current.blocks.extend(obj if isinstance(obj, list) else [obj])
            elif kind == "section":
                current.page = self._page(obj)
                self.doc.sections.append(current)
                current = Section()
            elif kind == "final":
                current.page = self._page(obj)
        self.doc.sections.append(current)
        # 끝에 빈 구역이 생기면 버린다
        self.doc.sections = [s for s in self.doc.sections if s.blocks] or [current]

        if self.footnotes:
            self.doc.warn(f"각주·미주 {self.footnotes}개는 옮기지 못했습니다. 본문의 표시 번호도 빠집니다.")
        sect = body.find(w("sectPr"))
        if sect is not None and (sect.find(w("headerReference")) is not None
                                 or sect.find(w("footerReference")) is not None):
            if self._has_header_text(sect):
                self.doc.warn("머리말·꼬리말은 옮기지 못했습니다.")
        return self.doc

    def _walk_blocks(self, elements, table_style):
        for el in elements:
            tag = el.tag
            if tag == w("p"):
                yield "block", self._paragraphs(el, table_style)
                sect = el.find(f"{W}pPr/{W}sectPr")
                if sect is not None:
                    yield "section", sect
            elif tag == w("tbl"):
                yield "block", self._table(el)
            elif tag == w("sdt"):
                content = el.find(w("sdtContent"))
                if content is not None:
                    yield from self._walk_blocks(list(content), table_style)
            elif tag == w("customXml"):
                yield from self._walk_blocks(list(el), table_style)
            elif tag == w("sectPr"):
                yield "final", el

    def _has_header_text(self, sect) -> bool:
        for ref in list(sect.findall(w("headerReference"))) + list(sect.findall(w("footerReference"))):
            rid = ref.get(R + "id")
            try:
                part = self.docx.part.related_parts[rid]
                if "".join(part.element.itertext()).strip():
                    return True
            except KeyError:
                continue
        return False

    # ---------------------------------------------------------------- 용지
    def _page(self, sect) -> PageSetup:
        page = PageSetup()
        sz = sect.find(w("pgSz"))
        if sz is not None:
            wd = units.twip(to_int(wattr(sz, "w"), 11906))
            ht = units.twip(to_int(wattr(sz, "h"), 16838))
            page.landscape = wattr(sz, "orient") == "landscape" or wd > ht
            # 한글은 용지 크기를 세로 기준으로 적고 방향만 따로 둔다
            page.width, page.height = min(wd, ht), max(wd, ht)
        mar = sect.find(w("pgMar"))
        if mar is not None:
            g = lambda k, d: units.twip(abs(to_int(wattr(mar, k), d)))
            top, bottom = g("top", 1440), g("bottom", 1440)
            header, footer = g("header", 720), g("footer", 720)
            page.left, page.right, page.gutter = g("left", 1440), g("right", 1440), g("gutter", 0)
            # Word의 위 여백은 본문까지의 거리, 머리말 값은 머리말까지의 거리.
            # 한글은 '위 여백(머리말까지) + 머리말 높이 = 본문 시작'이다.
            page.top = min(header, top)
            page.header = max(0, top - header)
            page.bottom = min(footer, bottom)
            page.footer = max(0, bottom - footer)
        return page

    # ---------------------------------------------------------------- 서식
    def _base_layers(self, style_id, table_style):
        ppr = dict(self.styles.doc_ppr)
        rpr = dict(self.styles.doc_rpr)
        if table_style:
            ppr = merge_ppr(ppr, self.styles.style_ppr(table_style))
            rpr = merge_rpr(rpr, self.styles.style_rpr(table_style))
        ppr = merge_ppr(ppr, self.styles.style_ppr(style_id))
        rpr = merge_rpr(rpr, self.styles.style_rpr(style_id))
        return ppr, rpr

    def _char_props(self, r: dict) -> CharProps:
        fonts = r.get("fonts", {})
        east = self.styles.resolve_font(fonts.get("eastAsia"))
        latin = self.styles.resolve_font(fonts.get("ascii")) or self.styles.resolve_font(fonts.get("hAnsi"))
        east = east or latin or DEFAULT_FONT
        latin = latin or east
        size = units.half_points_to_size(r.get("sz", 20))

        shade = None
        if r.get("highlight") and r["highlight"] != "none":
            shade = HIGHLIGHT.get(r["highlight"])
        elif r.get("shd"):
            shade = _hex(r["shd"][1])

        spacing = 0
        if r.get("spacing"):
            spacing = int(round(r["spacing"] / 20.0 / (size / 100.0) * 100))
            spacing = max(-50, min(50, spacing))

        strike = "SOLID" if r.get("strike") else ("DOUBLE_SLIM" if r.get("dstrike") else None)
        u = r.get("u")
        return CharProps(
            font_hangul=east, font_latin=latin, size=size,
            bold=bool(r.get("b")), italic=bool(r.get("i")),
            underline=UNDERLINE.get(u) if u and u != "none" else None,
            underline_color=_hex(r.get("u_color")) or "#000000",
            strike=strike,
            color=_hex(r.get("color")) or "#000000",
            shade=shade,
            superscript=r.get("vertAlign") == "superscript",
            subscript=r.get("vertAlign") == "subscript",
            ratio=max(50, min(200, r.get("w", 100))),
            spacing=spacing,
        )

    def _para_props(self, p: dict, font_for_ratio: CharProps, is_list: bool) -> ParaProps:
        ind = p.get("ind", {})
        left, right = units.twip(ind.get("left", 0)), units.twip(ind.get("right", 0))
        indent = 0
        if "hanging" in ind:
            # Word hanging: 첫 줄이 왼쪽 여백에서 빠져나온다.
            # 한글 내어쓰기: 첫 줄은 왼쪽 여백, 나머지가 들어간다. → 여백을 앞당긴다.
            hang = units.twip(ind["hanging"])
            left -= hang
            indent = -hang
        elif "firstLine" in ind:
            indent = units.twip(ind["firstLine"])

        sp = p.get("spacing", {})
        before = units.twip(sp.get("before", 0))
        after = units.twip(sp.get("after", 0))
        if sp.get("beforeAutospacing"):
            before = units.twip(280)
        if sp.get("afterAutospacing"):
            after = units.twip(280)

        rule = sp.get("lineRule", "auto")
        line = sp.get("line")
        if line is None:
            line_type, value = "PERCENT", units.word_multiple_to_hwp_percent(1.0, font_for_ratio.font_hangul)
        elif rule == "exact":
            line_type, value = "FIXED", units.twip(line)
        elif rule == "atLeast":
            line_type, value = "AT_LEAST", units.twip(line)
        else:
            line_type, value = "PERCENT", units.word_multiple_to_hwp_percent(line / 240.0, font_for_ratio.font_hangul)

        return ParaProps(
            align=ALIGN.get(p.get("jc"), "LEFT"),   # Word 기본은 왼쪽 정렬
            left=max(0, left), right=max(0, right), indent=indent,
            before=max(0, before), after=max(0, after),
            line_type=line_type, line_value=value,
            keep_with_next=bool(p.get("keepNext")), keep_lines=bool(p.get("keepLines")),
            page_break_before=bool(p.get("pageBreakBefore")),
            widow_orphan=p.get("widowControl", True),
            auto_tab=is_list and indent < 0,
        )

    # ---------------------------------------------------------------- 문단
    def _paragraphs(self, p_el, table_style) -> list[Paragraph]:
        ppr_el = p_el.find(w("pPr"))
        style_id = wattr(ppr_el.find(w("pStyle")) if ppr_el is not None else None, "val") \
            or self.styles.default_para
        ppr, base_rpr = self._base_layers(style_id, table_style)
        direct = parse_ppr(ppr_el)

        # 번호: 문단 스타일 뒤, 직접 서식 앞
        num = direct.get("numPr") or ppr.get("numPr")
        label = None
        if num and num[0]:
            got = self.numbering.next_label(num[0], num[1] or 0)
            if got:
                label, lvl = got
                ppr = merge_ppr(ppr, lvl.ppr)
        ppr = merge_ppr(ppr, direct)

        mark_rpr = merge_rpr(base_rpr, parse_rpr(ppr_el.find(w("rPr")) if ppr_el is not None else None))
        mark = self._char_props(mark_rpr)

        runs: list[Run] = []
        extra_blocks: list = []
        if label is not None:
            lvl_props = self._char_props(merge_rpr(mark_rpr, lvl.rpr))
            if label:
                runs.append(Run("text", lvl_props, text=label))
            if lvl.suffix == "tab":
                runs.append(Run("tab", lvl_props))
            elif lvl.suffix == "space":
                runs.append(Run("text", lvl_props, text=" "))

        self._inline(list(p_el), base_rpr, runs, extra_blocks)

        # 줄 간격 환산에 쓸 글꼴: 첫 글자 조각, 없으면 문단 기호
        first = next((r.props for r in runs if r.kind == "text" and r.text.strip()), mark)
        props = self._para_props(ppr, first, label is not None)

        # 쪽 나눔 표식(kind='page')으로 문단을 나눈다
        out: list[Paragraph] = []
        cur = Paragraph(props=props, mark=mark)
        for run in runs:
            if run.kind in ("page", "column"):
                if cur.runs:
                    out.append(cur)
                    cur = Paragraph(props=props, mark=mark)
                if run.kind == "page":
                    cur.page_break = True
                else:
                    cur.column_break = True
                continue
            cur.runs.append(run)
        out.append(cur)

        # 문단 앞뒤 간격 무시(contextualSpacing)는 이웃 문단을 봐야 해서
        # 여기선 표시만 남기고 작성 단계 전에 처리한다.
        for para in out:
            para._contextual = bool(ppr.get("contextualSpacing"))   # type: ignore[attr-defined]
            para._style = style_id                                   # type: ignore[attr-defined]
        return out + extra_blocks

    def _inline(self, elements, base_rpr, runs, extra_blocks, char_style=None):
        for el in elements:
            tag = el.tag
            if tag == w("r"):
                self._run(el, base_rpr, runs, extra_blocks)
            elif tag in (w("hyperlink"), w("smartTag"), w("customXml"), w("ins"),
                         w("moveTo"), w("fldSimple"), w("dir"), w("bdo")):
                self._inline(list(el), base_rpr, runs, extra_blocks)
            elif tag == w("sdt"):
                content = el.find(w("sdtContent"))
                if content is not None:
                    self._inline(list(content), base_rpr, runs, extra_blocks)
            elif tag in (M + "oMath", M + "oMathPara"):
                text = "".join(t.text or "" for t in el.iter(M + "t"))
                if text:
                    runs.append(Run("text", self._char_props(base_rpr), text=text))
                self.doc.warn("수식은 글자로만 옮겼습니다(서식이 빠짐).")
            # del, moveFrom, bookmark, proofErr, commentRange 등은 건너뛴다

    def _run(self, r_el, base_rpr, runs, extra_blocks):
        rpr_el = r_el.find(w("rPr"))
        rpr = base_rpr
        rstyle = wattr(rpr_el.find(w("rStyle")) if rpr_el is not None else None, "val")
        if rstyle:
            rpr = merge_rpr(rpr, self.styles.style_rpr(rstyle))
        raw = merge_rpr(rpr, parse_rpr(rpr_el))
        props = self._char_props(raw)
        hidden = raw.get("vanish")
        caps = raw.get("caps") or raw.get("smallCaps")

        for child in r_el:
            tag = child.tag
            if tag == w("fldChar"):
                kind = wattr(child, "fldCharType")
                if kind == "begin":
                    self.field_stack.append("instr")
                elif kind == "separate" and self.field_stack:
                    self.field_stack[-1] = "result"
                elif kind == "end" and self.field_stack:
                    self.field_stack.pop()
                continue
            if self.field_stack and self.field_stack[-1] == "instr":
                continue          # 필드 명령문은 보이지 않는 글자
            if hidden:
                continue
            if tag == w("t"):
                text = child.text or ""
                if caps:
                    text = text.upper()
                if text:
                    if runs and runs[-1].kind == "text" and runs[-1].props == props:
                        runs[-1].text += text
                    else:
                        runs.append(Run("text", props, text=text))
            elif tag in (w("tab"), w("ptab")):
                runs.append(Run("tab", props))
            elif tag == w("br"):
                kind = wattr(child, "type")
                runs.append(Run("page" if kind == "page" else "column" if kind == "column" else "break", props))
            elif tag == w("cr"):
                runs.append(Run("break", props))
            elif tag == w("noBreakHyphen"):
                runs.append(Run("text", props, text="-"))
            elif tag == w("sym"):
                code = to_int(int(wattr(child, "char", "0"), 16), 0) if wattr(child, "char") else 0
                if code:
                    from .numbering import BULLET_MAP
                    ch = chr(code if code < 0xF000 else code)
                    runs.append(Run("text", props, text=BULLET_MAP.get(ch, ch)))
            elif tag in (w("footnoteReference"), w("endnoteReference")):
                self.footnotes += 1
            elif tag == w("drawing"):
                self._drawing(child, props, runs, extra_blocks)
            elif tag in (w("pict"), w("object")):
                self._vml(child, props, runs, extra_blocks)
            elif tag == MC + "AlternateContent":
                choice = child.find(MC + "Choice")
                target = choice if choice is not None else child.find(MC + "Fallback")
                if target is not None:
                    for sub in target:
                        if sub.tag == w("drawing"):
                            self._drawing(sub, props, runs, extra_blocks)
                        elif sub.tag in (w("pict"), w("object")):
                            self._vml(sub, props, runs, extra_blocks)

    # ---------------------------------------------------------------- 그림
    def _image_part(self, rid):
        try:
            part = self.docx.part.related_parts[rid]
        except KeyError:
            return None, None
        ext = IMAGE_EXT.get(part.content_type)
        if ext is None:
            name = str(part.partname).rsplit(".", 1)[-1].lower()
            ext = {"jpeg": "jpg"}.get(name, name) if name in ("png", "jpg", "jpeg", "gif", "bmp", "wmf", "emf") else None
        return part.blob, ext

    def _drawing(self, el, props, runs, extra_blocks):
        container = el.find(WP + "inline")
        if container is None:
            container = el.find(WP + "anchor")
            if container is not None:
                self.doc.warn("떠 있는 그림·도형은 글자처럼 취급해 문단 안에 넣었습니다. 위치가 달라질 수 있습니다.")
        if container is None:
            return
        ext_el = container.find(WP + "extent")
        cx = to_int(ext_el.get("cx"), 0) if ext_el is not None else 0
        cy = to_int(ext_el.get("cy"), 0) if ext_el is not None else 0

        blip = container.find(f".//{PIC}pic//{A}blip")
        if blip is not None:
            data, ext = self._image_part(blip.get(R + "embed"))
            if data and ext:
                runs.append(Run("image", props, image=Image(data, ext, units.emu(cx), units.emu(cy))))
            else:
                self.doc.warn("형식을 알 수 없는 그림 하나를 건너뛰었습니다.")
            return
        txbx = container.find(f".//{WPS}txbx/{W}txbxContent")
        if txbx is not None:
            self._textbox(txbx, extra_blocks)
            return
        if container.find(f".//{{http://schemas.openxmlformats.org/drawingml/2006/chart}}chart") is not None:
            self.doc.warn("차트는 옮기지 못했습니다.")
        else:
            self.doc.warn("그리기 개체(도형·SmartArt 등)는 옮기지 못했습니다.")

    def _vml(self, el, props, runs, extra_blocks):
        img = el.find(f".//{V}imagedata")
        if img is not None:
            data, ext = self._image_part(img.get(R + "id"))
            if data and ext:
                shape = el.find(f".//{V}shape")
                style = shape.get("style", "") if shape is not None else ""
                dims = dict(kv.split(":", 1) for kv in style.split(";") if ":" in kv)

                def pt(v):
                    v = (v or "").strip()
                    try:
                        if v.endswith("pt"):
                            return int(float(v[:-2]) * 100)
                        if v.endswith("in"):
                            return int(float(v[:-2]) * 7200)
                    except ValueError:
                        pass
                    return 0
                runs.append(Run("image", props, image=Image(data, ext, pt(dims.get("width")), pt(dims.get("height")))))
                if el.tag == w("object"):
                    self.doc.warn("삽입 개체(OLE, 옛 수식 등)는 그림으로만 옮겼습니다.")
                return
        txbx = el.find(f".//{W}txbxContent")
        if txbx is not None:
            self._textbox(txbx, extra_blocks)

    def _textbox(self, txbx, extra_blocks):
        self.doc.warn("글상자는 본문 문단으로 풀어서 옮겼습니다. 위치가 달라집니다.")
        for kind, obj in self._walk_blocks(list(txbx), None):
            if kind == "block":
                extra_blocks.extend(obj if isinstance(obj, list) else [obj])

    # ------------------------------------------------------------------ 표
    def _borders(self, el) -> dict:
        out = {}
        if el is None:
            return out
        for side, names in (("top", ("top",)), ("bottom", ("bottom",)), ("left", ("left", "start")),
                            ("right", ("right", "end")), ("insideH", ("insideH",)),
                            ("insideV", ("insideV",))):
            for n in names:
                b = el.find(w(n))
                if b is not None:
                    val = wattr(b, "val", "nil")
                    if val in ("nil", "none"):
                        out[side] = NO_LINE
                    else:
                        out[side] = BorderLine(
                            BORDER.get(val, "SOLID"),
                            units.border_width(to_int(wattr(b, "sz"), 4)),
                            _hex(wattr(b, "color")) or "#000000",
                        )
                    break
        return out

    def _margins(self, el, base):
        if el is None:
            return base
        left, right, top, bottom = base
        for side in ("top", "bottom", "left", "start", "right", "end"):
            node = el.find(w(side))
            if node is None:
                continue
            v = units.twip(to_int(wattr(node, "w"), 0))
            if side == "top":
                top = v
            elif side == "bottom":
                bottom = v
            elif side in ("left", "start"):
                left = v
            else:
                right = v
        return (left, right, top, bottom)

    @staticmethod
    def _shade(shd) -> str | None:
        if shd is None:
            return None
        fill = _hex(wattr(shd, "fill"))
        if fill:
            return fill
        val, color = wattr(shd, "val", ""), _hex(wattr(shd, "color"))
        if val.startswith("pct") and color:
            pct = to_int(val[3:], 0) / 100.0
            rgb = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
            mix = [int(round(255 - (255 - c) * pct)) for c in rgb]
            return "#%02X%02X%02X" % tuple(mix)
        if val == "solid" and color:
            return color
        return None

    def _table(self, tbl) -> Table:
        tbl_pr = tbl.find(w("tblPr"))
        style_id = wattr(tbl_pr.find(w("tblStyle")) if tbl_pr is not None else None, "val") \
            or self.styles.default_table

        borders: dict = {}
        cell_margin = (units.twip(108), units.twip(108), 0, 0)
        align, indent, tbl_shade = "LEFT", 0, None
        for pr in self.styles.style_element(style_id, "tblPr") + ([tbl_pr] if tbl_pr is not None else []):
            borders.update(self._borders(pr.find(w("tblBorders"))))
            cell_margin = self._margins(pr.find(w("tblCellMar")), cell_margin)
            jc = wattr(pr.find(w("jc")), "val")
            if jc:
                align = ALIGN.get(jc, "LEFT")
            ind = pr.find(w("tblInd"))
            if ind is not None:
                indent = units.twip(to_int(wattr(ind, "w"), 0))
            tbl_shade = self._shade(pr.find(w("shd"))) or tbl_shade
        style_tcpr = self.styles.style_element(style_id, "tcPr")

        table = Table(align=align, indent=indent, margins=cell_margin)
        table.col_widths = [units.twip(to_int(wattr(g, "w"), 0))
                            for g in tbl.findall(f"{W}tblGrid/{W}gridCol")]

        rows = tbl.findall(w("tr"))
        anchors: dict[int, Cell] = {}          # 열 → 세로 병합 중인 셀
        for r_index, tr in enumerate(rows):
            tr_pr = tr.find(w("trPr"))
            height = 0
            if tr_pr is not None:
                h = tr_pr.find(w("trHeight"))
                if h is not None:
                    height = units.twip(to_int(wattr(h, "val"), 0))
                if onoff(tr_pr.find(w("tblHeader"))) and r_index == len(table.row_heights):
                    table.repeat_header = True
            table.row_heights.append(height)
            col = to_int(wattr(tr_pr.find(w("gridBefore")) if tr_pr is not None else None, "val"), 0)

            for tc in tr.findall(w("tc")):
                tc_pr = tc.find(w("tcPr"))
                span = max(1, to_int(wattr(tc_pr.find(w("gridSpan")) if tc_pr is not None else None, "val"), 1))
                vmerge = tc_pr.find(w("vMerge")) if tc_pr is not None else None
                if vmerge is not None and wattr(vmerge, "val", "continue") == "continue" and col in anchors:
                    anchors[col].row_span += 1
                    col += span
                    continue

                cell = Cell(row=r_index, col=col, col_span=span, margins=cell_margin)
                own_borders: dict = {}
                shade = tbl_shade
                for pr in style_tcpr + ([tc_pr] if tc_pr is not None else []):
                    own_borders.update(self._borders(pr.find(w("tcBorders"))))
                    shade = self._shade(pr.find(w("shd"))) or shade
                    cell.margins = self._margins(pr.find(w("tcMar")), cell.margins)
                    va = wattr(pr.find(w("vAlign")), "val")
                    if va:
                        cell.valign = {"top": "TOP", "center": "CENTER", "bottom": "BOTTOM"}.get(va, "TOP")
                if tc_pr is None or tc_pr.find(w("vAlign")) is None:
                    cell.valign = "TOP"      # Word 기본은 위쪽 정렬
                cell._own_borders = own_borders   # type: ignore[attr-defined]
                cell._shade = shade               # type: ignore[attr-defined]
                cell.is_header = table.repeat_header and r_index == 0

                for kind, obj in self._walk_blocks(list(tc), style_id):
                    if kind == "block":
                        cell.blocks.extend(obj if isinstance(obj, list) else [obj])
                if not cell.blocks:
                    cell.blocks.append(Paragraph())

                table.cells.append(cell)
                for c in range(col, col + span):
                    if vmerge is not None:
                        anchors[c] = cell
                    else:
                        anchors.pop(c, None)
                col += span

        # 격자 정보가 없는 표: 첫 행의 셀 폭으로 만든다
        n_cols = max((c.col + c.col_span for c in table.cells), default=0)
        if len(table.col_widths) < n_cols:
            table.col_widths += [units.twip(1440)] * (n_cols - len(table.col_widths))

        # 셀마다 네 변의 테두리를 확정한다 (셀 지정 > 표 바깥/안쪽 선)
        n_rows = len(table.row_heights)
        for cell in table.cells:
            own = cell._own_borders            # type: ignore[attr-defined]
            last_row = cell.row + cell.row_span >= n_rows
            last_col = cell.col + cell.col_span >= n_cols
            pick = lambda side, edge, inside: own.get(side, borders.get(edge if cond[side] else inside, NO_LINE))
            cond = {"top": cell.row == 0, "bottom": last_row, "left": cell.col == 0, "right": last_col}
            cell.border = BorderFillProps(
                left=pick("left", "left", "insideV"),
                right=pick("right", "right", "insideV"),
                top=pick("top", "top", "insideH"),
                bottom=pick("bottom", "bottom", "insideH"),
                fill=cell._shade,              # type: ignore[attr-defined]
            )
        return table


def read_docx(path: str) -> Document:
    return DocxReader(path).read()
