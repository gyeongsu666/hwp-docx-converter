"""중간 모델을 .hwpx로 쓴다.

한글 2024가 직접 저장한 빈 문서(Skeleton.hwpx)를 뼈대로 쓴다. 뼈대의 서식표
(글꼴·글자 모양·문단 모양·스타일)는 그대로 두고 그 뒤에 필요한 것만 덧붙이며,
본문도 뼈대의 첫 문단(구역 설정을 품은 문단)을 살려서 그 뒤를 채운다.
한글이 받아들이는 구조에서 최대한 벗어나지 않으려는 것이다.

한글이 파일을 여는 조건 (python-hwpx 검사기와 실제 사례 기준):
  - 구역 설정(secPr)은 첫 문단의 첫 run에 있어야 한다
  - 표는 sz/pos/outMargin/inMargin, 셀은 subList/cellAddr/cellSpan/cellSz/cellMargin 필수
  - 줄 배치 캐시(linesegarray)는 넣지 않는다. 넣으면 한글이 다시 계산하지 않고
    그 값을 그대로 써서 글자가 겹친다
  - mimetype은 zip의 첫 항목, 무압축
"""

from __future__ import annotations

import copy
import io
import os
import zipfile

from lxml import etree

from .model import (
    BorderFillProps, CharProps, Document, Image, Paragraph, ParaProps, Run,
    Section, Table,
)

HH = "http://www.hancom.co.kr/hwpml/2011/head"
HC = "http://www.hancom.co.kr/hwpml/2011/core"
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS = "http://www.hancom.co.kr/hwpml/2011/section"
OPF = "http://www.idpf.org/2007/opf/"
NS = {"hh": HH, "hc": HC, "hp": HP, "hs": HS, "opf": OPF}

SKELETON = os.path.join(os.path.dirname(__file__), "data", "Skeleton.hwpx")
LANGS = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")
MEDIA_TYPES = {"png": "image/png", "jpg": "image/jpg", "jpeg": "image/jpg", "gif": "image/gif",
               "bmp": "image/bmp", "wmf": "image/x-wmf", "emf": "image/x-emf", "tif": "image/tif"}


def q(ns: str, tag: str) -> str:
    return "{%s}%s" % (ns, tag)


def sub(parent, ns: str, tag: str, **attrs):
    el = etree.SubElement(parent, q(ns, tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


class HwpxWriter:
    def __init__(self, doc: Document):
        self.doc = doc
        with zipfile.ZipFile(SKELETON) as zf:
            self.parts = {name: zf.read(name) for name in zf.namelist()}
        self.header = etree.fromstring(self.parts["Contents/header.xml"])
        self.skeleton_section = etree.fromstring(self.parts["Contents/section0.xml"])
        self.ref = self.header.find("hh:refList", NS)

        self._fonts: dict[str, int] = {}
        self._char_ids: dict[CharProps, int] = {}
        self._para_ids: dict[ParaProps, int] = {}
        self._fill_ids: dict[BorderFillProps, int] = {}
        self._images: list[tuple[str, bytes, str]] = []     # (item id, data, ext)
        self._next_id = 1000000
        for f in self.ref.find("hh:fontfaces", NS)[0].findall("hh:font", NS):
            self._fonts[f.get("face")] = int(f.get("id"))

    # ------------------------------------------------------------ 서식표
    def _uid(self) -> str:
        self._next_id += 1
        return str(self._next_id)

    def _font(self, name: str) -> int:
        if name in self._fonts:
            return self._fonts[name]
        fid = len(self._fonts)
        self._fonts[name] = fid
        for face in self.ref.find("hh:fontfaces", NS).findall("hh:fontface", NS):
            font = sub(face, HH, "font", id=fid, face=name, type="TTF", isEmbedded="0")
            sub(font, HH, "typeInfo", familyType="FCAT_GOTHIC", weight="6", proportion="4",
                contrast="0", strokeVariation="1", armStyle="1", letterform="1",
                midline="1", xHeight="1")
            face.set("fontCnt", str(len(face.findall("hh:font", NS))))
        return fid

    def _container(self, tag: str):
        el = self.ref.find(f"hh:{tag}", NS)
        return el

    def border_fill(self, props: BorderFillProps) -> int:
        if props in self._fill_ids:
            return self._fill_ids[props]
        box = self._container("borderFills")
        bid = len(box.findall("hh:borderFill", NS)) + 1
        bf = sub(box, HH, "borderFill", id=bid, threeD="0", shadow="0",
                 centerLine="NONE", breakCellSeparateLine="0")
        sub(bf, HH, "slash", type="NONE", Crooked="0", isCounter="0")
        sub(bf, HH, "backSlash", type="NONE", Crooked="0", isCounter="0")
        for side in ("left", "right", "top", "bottom"):
            line = getattr(props, side)
            sub(bf, HH, f"{side}Border", type=line.type, width=line.width, color=line.color)
        sub(bf, HH, "diagonal", type="NONE", width="0.1 mm", color="#000000")
        if props.fill:
            brush = sub(bf, HC, "fillBrush")
            sub(brush, HC, "winBrush", faceColor=props.fill, hatchColor="#FF000000", alpha="0")
        box.set("itemCnt", str(bid))
        self._fill_ids[props] = bid
        return bid

    def char_pr(self, p: CharProps) -> int:
        if p in self._char_ids:
            return self._char_ids[p]
        box = self._container("charProperties")
        cid = len(box.findall("hh:charPr", NS))
        hangul, latin = self._font(p.font_hangul), self._font(p.font_latin)
        c = sub(box, HH, "charPr", id=cid, height=p.size, textColor=p.color,
                shadeColor=p.shade or "none", useFontSpace="0", useKerning="0",
                symMark="NONE", borderFillIDRef="2")
        east = {"hangul": hangul, "hanja": hangul, "japanese": hangul}
        west = {"latin": latin, "other": latin, "symbol": latin, "user": latin}
        order = ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")
        refs = {**east, **west}
        sub(c, HH, "fontRef", **{k: refs[k] for k in order})
        sub(c, HH, "ratio", **{k: p.ratio for k in order})
        sub(c, HH, "spacing", **{k: p.spacing for k in order})
        sub(c, HH, "relSz", **{k: 100 for k in order})
        sub(c, HH, "offset", **{k: 0 for k in order})
        if p.italic:
            sub(c, HH, "italic")
        if p.bold:
            sub(c, HH, "bold")
        sub(c, HH, "underline", type="BOTTOM" if p.underline else "NONE",
            shape=p.underline or "SOLID", color=p.underline_color)
        sub(c, HH, "strikeout", shape=p.strike or "NONE", color=p.color)
        sub(c, HH, "outline", type="NONE")
        sub(c, HH, "shadow", type="NONE", color="#C0C0C0", offsetX="10", offsetY="10")
        if p.superscript:
            sub(c, HH, "supscript")
        elif p.subscript:
            sub(c, HH, "subscript")
        box.set("itemCnt", str(cid + 1))
        self._char_ids[p] = cid
        return cid

    def para_pr(self, p: ParaProps) -> int:
        if p in self._para_ids:
            return self._para_ids[p]
        box = self._container("paraProperties")
        pid = len(box.findall("hh:paraPr", NS))
        el = sub(box, HH, "paraPr", id=pid, tabPrIDRef="1" if p.auto_tab else "0",
                 condense="0", fontLineHeight="0", snapToGrid="1",
                 suppressLineNumbers="0", checked="0", textDir="LTR")
        sub(el, HH, "align", horizontal=p.align, vertical="BASELINE")
        sub(el, HH, "heading", type="NONE", idRef="0", level="0")
        # Word는 한글 문장을 어절 단위로 줄바꿈한다(한글 기본은 글자 단위).
        sub(el, HH, "breakSetting", breakLatinWord="KEEP_WORD", breakNonLatinWord="KEEP_WORD",
            widowOrphan="1" if p.widow_orphan else "0",
            keepWithNext="1" if p.keep_with_next else "0",
            keepLines="1" if p.keep_lines else "0",
            pageBreakBefore="1" if p.page_break_before else "0", lineWrap="BREAK")
        sub(el, HH, "autoSpacing", eAsianEng="0", eAsianNum="0")
        switch = sub(el, HP, "switch")
        # hp:case는 실제 HWPUNIT, hp:default는 옛 판독기용으로 길이를 2배로 적는다.
        for branch, scale in ((sub(switch, HP, "case"), 1), (sub(switch, HP, "default"), 2)):
            if scale == 1:
                branch.set(q(HP, "required-namespace"), "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar")
            m = sub(branch, HH, "margin")
            for tag, v in (("intent", p.indent), ("left", p.left), ("right", p.right),
                           ("prev", p.before), ("next", p.after)):
                sub(m, HC, tag, value=v * scale, unit="HWPUNIT")
            value = p.line_value if p.line_type == "PERCENT" else p.line_value * scale
            sub(branch, HH, "lineSpacing", type=p.line_type, value=value, unit="HWPUNIT")
        sub(el, HH, "border", borderFillIDRef="2", offsetLeft="0", offsetRight="0",
            offsetTop="0", offsetBottom="0", connect="0", ignoreMargin="0")
        box.set("itemCnt", str(pid + 1))
        self._para_ids[p] = pid
        return pid

    # ---------------------------------------------------------------- 본문
    def _paragraph(self, parent, para: Paragraph, first_in_section=None):
        pid = self.para_pr(para.props)
        if first_in_section is not None:
            p = first_in_section
            p.set("paraPrIDRef", str(pid))
        else:
            p = sub(parent, HP, "p", id=self._uid(), paraPrIDRef=pid, styleIDRef="0",
                    pageBreak="1" if para.page_break else "0",
                    columnBreak="1" if para.column_break else "0", merged="0")
        if first_in_section is not None and para.page_break:
            p.set("pageBreak", "1")

        runs = [r for r in para.runs]
        if not any(r.kind in ("text", "tab", "break", "image") for r in runs):
            run = sub(p, HP, "run", charPrIDRef=self.char_pr(para.mark))
            sub(run, HP, "t")
            return p

        current = None     # (run 요소, charPr id, 마지막 hp:t)
        for r in runs:
            cid = self.char_pr(r.props)
            if current is None or current[1] != cid or r.kind == "image":
                run_el = sub(p, HP, "run", charPrIDRef=cid)
                current = [run_el, cid, None]
            run_el = current[0]
            if r.kind == "text":
                if current[2] is None:
                    current[2] = sub(run_el, HP, "t")
                t = current[2]
                if len(t):
                    t[-1].tail = (t[-1].tail or "") + r.text
                else:
                    t.text = (t.text or "") + r.text
            elif r.kind == "break":
                if current[2] is None:
                    current[2] = sub(run_el, HP, "t")
                sub(current[2], HP, "lineBreak")
            elif r.kind == "tab":
                sub(run_el, HP, "tab")
                current[2] = None
            elif r.kind == "image":
                self._picture(run_el, r.image)
                current = None
        return p

    def _picture(self, run_el, image: Image):
        n = len(self._images) + 1
        item = "BIN%04d" % n
        self._images.append((item, image.data, image.ext))
        w, h = max(image.width, 1), max(image.height, 1)
        uid = self._uid()
        pic = sub(run_el, HP, "pic", textWrap="SQUARE", textFlow="BOTH_SIDES", reverse="0",
                  id=uid, zOrder="0", numberingType="PICTURE", lock="0", dropcapstyle="None",
                  href="", groupLevel="0", instid=uid)
        sub(pic, HP, "offset", x="0", y="0")
        sub(pic, HP, "orgSz", width=w, height=h)
        sub(pic, HP, "curSz", width=w, height=h)
        sub(pic, HP, "flip", horizontal="0", vertical="0")
        sub(pic, HP, "rotationInfo", angle="0", centerX=w // 2, centerY=h // 2, rotateimage="1")
        ri = sub(pic, HP, "renderingInfo")
        for m in ("transMatrix", "scaMatrix", "rotMatrix"):
            sub(ri, HC, m, e1="1", e2="0", e3="0", e4="0", e5="1", e6="0")
        rect = sub(pic, HP, "imgRect")
        for i, (x, y) in enumerate(((0, 0), (w, 0), (w, h), (0, h))):
            sub(rect, HC, f"pt{i}", x=x, y=y)
        sub(pic, HP, "imgClip", left="0", right=w, top="0", bottom=h)
        sub(pic, HP, "inMargin", left="0", right="0", top="0", bottom="0")
        sub(pic, HP, "imgDim", dimwidth=w, dimheight=h)
        sub(pic, HC, "img", binaryItemIDRef=item, bright="0", contrast="0", effect="REAL_PIC", alpha="0")
        sub(pic, HP, "effects")
        sub(pic, HP, "sz", width=w, widthRelTo="ABSOLUTE", height=h, heightRelTo="ABSOLUTE", protect="0")
        sub(pic, HP, "pos", treatAsChar="1", affectLSpacing="0", flowWithText="1", allowOverlap="0",
            holdAnchorAndSO="0", vertRelTo="PARA", horzRelTo="COLUMN", vertAlign="TOP",
            horzAlign="LEFT", vertOffset="0", horzOffset="0")
        sub(pic, HP, "outMargin", left="0", right="0", top="0", bottom="0")
        sub(pic, HP, "shapeComment")

    # ------------------------------------------------------------------ 표
    @staticmethod
    def _content_height(blocks) -> int:
        """셀 내용의 최소 높이 추정. 모자라면 한글이 행을 늘려 주므로 작게 잡는다."""
        total = 0
        for b in blocks:
            if isinstance(b, Table):
                total += sum(b.row_heights) or 1000
                continue
            size = max([r.props.size for r in b.runs if r.kind == "text"] or [b.mark.size])
            if b.props.line_type == "PERCENT":
                line = size * b.props.line_value // 100
            else:
                line = b.props.line_value
            total += line + b.props.before + b.props.after
        return total

    def _table(self, run_el, table: Table):
        n_rows, n_cols = table.rows, table.cols
        heights = list(table.row_heights)
        for cell in table.cells:
            if cell.row_span == 1:
                need = self._content_height(cell.blocks) + cell.margins[2] + cell.margins[3]
                heights[cell.row] = max(heights[cell.row], need)
        heights = [h or 1000 for h in heights]
        widths = table.col_widths
        total_w, total_h = sum(widths), sum(heights)

        uid = self._uid()
        no_border = self.border_fill(BorderFillProps())
        tbl = sub(run_el, HP, "tbl", id=uid, zOrder="0", numberingType="TABLE",
                  textWrap="TOP_AND_BOTTOM", textFlow="BOTH_SIDES", lock="0",
                  dropcapstyle="None", pageBreak="CELL",
                  repeatHeader="1" if table.repeat_header else "0",
                  rowCnt=n_rows, colCnt=n_cols, cellSpacing="0",
                  borderFillIDRef=no_border, noAdjust="0")
        sub(tbl, HP, "sz", width=total_w, widthRelTo="ABSOLUTE", height=total_h,
            heightRelTo="ABSOLUTE", protect="0")
        sub(tbl, HP, "pos", treatAsChar="1", affectLSpacing="0", flowWithText="1",
            allowOverlap="0", holdAnchorAndSO="0", vertRelTo="PARA", horzRelTo="COLUMN",
            vertAlign="TOP", horzAlign="LEFT", vertOffset="0", horzOffset="0")
        sub(tbl, HP, "outMargin", left="0", right="0", top="0", bottom="0")
        ml, mr, mt, mb = table.margins
        sub(tbl, HP, "inMargin", left=ml, right=mr, top=mt, bottom=mb)

        for r in range(n_rows):
            tr = sub(tbl, HP, "tr")
            for cell in sorted((c for c in table.cells if c.row == r), key=lambda c: c.col):
                tc = sub(tr, HP, "tc", name="", header="1" if cell.is_header else "0",
                         hasMargin="1" if cell.margins != table.margins else "0",
                         protect="0", editable="0", dirty="0",
                         borderFillIDRef=self.border_fill(cell.border))
                sl = sub(tc, HP, "subList", id="", textDirection="HORIZONTAL", lineWrap="BREAK",
                         vertAlign=cell.valign, linkListIDRef="0", linkListNextIDRef="0",
                         textWidth="0", textHeight="0", hasTextRef="0", hasNumRef="0")
                self._blocks(sl, cell.blocks)
                sub(tc, HP, "cellAddr", colAddr=cell.col, rowAddr=cell.row)
                sub(tc, HP, "cellSpan", colSpan=cell.col_span, rowSpan=cell.row_span)
                sub(tc, HP, "cellSz", width=sum(widths[cell.col:cell.col + cell.col_span]),
                    height=sum(heights[cell.row:cell.row + cell.row_span]))
                l, rr, t, b = cell.margins
                sub(tc, HP, "cellMargin", left=l, right=rr, top=t, bottom=b)

    # ---------------------------------------------------------------- 흐름
    def _blocks(self, parent, blocks, first_holder=None):
        """블록 목록을 parent 아래에 쓴다. first_holder가 있으면 첫 블록을 거기 담는다."""
        self._apply_contextual_spacing(blocks)
        for i, block in enumerate(blocks):
            holder = first_holder if i == 0 else None
            if isinstance(block, Table):
                # 표는 '글자처럼 취급'한 채로 전용 문단에 담는다. 문단 정렬로 표 위치를 정한다.
                host = Paragraph(props=ParaProps(align=block.align, left=block.indent,
                                                 line_type="PERCENT", line_value=100))
                p = self._paragraph(parent, host, holder)
                for child in list(p):
                    if child.tag == q(HP, "run") and child.find(q(HP, "secPr")) is None \
                            and child.find(q(HP, "ctrl")) is None:
                        p.remove(child)
                run_el = sub(p, HP, "run", charPrIDRef=self.char_pr(host.mark))
                self._table(run_el, block)
                sub(run_el, HP, "t")
            else:
                self._paragraph(parent, block, holder)

    @staticmethod
    def _apply_contextual_spacing(blocks):
        """Word의 '같은 스타일 문단 사이 간격 없음'을 이웃 문단 간격 0으로 바꾼다."""
        from dataclasses import replace
        paras = [b for b in blocks]
        for a, b in zip(paras, paras[1:]):
            if not isinstance(a, Paragraph) or not isinstance(b, Paragraph):
                continue
            if getattr(a, "_contextual", False) and getattr(b, "_contextual", False) \
                    and getattr(a, "_style", None) == getattr(b, "_style", None):
                a.props = replace(a.props, after=0)
                b.props = replace(b.props, before=0)

    def _section_xml(self, section: Section) -> bytes:
        root = copy.deepcopy(self.skeleton_section)
        first = root.find("hp:p", NS)
        for p in root.findall("hp:p", NS)[1:]:
            root.remove(p)
        # 뼈대 첫 문단에서 구역 설정 run만 남긴다 (빈 글자 run과 줄 배치 캐시는 버린다)
        for child in list(first):
            if child.tag == q(HP, "run") and child.find(q(HP, "secPr")) is None:
                first.remove(child)
            elif child.tag == q(HP, "linesegarray"):
                first.remove(child)
        first.set("id", self._uid())

        page = section.page
        pr = first.find(".//hp:pagePr", NS)
        # 한글 2024가 저장한 세로 A4가 WIDELY. 가로는 NARROWLY (실파일 미확인).
        pr.set("landscape", "NARROWLY" if page.landscape else "WIDELY")
        pr.set("width", str(page.width))
        pr.set("height", str(page.height))
        m = pr.find("hp:margin", NS)
        for k in ("header", "footer", "gutter", "left", "right", "top", "bottom"):
            m.set(k, str(getattr(page, k)))

        blocks = section.blocks or [Paragraph()]
        self._blocks(root, blocks, first_holder=first)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # -------------------------------------------------------------- 패키지
    def build(self) -> bytes:
        sections = [self._section_xml(s) for s in self.doc.sections]
        self.header.set("secCnt", str(len(sections)))

        if self._images:
            box = self.ref.find("hh:binDataList", NS)
            if box is None:
                box = sub(self.ref, HH, "binDataList")
            for i, (item, _data, ext) in enumerate(self._images):
                sub(box, HH, "binItem", id=i, Type="Embedding", BinData=f"{item}.{ext}", Format=ext)
            box.set("itemCnt", str(len(self._images)))

        # content.hpf: 구역과 그림을 목록에 올린다
        hpf = etree.fromstring(self.parts["Contents/content.hpf"])
        manifest = hpf.find("opf:manifest", NS)
        spine = hpf.find("opf:spine", NS)
        for item in list(manifest):
            if item.get("id", "").startswith("section"):
                manifest.remove(item)
        for ref in list(spine):
            if ref.get("idref", "").startswith("section"):
                spine.remove(ref)
        settings = next((i for i in manifest if i.get("id") == "settings"), None)
        for n in range(len(sections)):
            item = etree.Element(q(OPF, "item"), id=f"section{n}",
                                 href=f"Contents/section{n}.xml", **{"media-type": "application/xml"})
            if settings is not None:
                settings.addprevious(item)
            else:
                manifest.append(item)
            etree.SubElement(spine, q(OPF, "itemref"), idref=f"section{n}", linear="yes")
        for item, _data, ext in self._images:
            el = etree.SubElement(manifest, q(OPF, "item"), id=item, href=f"BinData/{item}.{ext}")
            el.set("media-type", MEDIA_TYPES.get(ext, "image/" + ext))
            el.set("isEmbeded", "1")

        rdf = self.parts["META-INF/container.rdf"].decode("utf-8")
        extra = "".join(
            f'<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" '
            f'rdf:resource="Contents/section{n}.xml"/></rdf:Description>'
            f'<rdf:Description rdf:about="Contents/section{n}.xml"><rdf:type '
            f'rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description>'
            for n in range(1, len(sections)))
        rdf = rdf.replace("</rdf:RDF>", extra + "</rdf:RDF>")

        text = "\r\n".join(b.text for s in self.doc.sections for b in s.blocks
                           if isinstance(b, Paragraph))[:1024]

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w") as zf:
            zf.writestr(zipfile.ZipInfo("mimetype"), self.parts["mimetype"], compress_type=zipfile.ZIP_STORED)
            D = zipfile.ZIP_DEFLATED
            zf.writestr("version.xml", self.parts["version.xml"], D)
            zf.writestr("Contents/header.xml", etree.tostring(
                self.header, xml_declaration=True, encoding="UTF-8", standalone=True), D)
            for n, xml in enumerate(sections):
                zf.writestr(f"Contents/section{n}.xml", xml, D)
            zf.writestr("Preview/PrvText.txt", (text + "\r\n").encode("utf-8"), D)
            zf.writestr("settings.xml", self.parts["settings.xml"], D)
            zf.writestr("Preview/PrvImage.png", self.parts["Preview/PrvImage.png"], D)
            zf.writestr("META-INF/container.rdf", rdf.encode("utf-8"), D)
            zf.writestr("Contents/content.hpf", etree.tostring(
                hpf, xml_declaration=True, encoding="UTF-8", standalone=True), D)
            zf.writestr("META-INF/container.xml", self.parts["META-INF/container.xml"], D)
            zf.writestr("META-INF/manifest.xml", self.parts["META-INF/manifest.xml"], D)
            for item, data, ext in self._images:
                zf.writestr(f"BinData/{item}.{ext}", data, D)
        return out.getvalue()


def write_hwpx(doc: Document, path: str) -> None:
    data = HwpxWriter(doc).build()
    with open(path, "wb") as fh:
        fh.write(data)
