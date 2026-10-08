"""작성기 경로 시험.

원본 .hwp 하나로는 지나가지 않는 경로 — 가로/세로 병합, 중첩 표, 그림,
여러 구역, 배치 정보가 없는 문서 — 를 합성 모델로 만들어 확인한다.
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from hwp2docx.docinfo import (  # noqa: E402
    ALIGN_CENTER,
    Border,
    BorderFill,
    CharShape,
    DocInfo,
    FaceName,
    ParaShape,
)
from hwp2docx.model import Cell, Document, Image, Paragraph, Run, Section, Table  # noqa: E402
from hwp2docx.writer import write_docx  # noqa: E402

NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _png() -> bytes:
    from PIL import Image
    import io
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), (200, 30, 30)).save(buf, "PNG")
    return buf.getvalue()


PNG_1X1 = _png()


def make_info() -> DocInfo:
    info = DocInfo()
    info.face_names = [FaceName("맑은 고딕"), FaceName("Calibri")]
    info.char_shapes = [
        CharShape(face_ids=(0, 1, 0, 0, 0, 0, 0), base_size=1000),
        CharShape(face_ids=(0, 1, 0, 0, 0, 0, 0), base_size=1200, bold=True, color="FF0000"),
    ]
    info.para_shapes = [
        ParaShape(),
        ParaShape(align=ALIGN_CENTER, space_before=400, space_after=400, line_spacing=160),
    ]
    solid = Border("single", 4, "000000")
    info.border_fills = [
        BorderFill(solid, solid, solid, solid, fill_color=None),
        BorderFill(solid, solid, solid, solid, fill_color="DDDDDD"),
    ]
    return info


def para(text: str = "", shape: int = 0, char: int = 0, line_height: int = 0,
         top: int | None = None) -> Paragraph:
    p = Paragraph(para_shape_id=shape, default_char_shape_id=char)
    if text:
        p.runs.append(Run(text, char))
    if line_height:
        p.line_height = line_height
        p.layout_top = top or 0
        p.layout_bottom = (top or 0) + line_height
    return p


def cell(row: int, col: int, text: str = "", **kw) -> Cell:
    c = Cell(col=col, row=row, width=kw.pop("width", 10000),
             height=kw.pop("height", 1000), border_fill_id=kw.pop("bf", 1), **kw)
    c.paragraphs = [para(text)]
    return c


def build(path: str, doc: Document, **kw) -> None:
    write_docx(doc, make_info(), path, **kw)


def read_xml(path: str) -> bytes:
    with zipfile.ZipFile(path) as zf:
        return zf.read("word/document.xml")


def validate(path: str) -> str:
    script = "/mnt/skills/public/docx/scripts/office/validate.py"
    if not os.path.exists(script):
        return "SKIPPED"
    out = subprocess.run([sys.executable, script, path], capture_output=True, text=True)
    return out.stdout + out.stderr


# ---------------------------------------------------------------------------

def test_merged_cells():
    """가로 2칸 + 세로 2칸 병합이 gridSpan / vMerge로 나가야 한다."""
    table = Table(rows=3, cols=3, border_fill_id=1, total_width=30000, total_height=3000)
    table.cells = [
        Cell(col=0, row=0, col_span=2, width=20000, height=1000, border_fill_id=1,
             paragraphs=[para("가로병합")]),
        cell(0, 2, "C"),
        Cell(col=0, row=1, row_span=2, width=10000, height=1000, border_fill_id=1,
             paragraphs=[para("세로병합")]),
        cell(1, 1, "B2"), cell(1, 2, "C2"),
        cell(2, 1, "B3"), cell(2, 2, "C3"),
    ]
    doc = Document(sections=[Section(paragraphs=[
        Paragraph(tables=[table]), para("끝")])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "merge.docx")
        build(out, doc)
        xml = read_xml(out).decode()
        assert f'{NS}gridSpan'.replace(NS, "w:") or True
        assert 'w:gridSpan w:val="2"' in xml, "가로 병합이 없습니다"
        assert 'w:vMerge w:val="restart"' in xml, "세로 병합 시작이 없습니다"
        assert xml.count("<w:vMerge/>") == 1, "세로 병합 이어짐 칸이 1개여야 합니다"
        assert "PASSED" in validate(out) or "SKIPPED" == validate(out).strip()
        print("  병합 OK")


def test_nested_table():
    inner = Table(rows=1, cols=1, border_fill_id=1, total_width=8000, total_height=1000)
    inner.cells = [cell(0, 0, "안쪽", width=8000)]
    host = para()
    host.tables = [inner]
    outer = Table(rows=1, cols=1, border_fill_id=1, total_width=10000, total_height=2000)
    outer_cell = Cell(col=0, row=0, width=10000, height=2000, border_fill_id=1)
    outer_cell.paragraphs = [host]
    outer.cells = [outer_cell]

    doc = Document(sections=[Section(paragraphs=[Paragraph(tables=[outer]), para("끝")])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "nested.docx")
        build(out, doc)
        xml = read_xml(out).decode()
        assert xml.count("<w:tbl>") == 2, "중첩 표가 2개여야 합니다"
        assert "안쪽" in xml
        assert "PASSED" in validate(out) or validate(out).strip() == "SKIPPED"
        print("  중첩 표 OK")


def test_image():
    p = para()
    p.images = [Image(data=PNG_1X1, ext="png", width=7200, height=7200)]
    doc = Document(sections=[Section(paragraphs=[p])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "img.docx")
        build(out, doc)
        with zipfile.ZipFile(out) as zf:
            media = [n for n in zf.namelist() if n.startswith("word/media/")]
        assert media, "이미지가 패키지에 들어가지 않았습니다"
        assert "<w:drawing>" in read_xml(out).decode()
        assert "PASSED" in validate(out) or validate(out).strip() == "SKIPPED"
        print("  그림 OK")


def test_multi_section_and_landscape():
    s1 = Section(paragraphs=[para("첫 구역")])
    s2 = Section(paragraphs=[para("둘째 구역")])
    s2.page.landscape = True
    s2.page.width, s2.page.height = 84188, 59528
    doc = Document(sections=[s1, s2])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "multi.docx")
        build(out, doc)
        xml = read_xml(out).decode()
        assert xml.count("<w:sectPr") == 2, "구역이 2개여야 합니다"
        assert 'w:orient="landscape"' in xml
        assert "PASSED" in validate(out) or validate(out).strip() == "SKIPPED"
        print("  여러 구역 / 가로 용지 OK")


def test_fallback_without_layout():
    """배치 정보가 없으면 문단 모양 값으로 간격을 낸다."""
    doc = Document(sections=[Section(paragraphs=[para("가", shape=1), para("나", shape=1)])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "fallback.docx")
        build(out, doc)
        xml = read_xml(out).decode()
        assert 'w:before="80"' in xml, "문단 위 간격이 반영되지 않았습니다"
        assert 'w:jc w:val="center"' in xml
        assert "PASSED" in validate(out) or validate(out).strip() == "SKIPPED"
        print("  배치 정보 없는 문서 OK")


def test_layout_spacing():
    """배치 좌표가 있으면 그 간격이 before로 나가야 한다."""
    p1 = para("가", line_height=1000, top=1000)
    p2 = para("나", line_height=1000, top=3500)   # 앞 문단 끝(2000)에서 1500 떨어짐
    doc = Document(sections=[Section(paragraphs=[p1, p2])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "layout.docx")
        build(out, doc)
        xml = read_xml(out).decode()
        assert 'w:before="200"' in xml, "첫 문단 간격(1000 HWPUNIT = 200 twip)"
        assert 'w:before="300"' in xml, "둘째 문단 간격(1500 HWPUNIT = 300 twip)"
        assert 'w:lineRule="exact"' in xml
        print("  배치 좌표 기반 간격 OK")


def test_safe_fonts():
    info = make_info()
    info.face_names = [FaceName("함초롬바탕"), FaceName("함초롬바탕")]
    doc = Document(sections=[Section(paragraphs=[para("가")])])
    with tempfile.TemporaryDirectory() as tmp:
        plain = os.path.join(tmp, "plain.docx")
        safe = os.path.join(tmp, "safe.docx")
        write_docx(doc, info, plain)
        write_docx(doc, info, safe, safe_fonts=True)
        assert "함초롬바탕" in read_xml(plain).decode()
        assert "함초롬바탕" not in read_xml(safe).decode()
        assert "맑은 고딕" in read_xml(safe).decode()
        print("  글꼴 대체 OK")


def test_hanging_indent():
    """한글 내어쓰기(첫 줄은 왼쪽 여백, 나머지가 들어감) → Word left+hanging."""
    info = make_info()
    info.para_shapes.append(ParaShape(margin_left=1000, indent=-1310))
    doc = Document(sections=[Section(paragraphs=[para("가", shape=2)])])
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "hang.docx")
        write_docx(doc, info, out)
        xml = read_xml(out).decode()
        # left = (1000 + 1310) HWPUNIT = 462 twip, hanging = 1310 HWPUNIT = 262 twip
        assert 'w:left="462"' in xml and 'w:hanging="262"' in xml, xml[xml.find("<w:ind"):][:80]
        print("  내어쓰기 OK")


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for fn in tests:
        try:
            fn()
        except AssertionError as exc:
            failed += 1
            print(f"  ✗ {fn.__name__}: {exc}")
        except Exception as exc:
            failed += 1
            print(f"  ✗ {fn.__name__}: {type(exc).__name__}: {exc}")
    print(f"\n{len(tests) - failed}/{len(tests)} 통과")
    sys.exit(1 if failed else 0)
