"""원본 .docx와 변환된 .hwpx를 문단 단위로 대조한다.

양쪽 모두 변환기 코드를 거치지 않고 읽는다.
  - Word: document.xml을 직접 읽는다 (직접 서식만 쓰는 문서 기준)
  - 한글: python-hwpx(독립 구현)로 글자를, header.xml/section*.xml을 lxml로 서식을 읽는다
"""

import re
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def word_side(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    paras = []
    for p in body.iter(W + "p"):
        ppr = p.find(W + "pPr")
        sp = ppr.find(W + "spacing") if ppr is not None else None
        jc = ppr.find(W + "jc") if ppr is not None else None
        r = p.find(f".//{W}r/{W}rPr")
        g = lambda el, a, d=None: el.get(W + a, d) if el is not None else d
        fonts = r.find(W + "rFonts") if r is not None else None
        paras.append({
            "text": "".join(t.text or "" for t in p.iter(W + "t")),
            "align": {"center": "CENTER", "right": "RIGHT", "both": "JUSTIFY"}.get(g(jc, "val"), "LEFT"),
            "before": int(g(sp, "before", 0)), "after": int(g(sp, "after", 0)),
            "line": int(g(sp, "line", 240)), "lineRule": g(sp, "lineRule", "auto"),
            "size": int(g(r.find(W + "sz") if r is not None else None, "val", 20)) / 2,
            "bold": r is not None and r.find(W + "b") is not None,
            "font": g(fonts, "eastAsia") or g(fonts, "ascii"),
        })
    sect = body.find(W + "sectPr")
    sz, mar = sect.find(W + "pgSz"), sect.find(W + "pgMar")
    page = {k: int(mar.get(W + k)) for k in ("top", "bottom", "left", "right", "header", "footer")}
    page.update(w=int(sz.get(W + "w")), h=int(sz.get(W + "h")))
    return paras, page


def hwp_side(path):
    z = zipfile.ZipFile(path)
    head = etree.fromstring(z.read("Contents/header.xml"))
    fonts = {f.get("id"): f.get("face") for f in head.find(f".//{HH}fontface").findall(HH + "font")}
    chars = {c.get("id"): c for c in head.iter(HH + "charPr")}
    paras_pr = {p.get("id"): p for p in head.iter(HH + "paraPr")}
    sec = etree.fromstring(z.read("Contents/section0.xml"))
    out = []
    for p in sec.iterfind(HP + "p"):
        pp = paras_pr[p.get("paraPrIDRef")]
        case = pp.find(f"{HP}switch/{HP}case")
        m = case.find(HH + "margin")
        ls = case.find(HH + "lineSpacing")
        text_runs = [r for r in p.iterfind(HP + "run") if r.find(HP + "t") is not None]
        run = next((r for r in text_runs if "".join(r.find(HP + "t").itertext())), text_runs[0] if text_runs else None)
        c = chars[run.get("charPrIDRef")] if run is not None else None
        out.append({
            "text": "".join("".join(t.itertext()) for t in p.iter(HP + "t")),
            "align": pp.find(HH + "align").get("horizontal"),
            "before": int(m.find(HC + "prev").get("value")), "after": int(m.find(HC + "next").get("value")),
            "line_type": ls.get("type"), "line": int(ls.get("value")),
            "size": int(c.get("height")) / 100 if c is not None else None,
            "bold": c is not None and c.find(HH + "bold") is not None,
            "font": fonts.get(c.find(HH + "fontRef").get("hangul")) if c is not None else None,
            "wrap": pp.find(HH + "breakSetting").get("breakNonLatinWord"),
        })
    pr = sec.find(f".//{HP}pagePr")
    mg = pr.find(HP + "margin")
    page = {k: int(mg.get(k)) for k in ("top", "bottom", "left", "right", "header", "footer")}
    page.update(w=int(pr.get("width")), h=int(pr.get("height")), landscape=pr.get("landscape"))
    return out, page


def independent_text(path):
    from hwpx import HwpxDocument
    doc = HwpxDocument.open(path)
    return [p.text for p in doc.paragraphs]


def compare(docx_path, hwpx_path, ratio_note):
    wp, wpage = word_side(docx_path)
    hp, hpage = hwp_side(hwpx_path)
    problems, notes = [], []

    # 1. 글자 (독립 판독기)
    ind = [t for t in independent_text(hwpx_path)]
    if [p["text"] for p in wp] != ind:
        problems.append(f"글자 불일치: Word {len(wp)}문단 / 한글(python-hwpx) {len(ind)}문단")
        for i, (a, b) in enumerate(zip([p['text'] for p in wp], ind)):
            if a != b:
                problems.append(f"  문단 {i+1}: {a!r} ≠ {b!r}")
    chars = sum(len(p["text"]) for p in wp)

    # 2. 문단 서식
    if len(wp) != len(hp):
        problems.append(f"문단 수: Word {len(wp)} / 한글 {len(hp)}")
    for i, (a, b) in enumerate(zip(wp, hp), 1):
        exp = {
            "align": a["align"], "before": a["before"] * 5, "after": a["after"] * 5,
            "size": a["size"], "bold": a["bold"], "font": a["font"],
        }
        for k, v in exp.items():
            if b[k] != v:
                problems.append(f"문단 {i} {k}: 기대 {v} / 실제 {b[k]}")
        if a["lineRule"] == "auto":
            if b["line_type"] != "PERCENT":
                problems.append(f"문단 {i} 줄 간격 종류: PERCENT 기대 / {b['line_type']}")
        if b["wrap"] != "KEEP_WORD":
            problems.append(f"문단 {i} 줄바꿈 단위가 어절이 아님")
    ln = {(a["line"], a["lineRule"]) for a in wp}
    lh = {(b["line_type"], b["line"]) for b in hp}
    notes.append(f"줄 간격: Word {sorted(ln)} → 한글 {sorted(lh)}  ({ratio_note})")

    # 3. 용지
    exp_page = {
        "w": wpage["w"] * 5, "h": wpage["h"] * 5, "left": wpage["left"] * 5, "right": wpage["right"] * 5,
        "top": wpage["header"] * 5, "header": (wpage["top"] - wpage["header"]) * 5,
        "bottom": wpage["footer"] * 5, "footer": (wpage["bottom"] - wpage["footer"]) * 5,
    }
    for k, v in exp_page.items():
        if hpage[k] != v:
            problems.append(f"용지 {k}: 기대 {v} / 실제 {hpage[k]}")
    body_top_word = wpage["top"] * 5
    body_top_hwp = hpage["top"] + hpage["header"]
    if body_top_word != body_top_hwp:
        problems.append(f"본문 시작 높이: Word {body_top_word} / 한글 {body_top_hwp}")
    notes.append(f"용지 {hpage['w']/7200*25.4:.1f}×{hpage['h']/7200*25.4:.1f}mm ({hpage['landscape']}), "
                 f"본문 시작 {body_top_hwp/7200*25.4:.1f}mm, 좌우 여백 {hpage['left']/7200*25.4:.1f}mm")
    return problems, notes, len(wp), chars


if __name__ == "__main__":
    for docx_path, hwpx_path in zip(sys.argv[1::2], sys.argv[2::2]):
        from docx2hwpx.units import line_ratio
        problems, notes, n, chars = compare(docx_path, hwpx_path,
                                            f"휴먼명조 줄 높이 비율 {line_ratio('휴먼명조')} 사용")
        print(f"== {docx_path.split('/')[-1]}  ({n}문단, {chars}자)")
        print("   문제:", "없음" if not problems else "")
        for p in problems:
            print("    ✗", p)
        for x in notes:
            print("    ·", x)
