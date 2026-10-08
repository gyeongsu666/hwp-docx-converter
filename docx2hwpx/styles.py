"""Word 서식 상속 해석.

Word 문단 하나의 최종 서식은 여러 층이 겹친 결과다.

    문서 기본값(docDefaults)
      → 표 스타일 (표 안의 문단일 때)
        → 문단 스타일 (basedOn을 따라 뿌리부터)
          → 번호 매기기 수준 (번호 문단일 때)
            → 글자 스타일 (rStyle)
              → 직접 서식

아래층부터 차례로 덮어쓰며 합친다. 특히 표 스타일을 빼먹으면 안 된다:
Word 기본 'Table Grid'는 표 안 문단의 '뒤 간격'을 0으로 만드는데, 이걸
놓치면 한글에서 표의 모든 행이 한 줄씩 높아진다.
"""

from __future__ import annotations

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = "{%s}" % W_NS
A = "{%s}" % A_NS


def w(tag: str) -> str:
    return W + tag


def wattr(el, name: str, default=None):
    if el is None:
        return default
    return el.get(W + name, default)


def onoff(el) -> bool | None:
    """<w:b/> → True, <w:b w:val="0"/> → False, 없으면 None."""
    if el is None:
        return None
    val = el.get(W + "val")
    return val not in ("0", "false", "off")


def to_int(value, default=None):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# 테마 글꼴
# ---------------------------------------------------------------------------

class ThemeFonts:
    """theme1.xml의 글꼴 체계. rFonts의 asciiTheme="minorHAnsi" 같은 값을 푼다."""

    def __init__(self, theme_el=None, script: str = "Hang"):
        self.major = {"latin": "", "ea": "", "cs": ""}
        self.minor = {"latin": "", "ea": "", "cs": ""}
        if theme_el is None:
            return
        for kind, bucket in (("majorFont", self.major), ("minorFont", self.minor)):
            node = theme_el.find(f".//{A}{kind}")
            if node is None:
                continue
            for slot in ("latin", "ea", "cs"):
                el = node.find(A + slot)
                if el is not None:
                    bucket[slot] = el.get("typeface", "")
            # 동아시아 글꼴이 비어 있으면 문자 체계별 글꼴(한글은 Hang)을 쓴다.
            for f in node.findall(A + "font"):
                if f.get("script") == script and not bucket["ea"]:
                    bucket["ea"] = f.get("typeface", "")

    def resolve(self, theme_value: str) -> str:
        v = theme_value or ""
        bucket = self.major if v.startswith("major") else self.minor
        if v.endswith("EastAsia"):
            return bucket["ea"] or bucket["latin"]
        if v.endswith("Bidi"):
            return bucket["cs"] or bucket["latin"]
        return bucket["latin"]


# ---------------------------------------------------------------------------
# 층(layer) 읽기
# ---------------------------------------------------------------------------

def parse_rpr(el) -> dict:
    """<w:rPr> 하나를 '이 층이 정한 속성만' 담은 dict로."""
    d: dict = {}
    if el is None:
        return d
    rf = el.find(w("rFonts"))
    if rf is not None:
        fonts = {}
        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
            theme = rf.get(W + slot + "Theme")
            name = rf.get(W + slot)
            if theme:
                fonts[slot] = ("theme", theme)
            elif name:
                fonts[slot] = ("name", name)
        if fonts:
            d["fonts"] = fonts
    for tag in ("b", "i", "strike", "dstrike", "caps", "smallCaps", "vanish"):
        v = onoff(el.find(w(tag)))
        if v is not None:
            d[tag] = v
    sz = to_int(wattr(el.find(w("sz")), "val"))
    if sz:
        d["sz"] = sz
    u = el.find(w("u"))
    if u is not None:
        d["u"] = wattr(u, "val", "single")
        if wattr(u, "color") and wattr(u, "color") != "auto":
            d["u_color"] = wattr(u, "color")
    color = el.find(w("color"))
    if color is not None and wattr(color, "val"):
        d["color"] = wattr(color, "val")
    hl = el.find(w("highlight"))
    if hl is not None:
        d["highlight"] = wattr(hl, "val")
    shd = el.find(w("shd"))
    if shd is not None:
        d["shd"] = (wattr(shd, "val"), wattr(shd, "fill"), wattr(shd, "color"))
    va = el.find(w("vertAlign"))
    if va is not None:
        d["vertAlign"] = wattr(va, "val")
    sp = to_int(wattr(el.find(w("spacing")), "val"))
    if sp is not None:
        d["spacing"] = sp
    scale = to_int(wattr(el.find(w("w")), "val"))
    if scale:
        d["w"] = scale
    return d


def merge_rpr(base: dict, layer: dict) -> dict:
    out = dict(base)
    for key, value in layer.items():
        if key == "fonts":
            fonts = dict(out.get("fonts", {}))
            fonts.update(value)
            out["fonts"] = fonts
        else:
            out[key] = value
    return out


def parse_ppr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    jc = el.find(w("jc"))
    if jc is not None:
        d["jc"] = wattr(jc, "val")
    ind = el.find(w("ind"))
    if ind is not None:
        i = {}
        for key, names in (("left", ("left", "start")), ("right", ("right", "end"))):
            for n in names:
                v = to_int(wattr(ind, n))
                if v is not None:
                    i[key] = v
                    break
        fl = to_int(wattr(ind, "firstLine"))
        hg = to_int(wattr(ind, "hanging"))
        if hg is not None:
            i["hanging"] = hg
        elif fl is not None:
            i["firstLine"] = fl
        d["ind"] = i
    sp = el.find(w("spacing"))
    if sp is not None:
        s = {}
        for key in ("before", "after", "line"):
            v = to_int(wattr(sp, key))
            if v is not None:
                s[key] = v
        if wattr(sp, "lineRule"):
            s["lineRule"] = wattr(sp, "lineRule")
        for key in ("beforeAutospacing", "afterAutospacing"):
            v = wattr(sp, key)
            if v is not None:
                s[key] = v not in ("0", "false", "off")
        d["spacing"] = s
    for tag in ("contextualSpacing", "keepNext", "keepLines", "pageBreakBefore", "widowControl"):
        v = onoff(el.find(w(tag)))
        if v is not None:
            d[tag] = v
    num = el.find(w("numPr"))
    if num is not None:
        d["numPr"] = (to_int(wattr(num.find(w("numId")), "val")),
                      to_int(wattr(num.find(w("ilvl")), "val"), 0))
    return d


def merge_ppr(base: dict, layer: dict) -> dict:
    out = dict(base)
    for key, value in layer.items():
        if key == "ind":
            ind = dict(out.get("ind", {}))
            # 첫 줄 들여쓰기와 내어쓰기는 서로를 지운다.
            if "hanging" in value:
                ind.pop("firstLine", None)
            if "firstLine" in value:
                ind.pop("hanging", None)
            ind.update(value)
            out["ind"] = ind
        elif key == "spacing":
            spacing = dict(out.get("spacing", {}))
            spacing.update(value)
            out["spacing"] = spacing
        else:
            out[key] = value
    return out


# ---------------------------------------------------------------------------
# 스타일 표
# ---------------------------------------------------------------------------

class StyleSheet:
    def __init__(self, styles_el, theme: ThemeFonts):
        self.theme = theme
        self.styles: dict[str, object] = {}
        self.default_para = None
        self.default_char = None
        self.default_table = None
        self.doc_rpr: dict = {}
        self.doc_ppr: dict = {}
        if styles_el is None:
            return
        for st in styles_el.findall(w("style")):
            sid = wattr(st, "styleId")
            if not sid:
                continue
            self.styles[sid] = st
            if wattr(st, "default") in ("1", "true"):
                kind = wattr(st, "type")
                if kind == "paragraph":
                    self.default_para = sid
                elif kind == "character":
                    self.default_char = sid
                elif kind == "table":
                    self.default_table = sid
        dd = styles_el.find(w("docDefaults"))
        if dd is not None:
            self.doc_rpr = parse_rpr(dd.find(f"{W}rPrDefault/{W}rPr"))
            self.doc_ppr = parse_ppr(dd.find(f"{W}pPrDefault/{W}pPr"))

    def chain(self, style_id: str | None) -> list:
        """basedOn을 따라 올라간 스타일 목록. 뿌리가 앞에 온다."""
        out, seen = [], set()
        sid = style_id
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            out.append(st)
            sid = wattr(st.find(w("basedOn")), "val")
        return list(reversed(out))

    def style_ppr(self, style_id) -> dict:
        d: dict = {}
        for st in self.chain(style_id):
            d = merge_ppr(d, parse_ppr(st.find(w("pPr"))))
        return d

    def style_rpr(self, style_id) -> dict:
        d: dict = {}
        for st in self.chain(style_id):
            d = merge_rpr(d, parse_rpr(st.find(w("rPr"))))
        return d

    def style_element(self, style_id, tag: str) -> list:
        """스타일 체인에서 특정 하위 요소(tblPr, tcPr 등)를 뿌리부터 모은다."""
        return [el for st in self.chain(style_id) for el in [st.find(w(tag))] if el is not None]

    def resolve_font(self, spec) -> str:
        if not spec:
            return ""
        kind, value = spec
        return self.theme.resolve(value) if kind == "theme" else value
