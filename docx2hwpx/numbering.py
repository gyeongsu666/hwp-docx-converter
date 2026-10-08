"""Word 자동 번호·글머리표를 글자로 풀어 쓴다.

한글에도 번호 매기기가 있지만 Word와 규칙(수준별 들여쓰기, 이어 매기기,
재시작)이 달라 그대로 옮기면 번호가 어긋난다. 그래서 Word가 화면에 보여
주는 번호 글자('1.', '가.', '•')를 계산해 문단 앞에 실제 글자로 넣는다.
모양은 같고, 한글에서 번호를 고칠 때만 손으로 해야 한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .styles import W, parse_ppr, parse_rpr, to_int, w, wattr

GANADA = "가나다라마바사아자차카타파하"
CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
KOREAN_DIGITS = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

#: Symbol·Wingdings 글꼴의 사용자 영역 글머리표 → 표준 문자
BULLET_MAP = {
    "": "•", "": "▪", "": "➢", "": "❖", "": "✓",
    "": "■", "": "□", "": "●", "": "➔", "": "➤",
    "": "►", "": "•", "": "❑", "": "◆", "": "◇",
    "o": "◦", "§": "▪", "Ø": "➢", "ü": "✓", "q": "❑", "v": "❖",
}


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def _letters(n: int) -> str:
    # Word: a..z, aa..zz, aaa..
    n = max(1, n)
    return chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1)


def _korean_count(n: int) -> str:
    if n <= 0:
        return str(n)
    units = [(1000, "천"), (100, "백"), (10, "십")]
    out = []
    for v, name in units:
        d = n // v
        if d:
            out.append(("" if d == 1 else KOREAN_DIGITS[d]) + name)
        n %= v
    if n:
        out.append(KOREAN_DIGITS[n])
    return "".join(out)


def format_number(n: int, fmt: str) -> str:
    if fmt in ("decimal", "decimalHalfWidth", "decimalFullWidth"):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt == "upperLetter":
        return _letters(n).upper()
    if fmt == "lowerLetter":
        return _letters(n)
    if fmt == "ganada":
        return GANADA[(n - 1) % len(GANADA)]
    if fmt == "chosung":
        return CHOSUNG[(n - 1) % len(CHOSUNG)]
    if fmt in ("decimalEnclosedCircle", "decimalEnclosedCircleChinese"):
        return CIRCLED[n - 1] if 1 <= n <= len(CIRCLED) else str(n)
    if fmt in ("koreanDigital", "koreanCounting", "koreanLegal", "koreanDigital2"):
        return _korean_count(n)
    if fmt == "none":
        return ""
    return str(n)


@dataclass
class Level:
    start: int = 1
    fmt: str = "decimal"
    text: str = "%1."
    suffix: str = "tab"
    restart: int | None = None
    ppr: dict = field(default_factory=dict)
    rpr: dict = field(default_factory=dict)


def _parse_level(el) -> Level:
    lvl = Level()
    lvl.start = to_int(wattr(el.find(w("start")), "val"), 1)
    lvl.fmt = wattr(el.find(w("numFmt")), "val", "decimal")
    lvl.text = wattr(el.find(w("lvlText")), "val", "")
    lvl.suffix = wattr(el.find(w("suff")), "val", "tab")
    lvl.restart = to_int(wattr(el.find(w("lvlRestart")), "val"))
    lvl.ppr = parse_ppr(el.find(w("pPr")))
    lvl.rpr = parse_rpr(el.find(w("rPr")))
    return lvl


class Numbering:
    def __init__(self, numbering_el=None):
        self.abstract: dict[int, dict[int, Level]] = {}
        self.nums: dict[int, tuple[int, dict[int, Level], dict[int, int]]] = {}
        self.counters: dict[object, list[int]] = {}
        if numbering_el is None:
            return
        for an in numbering_el.findall(w("abstractNum")):
            aid = to_int(wattr(an, "abstractNumId"))
            levels = {}
            for lv in an.findall(w("lvl")):
                levels[to_int(wattr(lv, "ilvl"), 0)] = _parse_level(lv)
            self.abstract[aid] = levels
        for num in numbering_el.findall(w("num")):
            nid = to_int(wattr(num, "numId"))
            aid = to_int(wattr(num.find(w("abstractNumId")), "val"))
            overrides: dict[int, Level] = {}
            starts: dict[int, int] = {}
            for ov in num.findall(w("lvlOverride")):
                il = to_int(wattr(ov, "ilvl"), 0)
                so = ov.find(w("startOverride"))
                if so is not None:
                    starts[il] = to_int(wattr(so, "val"), 1)
                lv = ov.find(w("lvl"))
                if lv is not None:
                    overrides[il] = _parse_level(lv)
            self.nums[nid] = (aid, overrides, starts)

    def level(self, num_id: int | None, ilvl: int) -> Level | None:
        if not num_id or num_id not in self.nums:
            return None
        aid, overrides, _ = self.nums[num_id]
        if ilvl in overrides:
            return overrides[ilvl]
        return self.abstract.get(aid, {}).get(ilvl)

    def next_label(self, num_id: int, ilvl: int) -> tuple[str, Level] | None:
        """이 문단에 붙는 번호 글자를 계산하고 계수기를 한 칸 올린다."""
        lvl = self.level(num_id, ilvl)
        if lvl is None:
            return None
        aid, overrides, starts = self.nums[num_id]
        # 시작 번호를 다시 지정한 목록은 별개로 세고, 아니면 같은 추상 목록끼리 이어 센다.
        key = ("num", num_id) if starts or overrides else ("abs", aid)
        counts = self.counters.setdefault(key, [0] * 9)
        levels = self.abstract.get(aid, {})

        def start_of(i: int) -> int:
            if i in starts:
                return starts[i]
            lv = overrides.get(i) or levels.get(i)
            return lv.start if lv else 1

        if counts[ilvl] == 0:
            counts[ilvl] = start_of(ilvl)
        else:
            counts[ilvl] += 1
        # 하위 수준은 다시 시작
        for deeper in range(ilvl + 1, 9):
            counts[deeper] = 0

        text = lvl.text
        if lvl.fmt == "bullet":
            label = "".join(BULLET_MAP.get(ch, ch) for ch in text)
        else:
            label = text
            for i in range(ilvl + 1):
                lv = overrides.get(i) or levels.get(i) or lvl
                value = counts[i] if counts[i] else start_of(i)
                label = label.replace(f"%{i + 1}", format_number(value, lv.fmt))
        return label, lvl
