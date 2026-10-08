"""한컴 기호 글꼴의 사용자 영역(PUA) 문자를 표준 유니코드로 옮긴다.

한글은 원문자·괄호문자 같은 기호를 유니코드 코드포인트가 아니라
U+F000 대역의 사용자 정의 영역에 저장하는 경우가 있다. 그대로 두면
Word에서 네모(tofu)로 보이므로 대응되는 표준 문자로 바꾼다.

실제 파일에서 확인한 규칙: U+F081..U+F094 가 ①..⑳ 에 대응한다.
"""

from __future__ import annotations

# U+F081 -> ①(U+2460) ... U+F094 -> ⑳(U+2473)
_CIRCLED = {0xF081 + i: 0x2460 + i for i in range(20)}

# 자주 쓰이는 나머지 기호들
_EXTRA = {
    0xF06C: 0x25CF,   # ●
    0xF06D: 0x25A0,   # ■
    0xF06E: 0x25A1,   # □
    0xF0A1: 0x2665,   # ♥
    0xF0B7: 0x2022,   # •
    0xF0D8: 0x25B6,   # ▶
    0xF0E0: 0x2192,   # →
}

PUA_MAP: dict[int, int] = {**_CIRCLED, **_EXTRA}


def translate(ch: str) -> str:
    """PUA 문자 하나를 표준 문자로. 모르는 문자는 그대로 돌려준다."""
    mapped = PUA_MAP.get(ord(ch))
    return chr(mapped) if mapped else ch


def is_unmapped_pua(ch: str) -> bool:
    code = ord(ch)
    return 0xE000 <= code <= 0xF8FF and code not in PUA_MAP


def clean(text: str) -> tuple[str, bool]:
    """문자열 전체를 변환한다. (변환된 문자열, 못 바꾼 PUA가 있었는지)."""
    if not any(0xE000 <= ord(c) <= 0xF8FF for c in text):
        return text, False
    out, leftover = [], False
    for ch in text:
        if 0xE000 <= ord(ch) <= 0xF8FF:
            mapped = PUA_MAP.get(ord(ch))
            if mapped:
                out.append(chr(mapped))
            else:
                leftover = True
                out.append(ch)
        else:
            out.append(ch)
    return "".join(out), leftover
