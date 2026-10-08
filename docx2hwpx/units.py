"""단위 변환과 글꼴 줄 높이.

Word와 한글은 '줄 간격 %'의 기준이 다르다.

  한글:  줄 높이 = 글자 크기 × P%            (11pt, 160% → 17.6pt)
  Word:  줄 높이 = 글꼴 고유 줄 높이 × 배수   (맑은 고딕 11pt, 1.0 → 약 14.6pt)

'글꼴 고유 줄 높이'는 글꼴 파일의 OS/2 표에 적힌 winAscent + winDescent를
em 크기로 나눈 값이다(Word가 Windows에서 쓰는 값). 그래서 Word의 배수를
한글의 %로 옮기려면 그 글꼴의 비율을 알아야 한다.

윈도우에서는 설치된 글꼴 파일을 직접 읽어 정확한 비율을 쓰고, 찾지 못하면
아래 표의 근삿값을 쓴다.
"""

from __future__ import annotations

import os
import struct
import sys
from functools import lru_cache

# ---------------------------------------------------------------------------
# 길이
# ---------------------------------------------------------------------------
#: 1 twip(1/1440 inch) = 5 HWPUNIT(1/7200 inch)
HWPUNIT_PER_TWIP = 5
#: 1 EMU = 1/914400 inch → HWPUNIT = EMU / 127
EMU_PER_HWPUNIT = 127


def twip(v: int | float | None) -> int:
    return int(round((v or 0) * HWPUNIT_PER_TWIP))


def emu(v: int | float | None) -> int:
    return int(round((v or 0) / EMU_PER_HWPUNIT))


def half_points_to_size(v: int | float) -> int:
    """Word w:sz(1/2 pt) → 한글 글자 크기(1/100 pt)."""
    return int(round(v * 50))


# ---------------------------------------------------------------------------
# 테두리 두께: 한글은 정해진 값 중 하나만 받는다
# ---------------------------------------------------------------------------
HWP_BORDER_WIDTHS_MM = (0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5,
                        0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0)


def border_width(eighth_points: int | None) -> str:
    """Word w:sz(1/8 pt) → 가장 가까운 한글 테두리 두께 문자열."""
    mm = (eighth_points or 4) / 8.0 * 25.4 / 72.0
    best = min(HWP_BORDER_WIDTHS_MM, key=lambda w: abs(w - mm))
    return f"{best} mm" if best != int(best) else f"{best:.1f} mm"


# ---------------------------------------------------------------------------
# 글꼴 줄 높이 비율
# ---------------------------------------------------------------------------
#: 글꼴 파일을 못 찾을 때 쓰는 값 (winAscent + winDescent) / unitsPerEm.
#: 라틴 글꼴은 널리 알려진 값, 한글 글꼴은 근삿값이다.
FALLBACK_LINE_RATIO = {
    "calibri": 1.2207, "calibri light": 1.2207, "cambria": 1.1724,
    "arial": 1.1172, "times new roman": 1.1074, "segoe ui": 1.3301,
    "courier new": 1.1328, "consolas": 1.1709, "verdana": 1.2153,
    "tahoma": 1.2070, "georgia": 1.1362, "aptos": 1.2002,
    "맑은 고딕": 1.3301, "malgun gothic": 1.3301,
    "바탕": 1.1602, "batang": 1.1602, "바탕체": 1.1602,
    "굴림": 1.1602, "gulim": 1.1602, "굴림체": 1.1602,
    "돋움": 1.1602, "dotum": 1.1602, "돋움체": 1.1602,
    "궁서": 1.1602, "gungsuh": 1.1602,
    "나눔고딕": 1.2002, "nanumgothic": 1.2002,
    "나눔명조": 1.2002, "nanummyeongjo": 1.2002,
    "함초롬바탕": 1.2002, "함초롬돋움": 1.2002,
}
DEFAULT_LINE_RATIO = 1.2

#: 영문 이름 ↔ 한글 이름 (Word 문서와 윈도우 레지스트리가 서로 다르게 적는다)
FONT_ALIASES = {
    "malgun gothic": "맑은 고딕", "batang": "바탕", "batangche": "바탕체",
    "gulim": "굴림", "gulimche": "굴림체", "dotum": "돋움",
    "dotumche": "돋움체", "gungsuh": "궁서", "gungsuhche": "궁서체",
    "nanumgothic": "나눔고딕", "nanummyeongjo": "나눔명조",
}


def _read_ttf_ratio(path: str) -> float | None:
    """TTF/OTF/TTC 파일에서 (winAscent + winDescent) / unitsPerEm."""
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    base = 0
    if data[:4] == b"ttcf":
        # 글꼴 모음은 첫 번째 글꼴을 쓴다(같은 모음 안은 줄 높이가 같다).
        if len(data) < 16:
            return None
        base = struct.unpack_from(">I", data, 12)[0]
    try:
        num_tables = struct.unpack_from(">H", data, base + 4)[0]
        tables = {}
        for i in range(num_tables):
            rec = base + 12 + i * 16
            tag = data[rec : rec + 4]
            offset = struct.unpack_from(">I", data, rec + 8)[0]
            tables[tag] = offset
        head = tables.get(b"head")
        os2 = tables.get(b"OS/2")
        if head is None or os2 is None:
            return None
        units_per_em = struct.unpack_from(">H", data, head + 18)[0]
        win_ascent, win_descent = struct.unpack_from(">HH", data, os2 + 74)
        if not units_per_em:
            return None
        ratio = (win_ascent + win_descent) / units_per_em
        return ratio if 0.8 < ratio < 3.0 else None
    except (struct.error, IndexError):
        return None


@lru_cache(maxsize=1)
def _windows_font_files() -> dict[str, str]:
    """레지스트리에서 '글꼴 이름 → 파일 경로' 표를 만든다. 윈도우가 아니면 빈 표."""
    if sys.platform != "win32":
        return {}
    try:
        import winreg
    except ImportError:
        return {}
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    user_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    table: dict[str, str] = {}
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, key_path)
        except OSError:
            continue
        index = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, index)
            except OSError:
                break
            index += 1
            if not isinstance(value, str):
                continue
            path = value if os.path.isabs(value) else os.path.join(
                fonts_dir if hive == winreg.HKEY_LOCAL_MACHINE else user_dir, value)
            clean = name.split(" (")[0]
            for family in clean.split(" & "):
                table.setdefault(family.strip().lower(), path)
        winreg.CloseKey(key)
    return table


@lru_cache(maxsize=256)
def line_ratio(font_name: str) -> float:
    """글꼴의 줄 높이 / 글자 크기 비율. Word의 '줄 간격 1.0'이 글자 크기의 몇 배인지."""
    name = (font_name or "").strip().lower()
    candidates = [name]
    alias = FONT_ALIASES.get(name)
    if alias:
        candidates.append(alias.lower())
    for eng, kor in FONT_ALIASES.items():
        if kor.lower() == name:
            candidates.append(eng)

    files = _windows_font_files()
    for cand in candidates:
        path = files.get(cand)
        if path:
            ratio = _read_ttf_ratio(path)
            if ratio:
                return ratio
    for cand in candidates:
        if cand in FALLBACK_LINE_RATIO:
            return FALLBACK_LINE_RATIO[cand]
    return DEFAULT_LINE_RATIO


def word_multiple_to_hwp_percent(multiple: float, font_name: str) -> int:
    """Word '배수' 줄 간격 → 한글 '글자에 따라 %' 줄 간격."""
    return max(50, min(500, int(round(multiple * line_ratio(font_name) * 100))))
