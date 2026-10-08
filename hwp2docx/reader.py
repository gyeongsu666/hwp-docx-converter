"""HWP 5.x 파일 읽기: OLE 복합 파일 구조와 레코드 스트림.

HWP 5.x 문서는 OLE2 복합 파일이고, 그 안의 각 스트림은
(대개 zlib raw deflate로 압축된) 레코드의 나열이다.

레코드 헤더는 UINT32 하나:
    bit  0-9  : tag id
    bit 10-19 : level (트리 깊이)
    bit 20-31 : data size. 0xFFF이면 바로 뒤 UINT32가 실제 크기.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field
from typing import Iterator

import olefile

# ---------------------------------------------------------------------------
# 단위
# ---------------------------------------------------------------------------
#: HWPUNIT = 1/7200 inch. DXA(twip) = 1/1440 inch.
HWPUNIT_PER_INCH = 7200
DXA_PER_INCH = 1440
EMU_PER_INCH = 914400


def hwp_to_dxa(v: int) -> int:
    """HWPUNIT -> twip(DXA). Word의 거의 모든 길이 단위."""
    return int(round(v / HWPUNIT_PER_INCH * DXA_PER_INCH))


def hwp_to_emu(v: int) -> int:
    """HWPUNIT -> EMU. 그림 크기에 쓴다."""
    return int(round(v / HWPUNIT_PER_INCH * EMU_PER_INCH))


def colorref(v: int) -> str:
    """HWP COLORREF(0x00BBGGRR) -> OOXML 'RRGGBB'.

    최상위 바이트가 0xFF면 '색 없음'을 뜻하므로 None을 돌려준다.
    """
    if v is None:
        return None
    if (v >> 24) & 0xFF == 0xFF:
        return None
    return "%02X%02X%02X" % (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)


# ---------------------------------------------------------------------------
# 레코드
# ---------------------------------------------------------------------------
@dataclass
class Record:
    tag: int
    level: int
    payload: bytes
    #: 이 레코드보다 level이 깊은 뒤따르는 레코드들 (parse_tree가 채운다)
    children: list["Record"] = field(default_factory=list)

    # --- payload 읽기 도우미 -------------------------------------------------
    def u8(self, off: int) -> int:
        return self.payload[off]

    def i8(self, off: int) -> int:
        return struct.unpack_from("<b", self.payload, off)[0]

    def u16(self, off: int) -> int:
        return struct.unpack_from("<H", self.payload, off)[0]

    def i16(self, off: int) -> int:
        return struct.unpack_from("<h", self.payload, off)[0]

    def u32(self, off: int) -> int:
        return struct.unpack_from("<I", self.payload, off)[0]

    def i32(self, off: int) -> int:
        return struct.unpack_from("<i", self.payload, off)[0]

    def has(self, off: int, n: int = 1) -> bool:
        """payload가 off부터 n바이트를 갖고 있는지. 버전에 따라 뒤가 잘린다."""
        return len(self.payload) >= off + n

    def u16_or(self, off: int, default: int = 0) -> int:
        return self.u16(off) if self.has(off, 2) else default

    def u32_or(self, off: int, default: int = 0) -> int:
        return self.u32(off) if self.has(off, 4) else default

    def i32_or(self, off: int, default: int = 0) -> int:
        return self.i32(off) if self.has(off, 4) else default


def iter_records(data: bytes) -> Iterator[Record]:
    """레코드 스트림을 평평하게 순회한다."""
    pos, end = 0, len(data)
    while pos + 4 <= end:
        (header,) = struct.unpack_from("<I", data, pos)
        pos += 4
        tag = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            if pos + 4 > end:
                break
            (size,) = struct.unpack_from("<I", data, pos)
            pos += 4
        if pos + size > end:
            size = end - pos
        yield Record(tag, level, data[pos : pos + size])
        pos += size


def parse_tree(records: list[Record]) -> list[Record]:
    """평평한 레코드 목록을 level에 따라 트리로 만든다.

    HWP는 중첩(표 안의 셀, 셀 안의 문단, 문단 안의 또 다른 표)을
    level 값으로만 표현하므로, 표를 제대로 다루려면 이 단계가 필요하다.
    """
    roots: list[Record] = []
    stack: list[Record] = []
    for rec in records:
        while stack and stack[-1].level >= rec.level:
            stack.pop()
        if stack:
            stack[-1].children.append(rec)
        else:
            roots.append(rec)
        stack.append(rec)
    return roots


def read_wstring(data: bytes, off: int) -> tuple[str, int]:
    """WORD 길이 + UTF-16LE 문자열. (문자열, 다음 오프셋)을 돌려준다."""
    if off + 2 > len(data):
        return "", off
    (n,) = struct.unpack_from("<H", data, off)
    off += 2
    raw = data[off : off + n * 2]
    return raw.decode("utf-16le", errors="replace"), off + n * 2


# ---------------------------------------------------------------------------
# 파일
# ---------------------------------------------------------------------------
class HwpError(Exception):
    """변환을 계속할 수 없는 경우."""


@dataclass
class FileHeader:
    signature: str
    version: tuple[int, int, int, int]
    compressed: bool
    password: bool
    distributable: bool
    has_script: bool
    drm: bool
    ccl: bool

    @property
    def version_str(self) -> str:
        return "%d.%d.%d.%d" % self.version

    def at_least(self, mm: int, nn: int, pp: int, rr: int = 0) -> bool:
        return self.version >= (mm, nn, pp, rr)


class HwpFile:
    """열린 .hwp 파일. 스트림을 압축 해제해서 돌려준다."""

    def __init__(self, path: str):
        self.path = path
        if not olefile.isOleFile(path):
            raise HwpError(
                "OLE 복합 파일이 아닙니다. HWP 3.x 이하이거나 .hwpx(ZIP/XML) "
                "또는 손상된 파일일 수 있습니다."
            )
        self.ole = olefile.OleFileIO(path)
        self.header = self._read_header()

        if self.header.password:
            raise HwpError("암호가 걸린 문서라 열 수 없습니다. 한글에서 암호를 푼 뒤 다시 시도하세요.")
        if self.header.drm:
            raise HwpError("DRM이 적용된 문서라 열 수 없습니다.")
        if self.header.distributable:
            raise HwpError(
                "배포용 문서(복사 방지)라 본문이 암호화되어 있습니다. "
                "한글에서 '배포용 문서 해제' 후 다시 저장한 파일을 사용하세요."
            )

    # -- 내부 --------------------------------------------------------------
    def _read_header(self) -> FileHeader:
        if not self.ole.exists("FileHeader"):
            raise HwpError("FileHeader 스트림이 없습니다. HWP 5.x 파일이 아닙니다.")
        raw = self.ole.openstream("FileHeader").read()
        sig = raw[:32].rstrip(b"\x00").decode("ascii", errors="replace")
        if not sig.startswith("HWP Document File"):
            raise HwpError(f"서명이 맞지 않습니다: {sig!r}")
        ver = struct.unpack_from("<I", raw, 32)[0]
        version = ((ver >> 24) & 0xFF, (ver >> 16) & 0xFF, (ver >> 8) & 0xFF, ver & 0xFF)
        flags = struct.unpack_from("<I", raw, 36)[0]
        return FileHeader(
            signature=sig,
            version=version,
            compressed=bool(flags & 0x01),
            password=bool(flags & 0x02),
            distributable=bool(flags & 0x04),
            has_script=bool(flags & 0x08),
            drm=bool(flags & 0x10),
            ccl=bool(flags & 0x80),
        )

    def _stream(self, name: str) -> bytes:
        raw = self.ole.openstream(name).read()
        if not self.header.compressed:
            return raw
        try:
            return zlib.decompress(raw, -15)
        except zlib.error:
            # 일부 파일은 zlib 헤더를 붙여 둔다.
            return zlib.decompress(raw)

    # -- 공개 --------------------------------------------------------------
    def doc_info(self) -> bytes:
        return self._stream("DocInfo")

    def section_names(self) -> list[str]:
        names = []
        for entry in self.ole.listdir():
            if len(entry) == 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
                names.append(entry[1])
        return sorted(names, key=lambda n: int(n[len("Section") :] or 0))

    def section(self, name: str) -> bytes:
        return self._stream(f"BodyText/{name}")

    def bin_data(self, bin_id: int, ext: str = "") -> bytes | None:
        """BinData/BIN####.ext 스트림. 없으면 None."""
        prefix = "BIN%04X" % bin_id
        for entry in self.ole.listdir():
            if len(entry) == 2 and entry[0] == "BinData" and entry[1].upper().startswith(prefix):
                raw = self.ole.openstream("/".join(entry)).read()
                if self.header.compressed:
                    try:
                        return zlib.decompress(raw, -15)
                    except zlib.error:
                        try:
                            return zlib.decompress(raw)
                        except zlib.error:
                            return raw
                return raw
        return None

    def preview_image(self) -> bytes | None:
        """미리보기 PNG. 변환 결과 검증(verify)에 쓴다."""
        if self.ole.exists("PrvImage"):
            return self.ole.openstream("PrvImage").read()
        return None

    def preview_text(self) -> str | None:
        if self.ole.exists("PrvText"):
            return self.ole.openstream("PrvText").read().decode("utf-16le", errors="replace")
        return None

    def close(self) -> None:
        self.ole.close()

    def __enter__(self) -> "HwpFile":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
