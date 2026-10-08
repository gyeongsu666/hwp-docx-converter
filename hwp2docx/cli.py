"""명령줄 인터페이스.

    hwp2docx 문서.hwp                     한 파일 변환
    hwp2docx 폴더/ -o 출력폴더/            폴더 통째로
    hwp2docx *.hwp --safe-fonts --check   글꼴 대체 + 결과 검증
"""

from __future__ import annotations

import argparse
import glob
import os
import sys

from .convert import convert


def _collect(inputs: list[str], recursive: bool) -> list[str]:
    paths: list[str] = []
    for item in inputs:
        if os.path.isdir(item):
            pattern = "**/*.hwp" if recursive else "*.hwp"
            paths.extend(glob.glob(os.path.join(glob.escape(item), pattern), recursive=recursive))
        elif not os.path.exists(item) and any(ch in item for ch in "*?["):
            paths.extend(glob.glob(item, recursive=recursive))
        else:
            paths.append(item)
    seen, unique = set(), []
    for path in paths:
        key = os.path.abspath(path)
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def _parse_font_map(values: list[str] | None) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for item in values or []:
        if "=" not in item:
            raise SystemExit(f"글꼴 지정 형식이 잘못됐습니다: {item!r} (예: 함초롬바탕=맑은 고딕)")
        old, new = item.split("=", 1)
        mapping[old.strip()] = new.strip()
    return mapping


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="hwp2docx",
        description="한글 문서(.hwp)를 Word 문서(.docx)로 변환합니다.",
    )
    parser.add_argument("inputs", nargs="+", metavar="입력",
                        help=".hwp 파일, 폴더, 또는 패턴")
    parser.add_argument("-o", "--output", metavar="경로",
                        help="출력 파일 또는 폴더 (기본: 원본 옆에 같은 이름)")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="폴더를 하위까지 훑습니다")
    parser.add_argument("--safe-fonts", action="store_true",
                        help="함초롬 등 한컴 전용 글꼴을 어디서나 있는 글꼴로 바꿉니다")
    parser.add_argument("--font", action="append", metavar="원본=대체",
                        help="글꼴을 직접 지정합니다. 여러 번 쓸 수 있습니다")
    parser.add_argument("--no-layout", action="store_true",
                        help="한글이 저장해 둔 배치 좌표를 쓰지 않고 문단 모양 값만 씁니다")
    parser.add_argument("--check", action="store_true",
                        help="변환 결과를 원본 미리보기와 대조해 오차를 알려 줍니다")
    parser.add_argument("-q", "--quiet", action="store_true", help="경고를 숨깁니다")
    args = parser.parse_args(argv)

    paths = _collect(args.inputs, args.recursive)
    if not paths:
        print("변환할 .hwp 파일을 찾지 못했습니다.", file=sys.stderr)
        return 1

    font_map = _parse_font_map(args.font)
    out_dir = None
    single_out = None
    if args.output:
        if len(paths) > 1 or os.path.isdir(args.output) or args.output.endswith(os.sep):
            out_dir = args.output
            os.makedirs(out_dir, exist_ok=True)
        else:
            single_out = args.output

    failures = 0
    for path in paths:
        if single_out:
            output = single_out
        elif out_dir:
            output = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".docx")
        else:
            output = None

        result = convert(path, output, safe_fonts=args.safe_fonts,
                         use_layout=not args.no_layout, font_map=font_map,
                         check=args.check)

        name = os.path.basename(path)
        if not result.ok:
            failures += 1
            print(f"✗ {name}: {result.error}", file=sys.stderr)
            continue

        print(f"✓ {name} → {os.path.basename(result.output)}")
        if result.verify_message:
            print(f"    검증: {result.verify_message}")
        if result.warnings and not args.quiet:
            for warning in result.warnings:
                print(f"    · {warning}")

    if failures:
        print(f"\n{len(paths) - failures}개 성공, {failures}개 실패", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
