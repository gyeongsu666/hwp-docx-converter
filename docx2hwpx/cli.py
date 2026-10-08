"""명령줄: python -m docx2hwpx.cli 문서.docx [폴더] [-o 출력]"""

from __future__ import annotations

import argparse
import glob
import os
import sys

from .convert import convert


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="docx2hwpx", description="Word 문서(.docx)를 한글 문서(.hwpx)로 변환합니다.")
    ap.add_argument("inputs", nargs="+", metavar="입력", help=".docx 파일, 폴더, 또는 패턴")
    ap.add_argument("-o", "--output", metavar="경로", help="출력 파일 또는 폴더 (기본: 원본 옆)")
    ap.add_argument("-r", "--recursive", action="store_true", help="폴더를 하위까지 훑습니다")
    ap.add_argument("-q", "--quiet", action="store_true", help="경고를 숨깁니다")
    args = ap.parse_args(argv)

    paths = []
    for item in args.inputs:
        if os.path.isdir(item):
            paths += glob.glob(os.path.join(glob.escape(item), "**/*.docx" if args.recursive else "*.docx"),
                               recursive=args.recursive)
        elif not os.path.exists(item) and any(c in item for c in "*?["):
            paths += glob.glob(item, recursive=args.recursive)
        else:
            paths.append(item)
    # Word가 열어 둔 파일의 잠금 파일(~$...)은 건너뛴다
    paths = [p for p in dict.fromkeys(paths) if not os.path.basename(p).startswith("~$")]
    if not paths:
        print("변환할 .docx 파일을 찾지 못했습니다.", file=sys.stderr)
        return 1

    out_dir = single = None
    if args.output:
        if len(paths) > 1 or os.path.isdir(args.output):
            out_dir = args.output
            os.makedirs(out_dir, exist_ok=True)
        else:
            single = args.output

    failures = 0
    for path in paths:
        out = single or (os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".hwpx")
                         if out_dir else None)
        res = convert(path, out)
        name = os.path.basename(path)
        if not res.ok:
            failures += 1
            print(f"✗ {name}: {res.error}", file=sys.stderr)
            continue
        print(f"✓ {name} → {os.path.basename(res.output)}")
        if not args.quiet:
            for wmsg in res.warnings:
                print(f"    · {wmsg}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
