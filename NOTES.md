# 작업 기록

## hwp2docx (한글 → Word) — 완료, 0.1.1 배포
- PARA_LINE_SEG(한글이 저장한 줄 배치)로 세로 간격 재현. 실제 양식에서 오차 최대 0.9mm.
- 발견·수정한 버그 (0.1.2에 반영 예정)
  1. 바이너리 .hwp 문단 모양의 여백·들여쓰기·문단 간격은 실제 HWPUNIT의 2배로 저장됨
     (한글이 만든 .hwpx의 hp:case / hp:default 두 벌 비교로 확인) → 1/2로 읽도록 수정
  2. 한글 내어쓰기(첫 줄 = 왼쪽 여백, 나머지가 들어감) ≠ Word hanging → left에 |값| 더함
  3. 배치 파일 title 줄의 '>'가 빈 'Word' 파일을 만들던 문제 (0.1.1에서 수정)

## docx2hwpx (Word → 한글 .hwpx) — 0.2.0 배포
- 뼈대: python-hwpx(Apache-2.0)의 Skeleton.hwpx (한글 2024가 저장한 빈 문서)
- 한글이 여는 조건(python-hwpx 검사기 기준): secPr은 첫 문단 첫 run, 표에 sz/pos/outMargin/inMargin,
  셀에 subList/cellAddr/cellSpan/cellSz/cellMargin, linesegarray는 넣지 않음(한글이 캐시를 재사용해 겹침)
- 병합 셀은 가려진 칸을 빼 버림. 줄바꿈 <hp:t>a<hp:lineBreak/>b</hp:t>, 탭은 run 안 <hp:tab/>
- 문단 여백: hp:case = 실제값, hp:default = 2배
- Word 줄 간격 배수 → 한글 %: 배수 × (글꼴 winAscent+winDescent)/em × 100 (윈도우에선 실제 글꼴 파일 읽음)
- 번호 목록은 글자로 풀고 tabPr id=1(내어쓰기 자동 탭) 사용
- 미확인: 가로 용지 = landscape="NARROWLY" (세로는 WIDELY 확인됨)
- 실사용 파일 2개 변환: python-hwpx 한글 열림 검사 통과,
  문단 단위 대조(글자·서식·용지) 불일치 0. 대조 스크립트는 일부러 넣은 오류 5종을 모두 잡음.
- 남은 불확실성: 휴먼명조 줄 높이 비율(이 환경엔 글꼴 없음 → 1.2 근삿값, 줄 간격 240%)
