"""
PDF 처리 모듈 (PyMuPDF 기반)
- 금지어 탐지 (page.search_for) 및 플래그 적용
- Highlight Annotation 추가 (Quad 방식)
- 페이지별 중복 방지 북마크 생성
"""

from pathlib import Path
from typing import Set, List, Tuple, Dict

import fitz  # PyMuPDF

# 하이라이트 색상 (RGB 0~1)
HIGHLIGHT_COLOR = (1.0, 0.9, 0.0)  # 노란색


def process_pdf(input_path: str, keywords: Set[str]) -> Dict:
    """
    PDF 파일에서 단어를 검색하고 highlight + bookmark를 추가한 뒤 저장한다.

    Returns:
        {
            "output_path": str,
            "total_found": int,
            "details": [(page_no, keyword, rect_count), ...]
        }
    """
    src = Path(input_path)
    if not src.exists():
        raise FileNotFoundError(f"파일 없음: {input_path}")

    if not keywords:
        raise ValueError("검색 대상 단어 목록이 비어 있습니다.")

    output_path = src.parent / f"output_{src.name}"
    details: List[Tuple[int, str, int]] = []

    doc = fitz.open(str(src))

    # 💡 원본 PDF의 북마크(TOC)를 유지하고 싶다면 아래와 같이 시작합니다.
    # toc = doc.get_toc()
    toc: List[List] = []  # [level, title, page_number]

    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_no = page_index + 1  # 1-based (사용자 표시용)

            # 💡 한 페이지 내에서 어떤 단어들이 발견되었는지 추적 (북마크 중복 방지용)
            found_on_this_page = set()

            for keyword in keywords:
                # 💡 검색 플래그 설정
                # TEXT_PRESERVE_WHITESPACE: 공백 유지로 검색 정확도 향상
                # TEXT_DEHYPHENATE: 줄바꿈 시 하이픈으로 잘린 단어 결합 처리
                search_flags = fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_DEHYPHENATE | fitz.TEXT_IGNORE_CASE

                # 💡 quad=True: 사각형(Rect) 대신 사동형(Quad) 좌표를 반환받음
                quads = page.search_for(keyword, flags=search_flags, quads=True)

                if not quads:
                    continue

                # 1. 하이라이트 추가 (발견된 모든 위치)
                for q in quads:
                    # quad=True일 때는 add_highlight_annot에 quad 인자를 그대로 사용
                    annot = page.add_highlight_annot(q)
                    if annot:
                        annot.set_colors(stroke=HIGHLIGHT_COLOR)
                        annot.update()

                    # 상세 정보 기록 (전체 개수 카운트용)
                    details.append((page_no, keyword, 1))

                # 2. 북마크 추가 (페이지당 단어별로 1개씩만 추가하여 깔끔하게 유지)
                if keyword not in found_on_this_page:
                    toc.append([1, f"P{page_no}: {keyword}", page_no])
                    found_on_this_page.add(keyword)

        # 💡 생성된 모든 북마크를 PDF에 적용
        if toc:
            doc.set_toc(toc)

        # 저장 (최적화 옵션 포함)
        doc.save(str(output_path), garbage=4, deflate=True)

    finally:
        doc.close()

    return {
        "output_path": str(output_path),
        "total_found": len(details),
        "details": details,
    }
