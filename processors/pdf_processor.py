"""
PDF 처리 모듈 (PyMuPDF 기반)
- 금지어 탐지 (page.search_for)
- Highlight Annotation 추가
- Flat 북마크 생성
"""

from pathlib import Path
from typing import Set, List, Tuple, Dict
import fitz  # PyMuPDF


# 하이라이트 색상 (RGBA 0~1)
HIGHLIGHT_COLOR = (1.0, 0.9, 0.0)  # 노란색


def process_pdf(input_path: str, keywords: Set[str]) -> Dict:
    """
    PDF 파일에서 금지어를 탐지하고 highlight + bookmark를 추가한 뒤 저장한다.

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

    output_path = src.parent / f"output_{src.name}"
    details: List[Tuple[int, str, int]] = []

    doc = fitz.open(str(src))
    toc: List[List] = []  # [level, title, page_number]

    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_no = page_index + 1  # 1-based

            for keyword in keywords:
                rects = page.search_for(keyword)
                if not rects:
                    continue

                for rect in rects:
                    # Highlight annotation 추가
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=HIGHLIGHT_COLOR)
                    annot.update()

                    # 개별 북마크 생성
                    toc.append([1, f"{page_no} - {keyword}", page_no])
                    details.append((page_no, keyword, 1))

        # 북마크 적용 (flat 구조)
        if toc:
            doc.set_toc(toc)

        doc.save(str(output_path), garbage=4, deflate=True)

    finally:
        doc.close()

    return {
        "output_path": str(output_path),
        "total_found": len(details),
        "details": details,
    }
