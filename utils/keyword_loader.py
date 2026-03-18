import csv
from pathlib import Path
from typing import Set


def load_keywords(filepath: str) -> Set[str]:
    """
    CSV 파일에서 금지어 목록을 로드한다.
    - UTF-8 우선 시도, 실패 시 CP949 재시도
    - trim, 빈 값 제거, 중복 제거
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    keywords: Set[str] = set()

    for encoding in ("utf-8-sig", "utf-8", "cp949"):
        try:
            with open(path, newline="", encoding=encoding) as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        word = row[0].strip()
                        if word:
                            keywords.add(word)
            return keywords
        except (UnicodeDecodeError, UnicodeError):
            continue

    raise ValueError(f"인코딩을 감지할 수 없습니다: {filepath}")
