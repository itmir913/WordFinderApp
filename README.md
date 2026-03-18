# 금지어 탐지기

PDF 및 Excel(XLSX) 파일에서 금지어를 탐지하고, 결과를 시각적으로 표시하는 데스크톱 애플리케이션입니다.

---

## 프로젝트 구조

```
forbidden_word_detector/
├── main.py                        # 진입점
├── requirements.txt
├── ui/
│   ├── __init__.py
│   └── main_window.py             # PyQt5 메인 윈도우 (드래그 앤 드롭, 테이블, 로그)
├── workers/
│   ├── __init__.py
│   └── processor.py               # QThread + FileWorker (비동기 처리)
├── processors/
│   ├── __init__.py
│   ├── pdf_processor.py           # PyMuPDF 기반 PDF 처리
│   └── excel_processor.py         # pandas + openpyxl 기반 Excel 처리
└── utils/
    ├── __init__.py
    └── keyword_loader.py          # 금지어 CSV 로더
```

---

## 설치

```bash
pip install -r requirements.txt
```

---

## 실행

```bash
python main.py
```

---

## 사용법

1. **금지어 CSV 선택** — `금지어.csv` 파일을 선택하거나 드래그
   - A열에 금지어 목록 (UTF-8 또는 CP949 자동 감지)

2. **처리 파일 추가** — PDF/XLSX 파일을 버튼 또는 드래그로 추가

3. **처리 시작** — ▶ 버튼 클릭

4. **결과 확인**
   - PDF → `output_원본명.pdf` (하이라이트 + 북마크)
   - Excel → `output_원본명.xlsx` (탐지여부, 탐지된 금지어 컬럼 + 노란 행 강조)

---

## 클래스 설계

| 클래스 | 파일 | 역할 |
|--------|------|------|
| `MainWindow` | `ui/main_window.py` | PyQt5 메인 윈도우. 드래그 앤 드롭, 파일 테이블, 로그 출력, 버튼 제어 |
| `DropTableWidget` | `ui/main_window.py` | 드래그 앤 드롭을 지원하는 QTableWidget 서브클래스 |
| `FileWorker` | `workers/processor.py` | QObject 워커. 파일 목록을 순차 처리하고 signal로 UI에 상태 전달 |
| `ProcessorThread` | `workers/processor.py` | QThread 래퍼. Worker를 moveToThread로 이동시켜 실행 |
| `process_pdf()` | `processors/pdf_processor.py` | PyMuPDF로 PDF 탐지, 하이라이트, 북마크 저장 |
| `process_excel()` | `processors/excel_processor.py` | pandas 탐지 + openpyxl 조건부 서식 저장 |
| `load_keywords()` | `utils/keyword_loader.py` | CSV에서 금지어 set 로드 |

---

## 에러 처리 구조

| 에러 종류 | 처리 방식 |
|-----------|-----------|
| FileNotFoundError | 해당 파일 "실패" 처리 후 다음 파일 계속 |
| PermissionError | 저장 실패로 기록 후 계속 |
| UnicodeDecodeError | CSV 인코딩 자동 재시도 (UTF-8 → CP949) |
| 기타 Exception | 개별 파일 "실패" 처리, 전체 중단 없음 |

---

## 출력 형식

### PDF
- `output_파일명.pdf`
- 금지어 위치에 노란색 Highlight Annotation
- 북마크: `{페이지번호} - {금지어}` (flat 구조)

### Excel
- `output_파일명.xlsx`
- 새 컬럼 추가:
  - `탐지여부`: TRUE / FALSE
  - `탐지된 금지어`: 탐지된 금지어를 `, `로 연결
- 탐지된 행 전체 노란색 배경 (조건부 서식)
