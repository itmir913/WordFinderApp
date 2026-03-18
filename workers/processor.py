"""
멀티스레딩 처리 모듈
- QThread 서브클래싱 방식 (moveToThread 미사용 → C 레벨 크래시 방지)
- UI 프리징 없는 비동기 파일 처리
"""
import time
import traceback
from pathlib import Path
from typing import List, Set

from PyQt6.QtCore import QThread, pyqtSignal

from processors.excel_processor import process_excel
from processors.pdf_processor import process_pdf


class ProcessorThread(QThread):
    """
    QThread를 직접 서브클래싱하는 안전한 패턴.
    moveToThread + finished→quit 재진입 문제를 원천 차단한다.
    """

    # Signals (메인 스레드로 전달)
    progress = pyqtSignal(str, str)  # (파일명, 상태)
    log = pyqtSignal(str)  # 상세 로그
    error = pyqtSignal(str, str)  # (파일명, 에러 메시지)

    # finished 는 QThread 기본 제공 사용

    def __init__(self, file_paths: List[str], keywords: Set[str], parent=None):
        super().__init__(parent)
        self._file_paths = file_paths
        self._keywords = keywords
        self._stop_flag = False

    def stop(self):
        self._stop_flag = True

    # ── 핵심: QThread.run() 오버라이드 ──
    def run(self):
        try:
            self._process_all()
        except Exception as e:
            self.log.emit(f"❌ 치명적 오류: {e}\n{traceback.format_exc()}")

    def _process_all(self):
        for file_path in self._file_paths:
            if self._stop_flag:
                self.log.emit("⛔ 사용자에 의해 중지되었습니다.")
                break

            name = Path(file_path).name
            self.progress.emit(name, "처리중")
            self.log.emit(f"▶ 처리 시작: {file_path}")

            time.sleep(0.01)

            try:
                ext = Path(file_path).suffix.lower()

                if ext == ".pdf":
                    result = process_pdf(file_path, self._keywords)
                    self.progress.emit(name, "성공")
                    self.log.emit(
                        f"✅ PDF 완료 | 탐지 {result['total_found']}건 "
                        f"→ {result['output_path']}"
                    )

                elif ext == ".xlsx":
                    result = process_excel(file_path, self._keywords)
                    self.progress.emit(name, "성공")
                    self.log.emit(
                        f"✅ Excel 완료 | 전체 {result['total_rows']}행 중 "
                        f"탐지 {result['detected_rows']}행 "
                        f"→ {result['output_path']}"
                    )

                else:
                    self.progress.emit(name, "실패")
                    self.error.emit(name, f"지원하지 않는 파일 형식: {ext}")
                    self.log.emit(f"❌ 지원 안 함: {file_path}")

                time.sleep(0.01)

            except Exception as e:
                detail = traceback.format_exc()
                self.progress.emit(name, "실패")
                self.error.emit(name, str(e))
                self.log.emit(f"❌ [{name}] 상세 오류:\n{detail}")
