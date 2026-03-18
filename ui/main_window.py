"""
메인 윈도우 (PyQt6) - Frameless Custom UI 버전
"""

from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QTextEdit, QGroupBox, QLineEdit,
    QMessageBox, QAbstractItemView, QTabWidget, QApplication
)

from utils.keyword_loader import load_keywords
from workers.processor import ProcessorThread

# 칼럼 인덱스 변경
COL_NAME = 0
COL_STATUS = 1
COL_DELETE = 2

STATUS_COLORS = {
    "대기": "#FFFFFF",
    "처리중": "#FFF3CD",
    "성공": "#D4EDDA",
    "실패": "#F8D7DA",
}


# ──────────────────────────────────────────────
# 커스텀 타이틀 바 (개선된 디자인)
# ──────────────────────────────────────────────
class CustomTitleBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)  # 높이 살짝 축소
        self.start_pos = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 0, 0)  # 우측 여백 0으로 (버튼 밀착)
        layout.setSpacing(0)

        # 타이틀 텍스트
        title = QLabel("🔍 학교생활기록부 일괄 점검 프로그램 v2026.03.18")
        title.setStyleSheet(
            "color: #343A40; font-size: 16px; font-weight: bold; font-family: 'Malgun Gothic';")
        layout.addWidget(title)

        layout.addStretch()

        # 💡 제작자 정보 버튼 추가
        self.info_btn = QPushButton("ⓘ")
        self.info_btn.setFixedSize(45, 40)
        self.info_btn.setStyleSheet("""
        QPushButton {
            border: none; background: transparent; color: #6C757D; font-size: 18px;
        }
        QPushButton:hover { background-color: #E9ECEF; color: #343A40; }
        """)
        self.info_btn.clicked.connect(self.parent.show_about_dialog)  # 메인 윈도우의 함수 호출
        layout.addWidget(self.info_btn)

        # 최소화 버튼 (디자인 개선)
        btn_min = QPushButton("─")
        btn_min.setFixedSize(45, 40)
        btn_min.clicked.connect(self.parent.showMinimized)
        btn_min.setStyleSheet("""
            QPushButton { border: none; background: transparent; font-size: 14px; color: #495057; }
            QPushButton:hover { background: #E9ECEF; color: #212529; }
        """)

        # 닫기 버튼 (디자인 개선)
        btn_close = QPushButton("✕")
        btn_close.setFixedSize(45, 40)
        btn_close.clicked.connect(self.parent.close)
        btn_close.setStyleSheet("""
            QPushButton { border: none; background: transparent; font-size: 16px; color: #495057; }
            QPushButton:hover { background: #E81123; color: white; }
        """)

        layout.addWidget(btn_min)
        layout.addWidget(btn_close)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.start_pos:
            delta = event.globalPosition().toPoint() - self.start_pos
            self.parent.move(self.parent.pos() + delta)
            self.start_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.start_pos = None


# ──────────────────────────────────────────────
# 드래그 앤 드롭 지원 파일 리스트 테이블
# ──────────────────────────────────────────────
class DropTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls]
        main = self.window()

        if hasattr(main, "handle_dropped_files"):
            main.handle_dropped_files(paths)

        event.acceptProposedAction()


# ──────────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(800, 600)
        self._keywords: set = set()
        self._thread: Optional[ProcessorThread] = None

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Window)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self._init_ui()
        self._apply_styles()

    def _init_ui(self):
        central = QWidget()
        central.setObjectName("mainContainer")
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setSpacing(0)
        root.setContentsMargins(2, 2, 2, 10)

        self.title_bar = CustomTitleBar(self)
        root.addWidget(self.title_bar)

        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(15, 5, 15, 5)
        content_layout.setSpacing(10)

        # ── Step 1: 검색 기준 (CSV) ──
        content_layout.addWidget(self._build_csv_section())

        # ── Step 2: 검색 대상 (테이블 + 관리 버튼) ──
        # 탭 영역 내부에 관리 버튼을 포함시키거나, 탭 바로 위에 배치합니다.
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_file_management_page(), "📁 검색 대상 파일")
        self.tabs.addTab(self._build_log_panel(), "📜 시스템 로그")
        content_layout.addWidget(self.tabs, stretch=1)

        # ── Step 3: 실행 액션 ──
        content_layout.addWidget(self._build_action_buttons())

        root.addLayout(content_layout)

    def _build_csv_section(self) -> QGroupBox:
        box = QGroupBox("Step 1. 검색 기준 설정 (단어 목록)")
        layout = QVBoxLayout(box)

        row = QHBoxLayout()
        row.addWidget(QLabel("단어 목록 CSV:"))

        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("검색할 단어들이 담긴 CSV 파일을 등록하세요")
        self.csv_path_edit.setReadOnly(True)
        row.addWidget(self.csv_path_edit, stretch=1)

        self.btn_csv = QPushButton("CSV 파일 불러오기")
        self.btn_csv.clicked.connect(self._select_csv)
        row.addWidget(self.btn_csv)

        layout.addLayout(row)

        # 상태 표시를 우측 하단에 작게 배치
        status_layout = QHBoxLayout()
        status_layout.addStretch()
        self.kw_label = QLabel("검색 단어 미등록")
        self.kw_label.setStyleSheet("color: #E03131; font-size: 14px; font-weight: bold;")
        status_layout.addWidget(self.kw_label)
        layout.addLayout(status_layout)

        return box

    def _build_file_management_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)

        # 테이블 전용 툴바 (추가/비우기 버튼)
        tool_bar = QHBoxLayout()
        tool_bar.addWidget(QLabel("검사 대상 파일(*.pdf, *.xlsx)을 드래그해서 추가하세요"))
        tool_bar.addStretch()

        self.btn_add = QPushButton("➕ 파일 추가")
        self.btn_add.setObjectName("primaryBtn")
        self.btn_add.clicked.connect(self._select_files)
        tool_bar.addWidget(self.btn_add)

        self.btn_clear = QPushButton("🗑 전체 비우기")
        self.btn_clear.clicked.connect(self._clear_files)
        tool_bar.addWidget(self.btn_clear)

        layout.addLayout(tool_bar)

        # 테이블 배치
        self.table = DropTableWidget()
        self.table.setColumnCount(3)  # 3열로 변경
        self.table.setHorizontalHeaderLabels(["파일명", "상태", "삭제"])
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        # 비율 및 동작 설정
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(COL_NAME, QHeaderView.ResizeMode.Stretch)  # 8 비율 차지
        header.setSectionResizeMode(COL_STATUS, QHeaderView.ResizeMode.Fixed)  # 고정 너비
        header.setSectionResizeMode(COL_DELETE, QHeaderView.ResizeMode.Fixed)  # 고정 너비

        self.table.setColumnWidth(COL_STATUS, 80)
        self.table.setColumnWidth(COL_DELETE, 60)

        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setShowGrid(False)

        layout.addWidget(self.table)
        return page

    def _build_control_panel(self) -> QGroupBox:
        # 그룹 이름부터 목적을 명확히 함
        box = QGroupBox("검색 기준 설정")
        main_layout = QVBoxLayout(box)

        # ── 첫 번째 줄: 검색어 소스 ──
        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("단어 목록(CSV):"))

        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("검색 단어들이 나열된 CSV 파일을 선택하세요")
        self.csv_path_edit.setReadOnly(True)
        top_row.addWidget(self.csv_path_edit)

        self.btn_csv = QPushButton("불러오기")  # '찾기'보다 데이터 유입 느낌이 강함
        top_row.addWidget(self.btn_csv)
        main_layout.addLayout(top_row)

        # ── 두 번째 줄: 액션 ──
        bottom_row = QHBoxLayout()

        self.btn_add = QPushButton("검색 대상 파일 추가")  # PDF/Excel임을 암시
        self.btn_add.setObjectName("primaryBtn")
        bottom_row.addWidget(self.btn_add)

        self.btn_clear = QPushButton("목록 비우기")  # '전체'보다 간결
        bottom_row.addWidget(self.btn_clear)

        bottom_row.addStretch()

        # 상태 라벨: '로딩 대기중'보다 구체적으로
        self.kw_label = QLabel("검색 단어 파일 미등록")
        self.kw_label.setStyleSheet("color: #E03131; font-weight: bold;")  # 초기엔 경고색(빨강)
        bottom_row.addWidget(self.kw_label)

        main_layout.addLayout(bottom_row)
        return box

    def _build_file_table(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 5, 0, 0)

        self.table = DropTableWidget()
        self.table.setColumnCount(3)  # 3열로 변경
        self.table.setHorizontalHeaderLabels(["파일명", "상태", "삭제"])
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        # 비율 및 동작 설정
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(COL_NAME, QHeaderView.ResizeMode.Stretch)  # 8 비율 차지
        header.setSectionResizeMode(COL_STATUS, QHeaderView.ResizeMode.Fixed)  # 고정 너비
        header.setSectionResizeMode(COL_DELETE, QHeaderView.ResizeMode.Fixed)  # 고정 너비

        self.table.setColumnWidth(COL_STATUS, 80)
        self.table.setColumnWidth(COL_DELETE, 60)

        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setShowGrid(False)

        layout.addWidget(self.table)
        return widget

    def _build_log_panel(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 5, 0, 0)
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setFont(QFont("Consolas", 11))
        layout.addWidget(self.log_area)
        return widget

    def _build_action_buttons(self) -> QWidget:
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addStretch()

        self.btn_start = QPushButton("▶ 처리 시작")
        self.btn_start.setFixedHeight(40)
        self.btn_start.setStyleSheet(
            "QPushButton { background: #20C997; color: white; font-weight: bold; "
            "border-radius: 6px; padding: 0 30px; } "
            "QPushButton:hover { background: #12B886; }"
        )
        self.btn_start.clicked.connect(self._start_processing)
        layout.addWidget(self.btn_start)

        self.btn_stop = QPushButton("⛔ 중지")
        self.btn_stop.setFixedHeight(40)
        self.btn_stop.setEnabled(False)
        self.btn_stop.setStyleSheet(
            "QPushButton { background: #FA5252; color: white; font-weight: bold; "
            "border-radius: 6px; padding: 0 30px; } "
            "QPushButton:disabled { background: #FFE3E3; color: #FFA8A8; }"
            "QPushButton:hover:!disabled { background: #F03E3E; }"
        )
        self.btn_stop.clicked.connect(self._stop_processing)
        layout.addWidget(self.btn_stop)

        return widget

    def _apply_styles(self):
        # 폰트 변경 (맑은 고딕 우선), 크기 14px로 가독성 향상
        style = """
        #mainContainer { background-color: #F8F9FA; border: 1px solid #CED4DA; border-radius: 10px; }
        QWidget { font-family: 'Malgun Gothic', 'Segoe UI', sans-serif; font-size: 14px; color: #333333; }
        
        QGroupBox { 
            background-color: #FFFFFF; border: 1px solid #DEE2E6; 
            border-radius: 8px; margin-top: 15px; font-weight: bold; 
        }
        QGroupBox::title { 
            subcontrol-origin: margin; subcontrol-position: top left; 
            padding: 0 8px; left: 10px; top: 2px; color: #495057; 
        }
        
        QLineEdit { border: 1px solid #CED4DA; border-radius: 4px; padding: 5px; background-color: #F8F9FA; }
        QLineEdit:focus { border: 1px solid #4DABF7; background-color: #FFFFFF; }
        
        QPushButton { background-color: #E9ECEF; color: #495057; border: none; border-radius: 4px; padding: 6px 12px; font-weight: bold; }
        QPushButton:hover { background-color: #DEE2E6; }
        QPushButton:pressed { background-color: #CED4DA; }
        QPushButton#primaryBtn { background-color: #4DABF7; color: white; }
        QPushButton#primaryBtn:hover { background-color: #339AF0; }
        
        QTableWidget { 
            border: 1px solid #E9ECEF; 
            border-radius: 6px; 
            background-color: #FFFFFF; 
            alternate-background-color: #F8F9FA;
            selection-background-color: #E7F1FF; 
            selection-color: #000000; 
        }
        QHeaderView::section { background-color: #F1F3F5; color: #495057; border: none; border-bottom: 2px solid #DEE2E6; padding: 8px; font-weight: bold; }
        
        QTextEdit { border: 1px solid #E9ECEF; background-color: #212529; color: #C1C9D2; border-radius: 6px; padding: 10px; }
        
        QTabWidget::pane { border: 1px solid #E9ECEF; border-radius: 6px; background-color: #FFFFFF; top: -1px; }
        QTabBar::tab { background: #F8F9FA; color: #868E96; padding: 8px 25px; border: 1px solid #E9ECEF; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; font-weight: bold; }
        QTabBar::tab:selected { background: #FFFFFF; color: #212529; border-bottom: 1px solid #FFFFFF; }
        QTabBar::tab:hover:!selected { background: #F1F3F5; }
        /* =========================================================
           여기서부터 새로 추가할 부분 (QMessageBox 및 스크롤바)
           ========================================================= */

        /* 1. 알림창(QMessageBox) 스타일 */
        QMessageBox { 
            background-color: #FFFFFF; 
            border: 1px solid #CED4DA; 
        }
        QMessageBox QLabel { 
            color: #212529; /* 글씨 색상을 선명하게 */
        }
        QMessageBox QPushButton { 
            background-color: #F1F3F5; 
            color: #495057; 
            border-radius: 4px; 
            padding: 6px 20px; 
            min-width: 60px; 
        }
        QMessageBox QPushButton:hover { 
            background-color: #E9ECEF; 
        }

        /* 2. 스크롤바(QScrollBar) 스타일 - 모던하게 변경 */
        QScrollBar:vertical { 
            border: none; background: #F8F9FA; width: 10px; border-radius: 5px; margin: 0px; 
        }
        QScrollBar::handle:vertical { 
            background: #CED4DA; min-height: 20px; border-radius: 5px; 
        }
        QScrollBar::handle:vertical:hover { 
            background: #ADB5BD; 
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { 
            height: 0px; /* 위아래 화살표 버튼 숨김 */
        }
        
        QScrollBar:horizontal { 
            border: none; background: #F8F9FA; height: 10px; border-radius: 5px; margin: 0px; 
        }
        QScrollBar::handle:horizontal { 
            background: #CED4DA; min-width: 20px; border-radius: 5px; 
        }
        QScrollBar::handle:horizontal:hover { 
            background: #ADB5BD; 
        }
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { 
            width: 0px; /* 좌우 화살표 버튼 숨김 */
        }
        """
        self.setStyleSheet(style)

    def _select_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "CSV 선택", "", "CSV 파일 (*.csv)")
        if path: self._load_csv(path)

    def _load_csv(self, path: str):
        try:
            self._keywords = load_keywords(path)
            self.csv_path_edit.setText(path)
            self.kw_label.setText(f"검사할 단어 수: {len(self._keywords)}개")
            self._log(f"✅ CSV 로드 완료: {len(self._keywords)}개 | {Path(path).name}")
        except Exception as e:
            QMessageBox.critical(self, "CSV 로드 실패", str(e))
            self._log(f"❌ CSV 로드 실패: {e}")

    def _select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "처리 파일 선택", "", "지원 파일 (*.pdf *.xlsx);;PDF (*.pdf);;Excel (*.xlsx)"
        )
        self.add_files(paths)

    def add_files(self, paths: list):
        existing = self._existing_paths()
        added = 0
        for path in paths:
            p = Path(path)
            if p.suffix.lower() not in (".pdf", ".xlsx"): continue
            if str(p) in existing: continue

            row = self.table.rowCount()
            self.table.insertRow(row)

            # 1. 파일명 (내부에 실제 경로 숨겨둠)
            name_item = QTableWidgetItem(p.name)
            name_item.setData(Qt.ItemDataRole.UserRole, str(p))  # 경로 은닉
            self.table.setItem(row, COL_NAME, name_item)

            # 2. 상태
            status_item = QTableWidgetItem("대기")
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, COL_STATUS, status_item)

            # 3. 삭제 버튼
            del_btn = QPushButton("✖")
            del_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            del_btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent; 
                    border: none;
                    color: #ADB5BD; font-size: 14px; border-radius: 4px; 
                }
                QPushButton:hover { background-color: #FFE3E3; color: #FA5252; }
                QPushButton:pressed { background-color: #FFC9C9; color: #E03131; }
            """)
            del_btn.clicked.connect(lambda checked, b=del_btn: self._remove_row_by_button(b))

            btn_container = QWidget()
            btn_container.setStyleSheet("background-color: transparent;")
            btn_layout = QHBoxLayout(btn_container)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)  # 👈 2. 버튼 중앙 정렬 (추가)
            btn_layout.addWidget(del_btn)

            # 💡 버튼 뒤에 빈 도화지 깔기 (이게 있어야 교차 배경색이 칠해집니다)
            self.table.setItem(row, COL_DELETE, QTableWidgetItem(""))

            # 그 다음 버튼 배치
            self.table.setCellWidget(row, COL_DELETE, btn_container)
            self._set_row_color(row, "대기")
            added += 1

        if added: self._log(f"📂 {added}개 파일 추가됨.")

    def _remove_row_by_button(self, button: QPushButton):
        # 버튼의 위치를 추적하여 해당 행 삭제
        pos = button.parentWidget().pos()  # btn_container의 위치
        index = self.table.indexAt(pos)
        if index.isValid():
            filename = self.table.item(index.row(), COL_NAME).text()
            self.table.removeRow(index.row())
            self._log(f"🗑 '{filename}' 파일은 목록에서 제외되었습니다.")

    def _existing_paths(self) -> set:
        paths = set()
        for r in range(self.table.rowCount()):
            item = self.table.item(r, COL_NAME)
            if item:
                # 숨겨둔 경로(UserRole) 데이터 꺼내기
                paths.add(item.data(Qt.ItemDataRole.UserRole))
        return paths

    def _clear_files(self):
        self.table.setRowCount(0)
        self._log("🗑 파일 목록 전체 초기화.")

    def _start_processing(self):
        if not self._keywords:
            QMessageBox.warning(self, "CSV 파일이 선택되지 않음", "검색할 단어가 담긴 CSV를 먼저 입력하세요.")
            return
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "파일 없음", "처리할 PDF, Excel 파일을 추가하세요.")
            return

        # 숨겨둔 경로 데이터 리스트화
        file_paths = []
        for r in range(self.table.rowCount()):
            path = self.table.item(r, COL_NAME).data(Qt.ItemDataRole.UserRole)
            file_paths.append(path)
            self._update_row(r, "대기")

        # self.tabs.setCurrentIndex(1)  # 시작 시 로그 탭으로 자동 이동 (선택사항)
        # self.log_area.clear()
        self._log("🚀 처리 시작")

        self._thread = ProcessorThread(file_paths, self._keywords)
        self._thread.progress.connect(self._on_progress)
        self._thread.log.connect(self._log)
        self._thread.error.connect(self._on_error)
        self._thread.finished.connect(self._on_finished)

        # 1️⃣ 메인 버튼 제어
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)

        # 2️⃣ 상단 설정 버튼들 비활성화
        self.btn_csv.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.btn_clear.setEnabled(False)

        # 3️⃣ 드래그 앤 드롭 완전 차단
        self.setAcceptDrops(False)
        self.table.setAcceptDrops(False)

        # 4️⃣ 테이블 내부 요소 제어 (스크롤은 가능)
        self._set_table_buttons_enabled(False)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)

        self._thread.start()

    def _stop_processing(self):
        if self._thread:
            self._thread.stop()
            self._log("⛔ 중지 요청...")

    def _on_progress(self, filename: str, status: str):
        row = self._find_row(filename)
        if row is not None: self._update_row(row, status)

    def _on_error(self, filename: str, message: str):
        row = self._find_row(filename)
        if row is not None: self._update_row(row, "실패")
        self._log(f"❌ 오류 [{filename}]: {message}")

    def _on_finished(self):
        # 1️⃣ 메인 버튼 복구
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)

        # 2️⃣ 상단 설정 버튼들 활성화
        self.btn_csv.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_clear.setEnabled(True)

        # 3️⃣ 드래그 앤 드롭 다시 허용
        self.setAcceptDrops(True)
        self.table.setAcceptDrops(True)

        # 4️⃣ 테이블 내부 요소 복구
        self._set_table_buttons_enabled(True)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        self._log("✅ 모든 처리 완료.")
        self._thread = None

    def _find_row(self, filename: str) -> Optional[int]:
        # 파일 경로에서 파일명 추출해서 비교
        for r in range(self.table.rowCount()):
            item = self.table.item(r, COL_NAME)
            if item and item.text() == filename:
                return r
        return None

    def _update_row(self, row: int, status: str):
        item = self.table.item(row, COL_STATUS)
        if item:
            item.setText(status)
        self._set_row_color(row, status)
        QApplication.processEvents()

    def _set_row_color(self, row: int, status: str):
        if status == "대기":
            # 💡 '대기' 상태일 때는 강제로 색을 칠하지 않고 기본 교차 색상(흰색/연회색)이 보이게 둠
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setData(Qt.ItemDataRole.BackgroundRole, None)  # 배경색 초기화
        else:
            # 처리중, 성공, 실패 등의 상태일 때만 해당 색상으로 덮어씀
            color = QColor(STATUS_COLORS.get(status, "#FFFFFF"))
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(color)

    def _set_table_buttons_enabled(self, enabled: bool):
        """테이블 내의 모든 삭제 버튼을 한 번에 제어합니다."""
        for row in range(self.table.rowCount()):
            # 삭제 버튼이 들어있는 2번 열(COL_DELETE)의 위젯을 가져옴
            container = self.table.cellWidget(row, COL_DELETE)
            if container:
                # 컨테이너 안에 담긴 QPushButton을 찾아서 상태 변경
                btn = container.findChild(QPushButton)
                if btn:
                    btn.setEnabled(enabled)

    def _log(self, message: str):
        self.log_area.append(message)
        QApplication.processEvents()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls(): event.acceptProposedAction()

    def handle_dropped_files(self, paths: list):
        csv_paths = [p for p in paths if p.lower().endswith(".csv")]
        target_paths = [p for p in paths if p.lower().endswith((".pdf", ".xlsx"))]

        # 1. CSV 처리: 정확히 1개일 때만 자동 로드
        if len(csv_paths) == 1:
            self._load_csv(csv_paths[0])
        elif len(csv_paths) > 1:
            self._log("⚠️ 여러 개의 CSV가 감지되었습니다. 기준 파일은 하나씩만 등록 가능합니다.")
            QMessageBox.information(self, "알림", "CSV 파일은 한 번에 하나만 등록할 수 있습니다.")

        # 2. PDF/Excel 처리: 목록에 추가
        if target_paths:
            self.add_files(target_paths)

    # 윈도우 빈 공간에 드롭했을 때 처리
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls]
        self.handle_dropped_files(paths)
        event.acceptProposedAction()

    def show_about_dialog(self):
        """제작자 정보를 보여주는 팝업창 (여백 및 스타일 강화)"""
        about_text = (
            "<div style='margin: 10px;'>"
            "<h2 style='color: #2C3E50;'>학교생활기록부 일괄 점검 프로그램</h2>"
            "<hr>"
            "<p style='line-height: 150%;'>"
            "<b>버전:</b> v2026.03.16<br>"
            "<b>제작자:</b> 운양고등학교 이종환T<br>"
            "<b>GitHub:</b> <a href='https://github.com/itmir913' style='color: #3498DB;'>방문하기</a>"
            "</p>"
            "<br>"
            "<p style='line-height: 150%;'>"
            "본 프로그램은 <b>GNU LESSER GENERAL PUBLIC LICENSE Version 3 (LGPLv3)</b> 하에 배포됩니다.<br>"
            "누구나 자유롭게 사용, 복제 및 배포할 수 있으며,<br>"
            "수정하여 배포할 경우 해당 수정 사항의 소스코드는 LGPLv3에 따라 공개해야 합니다."
            "</p>"
            "<br>"
            "<p style='color: #7F8C8D;'><small>Copyright 2026. Licensed under LGPLv3.</small></p>"
            "</div>"
        )

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("프로그램 정보")
        msg_box.setText(about_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)

        # 💡 스타일시트로 최소 너비와 내부 여백 지정
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
            }
            QLabel {
                min-width: 400px;  /* 창의 최소 너비 설정 */
                padding: 20px;     /* 텍스트 주변 여백 추가 */
            }
            QPushButton {
                min-width: 80px;
                padding: 5px;
                margin-right: 10px;
                margin-bottom: 10px;
            }
        """)

        # 하이퍼링크 활성화
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.exec()
