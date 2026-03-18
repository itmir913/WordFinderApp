import sys
from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("학교생활기록부 일괄 점검 프로그램")
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    # PyQt6에서는 exec_() 대신 언더바가 빠진 exec()를 사용합니다!
    sys.exit(app.exec())

if __name__ == "__main__":
    main()