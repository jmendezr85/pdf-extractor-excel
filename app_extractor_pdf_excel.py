
import sys
from PySide6 import QtWidgets

from ui import MainWindow, APP_TITLE


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
