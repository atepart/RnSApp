import sys

from PySide6.QtWidgets import QApplication

from src.app import RnSApp
from src.logger import configure_logger

if __name__ == "__main__":
    configure_logger()
    app = QApplication(sys.argv)
    rns_app = RnSApp()
    rns_app.show()
    sys.exit(app.exec())
