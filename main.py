import sys

from PySide6.QtCore import QCoreApplication, QLocale
from PySide6.QtWidgets import QApplication

from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.xlsx_io import XlsxCellIO
from ui.app import RnSApp


def configure_logger():
    import logging
    import os

    # Console logging with level from env: RNS_LOG_LEVEL=DEBUG|INFO|WARNING
    level_name = os.getenv("RNS_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    datefmt = "%H:%M:%S"
    # Reset basicConfig by configuring root handlers explicitly
    root = logging.getLogger()
    root.setLevel(level)
    # Remove existing handlers to avoid duplicates
    for h in list(root.handlers):
        root.removeHandler(h)
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(fmt=fmt, datefmt=datefmt))
    root.addHandler(handler)
    # Tame noisy libs unless DEBUG requested
    if level > logging.DEBUG:
        logging.getLogger("urllib3").setLevel(logging.WARNING)
        logging.getLogger("requests").setLevel(logging.WARNING)


if __name__ == "__main__":
    configure_logger()
    # Configure settings scope and decimal separator
    QCoreApplication.setOrganizationName("AY_Inc")
    QCoreApplication.setApplicationName("RnSApp")
    # Use dot as decimal separator globally
    QLocale.setDefault(QLocale(QLocale.C))
    app = QApplication(sys.argv)
    # No QSS theme applied for now
    # Composition root: wire dependencies here
    repo = InMemoryCellRepository()
    excel = XlsxCellIO()
    rns_app = RnSApp(repo=repo, excel_io=excel)
    rns_app.show()
    sys.exit(app.exec())
