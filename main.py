import sys

from PySide6.QtCore import QCoreApplication, QLocale
from PySide6.QtWidgets import QApplication

from infrastructure.repository_memory import InMemoryCellRepository
from infrastructure.xlsx_io import XlsxCellIO
from ui.app import RnSApp


def configure_logger():
    import logging

    logging.basicConfig(level=logging.INFO)


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
