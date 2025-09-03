import sys

from PySide6.QtCore import QLocale
from PySide6.QtWidgets import QApplication

from infrastructure.repository_memory import InMemoryCellRepository
from ui.app import RnSApp


def configure_logger():
    import logging

    logging.basicConfig(level=logging.INFO)


if __name__ == "__main__":
    configure_logger()
    # Use dot as decimal separator globally
    QLocale.setDefault(QLocale(QLocale.C))
    app = QApplication(sys.argv)
    # No QSS theme applied for now
    # Composition root: wire dependencies here
    repo = InMemoryCellRepository()
    rns_app = RnSApp(repo=repo)
    rns_app.show()
    sys.exit(app.exec())
