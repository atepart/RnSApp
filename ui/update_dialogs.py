from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from infrastructure.updater import list_releases


class FetchReleasesWorker(QtCore.QObject):
    finished = QtCore.Signal(list)
    error = QtCore.Signal(str)
    status = QtCore.Signal(str)

    def __init__(self, repo_slug: str, limit: int = 10) -> None:
        super().__init__()
        self._repo = repo_slug
        self._limit = limit

    @QtCore.Slot()
    def run(self):
        try:
            self.status.emit("Подключение к GitHub API...")
            releases = list_releases(self._repo, self._limit)
            if not releases:
                # Считаем отсутствующий список релизов ошибкой, чтобы UI показал диалог
                self.error.emit("Не удалось получить список релизов")
                return
            avail = sum(1 for r in releases if getattr(r, "asset", None) and getattr(r.asset, "download_url", ""))
            self.status.emit(f"Получено релизов: {len(releases)} (доступно для вашей платформы: {avail})")
            self.finished.emit(releases)
        except Exception as e:
            self.error.emit(str(e))


class ReleasePickerDialog(QtWidgets.QDialog):
    def __init__(self, releases: list, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Выбор версии для установки")
        self.resize(560, 380)
        self._releases = releases
        self.selected = None

        layout = QtWidgets.QVBoxLayout()
        self.setLayout(layout)

        self.listw = QtWidgets.QListWidget()
        self.listw.itemDoubleClicked.connect(self._on_accept)
        layout.addWidget(self.listw)

        info = QtWidgets.QLabel(
            "Выберите версию для установки. Бета-версии помечены как 'bN'.\n"
            "Если для вашей платформы нет файла в релизе — элемент будет недоступен."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        self._populate()

    def _populate(self):
        self.listw.clear()
        for r in self._releases:
            text = r.tag
            if getattr(r, "published_at", None):
                text += f"  —  {r.published_at}"
            if getattr(r, "prerelease", False):
                text += "  [pre-release]"
            item = QtWidgets.QListWidgetItem(text)
            if not getattr(r, "asset", None) or not getattr(r.asset, "download_url", ""):
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEnabled)
            item.setData(QtCore.Qt.UserRole, r)
            self.listw.addItem(item)

        if self.listw.count() > 0:
            self.listw.setCurrentRow(0)

    def _on_accept(self):
        item = self.listw.currentItem()
        if not item or not (item.flags() & QtCore.Qt.ItemIsEnabled):
            return
        self.selected = item.data(QtCore.Qt.UserRole)
        self.accept()
