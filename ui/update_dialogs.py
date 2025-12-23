from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

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


class ReleaseDetailDialog(QtWidgets.QDialog):
    def __init__(self, release, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Релиз {release.tag}")
        self.resize(640, 420)
        layout = QtWidgets.QVBoxLayout(self)

        header = QtWidgets.QLabel()
        parts = [f"<b>{release.tag}</b>"]
        if getattr(release, "published_at", None):
            parts.append(f"от {release.published_at}")
        if getattr(release, "prerelease", False):
            parts.append("(pre-release)")
        header.setText(" ".join(parts))
        layout.addWidget(header)

        desc = QtWidgets.QTextEdit()
        desc.setReadOnly(True)
        desc.setPlainText(release.body or "Описание отсутствует.")
        layout.addWidget(desc)

        btns = QtWidgets.QDialogButtonBox()
        btn_choose = btns.addButton("Выбрать", QtWidgets.QDialogButtonBox.AcceptRole)
        btns.addButton(QtWidgets.QDialogButtonBox.Close)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        # Disable choose when нет файла для платформы
        if not getattr(release, "asset", None) or not getattr(release.asset, "download_url", ""):
            btn_choose.setEnabled(False)


class ReleasePickerDialog(QtWidgets.QDialog):
    def __init__(self, releases: list, parent=None, current_version: str | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Выбор версии для установки")
        self.resize(640, 520)
        self._releases = releases
        self._current_version = current_version
        self.selected = None

        layout = QtWidgets.QVBoxLayout()
        self.setLayout(layout)

        if current_version:
            cv_label = QtWidgets.QLabel(f"Текущая версия: {current_version}")
            cv_label.setStyleSheet("font-weight: bold; color: #d35400;")
            layout.addWidget(cv_label)

        self.listw = QtWidgets.QListWidget()
        self.listw.itemDoubleClicked.connect(self._open_detail)
        self.listw.currentItemChanged.connect(self._on_selection_changed)
        layout.addWidget(self.listw)

        self.desc = QtWidgets.QTextEdit()
        self.desc.setReadOnly(True)
        self.desc.setFixedHeight(160)
        layout.addWidget(self.desc)

        info = QtWidgets.QLabel(
            "Выберите версию для установки. Бета-версии помечены как 'bN'.\n"
            "Если для вашей платформы нет файла в релизе — элемент будет недоступен."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.detail_btn = btns.addButton("Подробнее", QtWidgets.QDialogButtonBox.ActionRole)
        self.detail_btn.clicked.connect(self._open_detail)
        btns.accepted.connect(self._on_accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        self._populate()
        self._on_selection_changed(self.listw.currentItem(), None)

    def _populate(self):
        self.listw.clear()
        current_row = 0
        for r in self._releases:
            text = r.tag
            if getattr(r, "published_at", None):
                text += f"  —  {r.published_at}"
            if getattr(r, "prerelease", False):
                text += "  [pre-release]"
            is_current = self._is_current(r.tag)
            if is_current:
                text += "  —  установлена"
            item = QtWidgets.QListWidgetItem(text)
            if not getattr(r, "asset", None) or not getattr(r.asset, "download_url", ""):
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEnabled)
            if is_current:
                item.setBackground(QtCore.Qt.GlobalColor.yellow)
                item.setForeground(QtCore.Qt.GlobalColor.black)
                item.setFont(self._bold_font(item.font()))
            item.setData(QtCore.Qt.UserRole, r)
            self.listw.addItem(item)
            if is_current:
                current_row = self.listw.count() - 1

        if self.listw.count() > 0:
            self.listw.setCurrentRow(current_row)

    def _on_accept(self):
        item = self.listw.currentItem()
        if not item or not (item.flags() & QtCore.Qt.ItemIsEnabled):
            return
        self.selected = item.data(QtCore.Qt.UserRole)
        self.accept()

    def _on_selection_changed(self, current, previous):
        release = current.data(QtCore.Qt.UserRole) if current else None
        if release is None:
            self.desc.setPlainText("")
            self.detail_btn.setEnabled(False)
            return
        self.detail_btn.setEnabled(bool(current.flags() & QtCore.Qt.ItemIsEnabled))
        body = getattr(release, "body", "") or "Описание отсутствует."
        self.desc.setPlainText(body)
        self.desc.moveCursor(QtGui.QTextCursor.Start)

    def _open_detail(self, item=None):
        if item is None:
            item = self.listw.currentItem()
        if not item or not (item.flags() & QtCore.Qt.ItemIsEnabled):
            return
        release = item.data(QtCore.Qt.UserRole)
        dlg = ReleaseDetailDialog(release, parent=self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.selected = release
            self.accept()

    def _is_current(self, tag: str) -> bool:
        if not self._current_version:
            return False

        def _norm(v: str) -> str:
            return v.lower().lstrip("v").strip()

        return _norm(tag) == _norm(self._current_version)

    @staticmethod
    def _bold_font(font):
        font.setBold(True)
        return font
