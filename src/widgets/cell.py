from PySide6 import QtCore, QtGui, QtWidgets

from src.constants import ParamTableColumns
from src.store import Store


class CellWidget(QtWidgets.QGroupBox):
    def __init__(self, parent, index: int, param_table) -> None:
        super().__init__(parent)
        self.index = index
        self.param_table = param_table
        self.initUI()

    def initUI(self):
        layout = QtWidgets.QVBoxLayout()
        hlayout1 = QtWidgets.QHBoxLayout()
        hlayout2 = QtWidgets.QHBoxLayout()

        self.number = QtWidgets.QLabel(self)
        self.number.setText(f"№{self.index}")
        hlayout1.addWidget(self.number)

        self.name = QtWidgets.QLabel(self)
        hlayout1.addWidget(self.name)

        self.drift = QtWidgets.QLabel(self)
        self.drift.setToolTip("Уход")
        hlayout2.addWidget(self.drift)
        self.drift.setVisible(False)

        self.rns = QtWidgets.QLabel(self)
        self.rns.setToolTip("RnS")
        hlayout2.addWidget(self.rns)
        self.rns.setVisible(False)

        layout.addLayout(hlayout1)
        layout.addLayout(hlayout2)

        self.writeButton = QtWidgets.QPushButton("Записать")
        self.writeButton.clicked.connect(self.openWriteDialog)
        layout.addWidget(self.writeButton)

        self.checkbox = QtWidgets.QCheckBox("Построить")
        self.checkbox.setVisible(False)
        self.checkbox.stateChanged.connect(self.buildGraph)
        layout.addWidget(self.checkbox)

        self.setLayout(layout)

        self.customContextMenuRequested.connect(self.showContextMenu)

    def openWriteDialog(self):
        if self.param_table.is_empty():
            QtWidgets.QMessageBox.warning(
                self,
                "Нет рассчитанных данных!",
                "Сперва нужно занести данные и сделать расчет, после записывать!",
            )
            return
        name, ok = QtWidgets.QInputDialog.getText(self, "Запись", "Введите уникальное имя:")
        if ok and name:
            if Store.data.exclude(cell=self.index).filter(name=name).exists():
                QtWidgets.QMessageBox.warning(self, "Ошибка", "Это имя уже существует. Пожалуйста, введите другое имя.")
                return
            self.name.setText(name)
            self.writeData()
            self.updateUI()

    def writeData(self):
        self.rns.setText(f"RnS: {round(self.param_table.get_column_value(0, ParamTableColumns.RNS), 1)}")
        self.drift.setText(f"Уход: {round(self.param_table.get_column_value(0, ParamTableColumns.DRIFT), 3)}")
        # parent().parent() points to the main app widget (RnSApp)
        self.parent().parent().calculate_means()
        self.parent().parent().addCellData(cell=self.index, name=self.name.text())

    def updateUI(self):
        if self.rns.text() is not None and self.drift.text() is not None:
            self.writeButton.setVisible(False)
            self.rns.setVisible(True)
            self.drift.setVisible(True)
            self.checkbox.setVisible(True)
            self.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)

    def buildGraph(self, state):
        cell_data = Store.data.get(cell=self.index)
        if state == QtCore.Qt.CheckState.Checked:
            self.parent().parent().plot_data(self.index)
            cell_data.is_plot = True
        else:
            self.parent().parent().remove_plot(self.index)
            cell_data.is_plot = False

    def openRenameDialog(self):
        # Логика для открытия окна для переименования
        name, ok = QtWidgets.QInputDialog.getText(self, "Переименование", "Введите новое имя:")
        if name and ok:
            if Store.data.exclude(cell=self.index).filter(name=name).exists():
                QtWidgets.QMessageBox.warning(self, "Ошибка", "Это имя уже существует. Пожалуйста, введите другое имя.")
                return
            self.name.setText(name)
            self.parent().parent().remove_plot(cell=self.index)
            cell_data = Store.update_or_create_item(cell=self.index, name=name)
            if cell_data.is_plot:
                self.parent().parent().plot_data(cell=self.index)

    def openRewriteDataDialog(self):
        # Логика для открытия окна для перезаписи
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Перезапись данных")
        dialog.setLayout(QtWidgets.QVBoxLayout())

        label = QtWidgets.QLabel(f"Вы уверены, что хотите перезаписать данные для ячейки {self.name.text()}?")
        dialog.layout().addWidget(label)

        buttonBox = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        dialog.layout().addWidget(buttonBox)

        cancelButton = buttonBox.button(QtWidgets.QDialogButtonBox.Cancel)

        cancelButton.setDefault(True)

        buttonBox.accepted.connect(dialog.accept)
        buttonBox.rejected.connect(dialog.reject)

        if dialog.exec() == QtWidgets.QDialog.Accepted:
            self.writeData()

    def showData(self):
        reply = QtWidgets.QMessageBox.question(
            self,
            "Показать данные",
            "Текущиее таблицы с данными и рассчетом будет перезаписаны, продолжить?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )

        if reply == QtWidgets.QMessageBox.Yes:
            self.parent().parent().reload_tables_from_cell_data(cell=self.index)

    def showContextMenu(self, position):
        menu = QtWidgets.QMenu(self)

        showAction = QtGui.QAction("Показать данные", self)
        showAction.triggered.connect(self.showData)
        menu.addAction(showAction)

        renameAction = QtGui.QAction("Переименовать", self)
        renameAction.triggered.connect(self.openRenameDialog)
        menu.addAction(renameAction)

        rewriteAction = QtGui.QAction("Перезаписать", self)
        rewriteAction.triggered.connect(self.openRewriteDataDialog)
        menu.addAction(rewriteAction)

        # Показ контекстного меню в позиции курсора
        menu.exec(self.mapToGlobal(position))

    def clear(self):
        self.name.setText("")
        self.rns.setText("")
        self.drift.setText("")
        self.rns.setVisible(False)
        self.drift.setVisible(False)
        self.checkbox.setChecked(False)
        self.checkbox.setVisible(False)
        self.writeButton.setVisible(True)
        self.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.NoContextMenu)

    def set_active(self, value: bool):
        if value:
            self.number.setText(f"<b>№{self.index}</b>")
        else:
            self.number.setText(f"№{self.index}")
