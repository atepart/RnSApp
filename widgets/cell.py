from PyQt5 import QtWidgets, QtCore

from constants import ParamTableColumns


class CellWidget(QtWidgets.QGroupBox):
    def __init__(self, parent, index: int, param_table):
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

        self.rewriteButton = QtWidgets.QPushButton("Переписать данные")
        self.rewriteButton.clicked.connect(self.openRewriteDataDialog)
        layout.addWidget(self.rewriteButton)
        self.rewriteButton.setVisible(False)

        self.checkbox = QtWidgets.QCheckBox("Построить")
        self.checkbox.setVisible(False)
        self.checkbox.stateChanged.connect(self.buildGraph)
        layout.addWidget(self.checkbox)

        self.setLayout(layout)

    def openWriteDialog(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Запись", "Введите новое имя:")
        if ok and name:
            self.name.setText(name)
            self.writeData()
            self.updateUI()

    def writeData(self):
        self.rns.setText(f"{self.param_table.get_column_value(0, ParamTableColumns.RNS)}")
        self.drift.setText(f"{self.param_table.get_column_value(0, ParamTableColumns.DRIFT)}")
        self.parent().parent().calculate_means()
        self.parent().parent().addCellData(cell=self.index, name=self.name.text())

    def updateUI(self):
        if self.rns.text() is not None and self.drift.text() is not None:
            self.writeButton.setVisible(False)
            self.rns.setVisible(True)
            self.drift.setVisible(True)
            self.checkbox.setVisible(True)
            self.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
            self.customContextMenuRequested.connect(self.showContextMenu)

    def buildGraph(self, state):
        if state == QtCore.Qt.CheckState.Checked:
            self.parent().parent().plot_data(self.index)
        else:
            self.parent().parent().remove_plot(self.index)

    def openRenameDialog(self):
        # Логика для открытия окна для переименования
        name, ok = QtWidgets.QInputDialog.getText(self, "Переименование", "Введите новое имя:")
        if name and ok:
            self.name.setText(name)

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

        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.writeData()

    def showContextMenu(self, position):
        menu = QtWidgets.QMenu(self)

        renameAction = QtWidgets.QAction("Переименовать", self)
        renameAction.triggered.connect(self.openRenameDialog)
        menu.addAction(renameAction)

        rewriteAction = QtWidgets.QAction("Перезаписать", self)
        rewriteAction.triggered.connect(self.openRewriteDataDialog)
        menu.addAction(rewriteAction)

        # Показ контекстного меню в позиции курсора
        menu.exec_(self.mapToGlobal(position))
