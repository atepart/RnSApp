import sys
import time

from PySide6 import QtCore, QtWidgets


class Worker(QtCore.QThread):
    progress = QtCore.Signal(int, int, float)

    def run(self):
        total = 1000
        dl = 0
        while dl < total:
            time.sleep(0.1)
            dl += 10
            self.progress.emit(dl, total, 1.5)


app = QtWidgets.QApplication(sys.argv)
dlg = QtWidgets.QProgressDialog("Test", "Cancel", 0, 100)
dlg.setMinimumDuration(0)
dlg.show()
w = Worker()


def on_p(d, t, s):
    dlg.setValue(int(d / t * 100))
    dlg.setLabelText(f"{d}/{t} ({s})")


w.progress.connect(on_p)
w.start()
QtCore.QTimer.singleShot(2000, app.quit)
app.exec()
