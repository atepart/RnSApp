import sys
import time

import requests
from PySide6 import QtCore, QtWidgets


class Worker(QtCore.QThread):
    progress = QtCore.Signal(int, int, float)

    def run(self):
        url = "https://github.com/atepart/RnSApp/releases/download/new107b3/RnSApp_macOS_arm64_new107b3.zip"
        resp = requests.get(url, stream=True, timeout=10)
        total = int(resp.headers.get("content-length", 0))
        dl = 0
        st = time.time()
        lt = st
        for chunk in resp.iter_content(chunk_size=131072):
            if chunk:
                dl += len(chunk)
                now = time.time()
                if total and (now - lt > 0.1):
                    self.progress.emit(dl, total, (dl / 1024 / 1024) / (now - st))
                    lt = now


app = QtWidgets.QApplication(sys.argv)
dlg = QtWidgets.QProgressDialog("Test", "Cancel", 0, 100)
dlg.setMinimumDuration(0)
dlg.show()
w = Worker()


def on_p(d, t, s):
    dlg.setValue(int(d / t * 100))
    dlg.setLabelText(f"{d}/{t} ({s:.1f} MB/s)")
    if d >= t:
        app.quit()


w.progress.connect(on_p)
w.start()
app.exec()
