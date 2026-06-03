import argparse
import logging
import os
import shutil
import subprocess
import sys
import time


def setup_logger(install_dir: str):
    log_file = os.path.join(install_dir, "updater.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )


def is_process_running(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def count_files(dir_path: str, exclude_names: set[str]) -> int:
    count = 0
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file not in exclude_names:
                count += 1
    return count


try:
    from PySide6 import QtCore, QtWidgets

    class UpdaterWorker(QtCore.QThread):
        progress = QtCore.Signal(int)
        status = QtCore.Signal(str)
        finished = QtCore.Signal(bool, str)

        def __init__(self, pid: int, src_dir: str, dst_dir: str, exe_path: str, parent=None) -> None:
            super().__init__(parent)
            self.pid = pid
            self.src_dir = src_dir
            self.dst_dir = dst_dir
            self.exe_path = exe_path
            self.exclude_names = {os.path.basename(sys.executable), "updater.log", "updater", "updater.exe"}

        def run(self):
            try:
                # 1. Wait for process
                self.status.emit("Ожидание завершения приложения...")
                timeout = 60
                start_time = time.time()
                while time.time() - start_time < timeout:
                    if not is_process_running(self.pid):
                        break
                    time.sleep(0.5)
                else:
                    self.finished.emit(False, f"Таймаут ожидания завершения процесса {self.pid}.")
                    return

                time.sleep(1)

                # 2. Copy files
                self.status.emit("Подсчет файлов...")
                total_files = count_files(self.src_dir, self.exclude_names)
                copied_files = 0

                self.status.emit("Копирование файлов...")
                self._copy_tree(self.src_dir, self.dst_dir, total_files, [copied_files])

                # 3. Start app
                self.status.emit("Запуск обновленного приложения...")
                if os.path.exists(self.exe_path):
                    if sys.platform == "darwin" and self.exe_path.endswith(".app"):
                        subprocess.Popen(["open", self.exe_path])
                    else:
                        subprocess.Popen([self.exe_path])
                    self.finished.emit(True, "")
                else:
                    self.finished.emit(False, f"Исполняемый файл не найден: {self.exe_path}")

            except Exception as e:
                logging.error(f"Update error: {e}", exc_info=True)
                self.finished.emit(False, str(e))

        def _copy_tree(self, src: str, dst: str, total: int, copied: list[int]):
            for item in os.listdir(src):
                if item in self.exclude_names:
                    continue

                s = os.path.join(src, item)
                d = os.path.join(dst, item)

                if os.path.isdir(s):
                    if not os.path.exists(d):
                        os.makedirs(d)
                    self._copy_tree(s, d, total, copied)
                else:
                    max_retries = 5
                    for attempt in range(max_retries):
                        try:
                            if os.path.exists(d):
                                os.remove(d)
                            shutil.copy2(s, d)
                            copied[0] += 1
                            if total > 0:
                                self.progress.emit(int((copied[0] / total) * 100))
                            break
                        except PermissionError:
                            time.sleep(1)
                    else:
                        raise PermissionError(f"Не удалось скопировать {s} в {d}")

    def run_gui(args):
        QtWidgets.QApplication(sys.argv)

        dlg = QtWidgets.QDialog()
        dlg.setWindowTitle("Обновление RnSApp")
        dlg.resize(400, 100)

        layout = QtWidgets.QVBoxLayout(dlg)

        status_label = QtWidgets.QLabel("Инициализация...", dlg)
        layout.addWidget(status_label)

        progress_bar = QtWidgets.QProgressBar(dlg)
        progress_bar.setRange(0, 100)
        progress_bar.setValue(0)
        layout.addWidget(progress_bar)

        worker = UpdaterWorker(args.pid, args.src, args.dst, args.exe)

        def on_status(msg):
            status_label.setText(msg)

        def on_progress(val):
            progress_bar.setValue(val)

        def on_finished(success, err_msg):
            if not success:
                QtWidgets.QMessageBox.critical(dlg, "Ошибка обновления", err_msg)
            dlg.accept()

        worker.status.connect(on_status)
        worker.progress.connect(on_progress)
        worker.finished.connect(on_finished)

        worker.start()
        dlg.exec()
        sys.exit(0)

except ImportError:
    # Fallback to CLI if PySide6 is not available in the environment
    def run_cli(args):
        # 1. Wait for main app to exit
        logging.info("Ожидание завершения приложения...")
        timeout = 60
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not is_process_running(args.pid):
                break
            time.sleep(0.5)
        else:
            logging.error(f"Таймаут ожидания завершения процесса {args.pid}.")
            sys.exit(1)

        time.sleep(1)

        # 2. Copy files
        updater_exe_name = os.path.basename(sys.executable)
        exclude_names = {updater_exe_name, "updater.log", "updater", "updater.exe"}

        try:
            logging.info("Начало копирования файлов...")

            def copy_tree_cli(src_dir: str, dst_dir: str):
                for item in os.listdir(src_dir):
                    if item in exclude_names:
                        continue

                    s = os.path.join(src_dir, item)
                    d = os.path.join(dst_dir, item)

                    if os.path.isdir(s):
                        if not os.path.exists(d):
                            os.makedirs(d)
                        copy_tree_cli(s, d)
                    else:
                        max_retries = 5
                        for attempt in range(max_retries):
                            try:
                                if os.path.exists(d):
                                    os.remove(d)
                                shutil.copy2(s, d)
                                break
                            except PermissionError:
                                time.sleep(1)
                        else:
                            raise PermissionError(f"Не удалось скопировать {s} в {d}")

            copy_tree_cli(args.src, args.dst)
            logging.info("Копирование успешно завершено.")

            # 3. Start the new app
            if os.path.exists(args.exe):
                logging.info(f"Запуск обновленного приложения: {args.exe}")
                if sys.platform == "darwin" and args.exe.endswith(".app"):
                    subprocess.Popen(["open", args.exe])
                else:
                    subprocess.Popen([args.exe])
            else:
                logging.error(f"Исполняемый файл не найден: {args.exe}")

        except Exception as e:
            logging.error(f"Ошибка во время обновления: {e}", exc_info=True)


def main():
    parser = argparse.ArgumentParser(description="RnSApp Updater")
    parser.add_argument("--pid", type=int, required=True, help="PID of the main app to wait for")
    parser.add_argument("--src", type=str, required=True, help="Source directory (unpacked new version)")
    parser.add_argument("--dst", type=str, required=True, help="Destination directory (current install dir)")
    parser.add_argument("--exe", type=str, required=True, help="Path to the executable to start after update")

    args = parser.parse_args()

    setup_logger(args.dst)
    logging.info("--- Запуск updater ---")
    logging.info(f"PID: {args.pid}")
    logging.info(f"Source: {args.src}")
    logging.info(f"Destination: {args.dst}")
    logging.info(f"Executable to start: {args.exe}")

    if "run_gui" in globals():
        run_gui(args)
    else:
        run_cli(args)

    logging.info("--- Завершение updater ---")


if __name__ == "__main__":
    main()
