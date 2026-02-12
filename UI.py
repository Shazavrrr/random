import sys
import os
import json
import win32com.client
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout,
    QHBoxLayout, QWidget, QComboBox, QLabel, QListWidget,
    QListWidgetItem, QFileDialog, QLineEdit, QFrame,
    QProgressBar, QTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont


class PlotWorker(QThread):
    progress_changed = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished_signal = pyqtSignal(object)

    def __init__(self, printer, output_dir, frames, plotter_module):
        super().__init__()
        self.printer = printer
        self.output_dir = output_dir
        self.frames = frames
        self.plotter = plotter_module

    def run(self):
        try:
            total = len(self.frames)
            result = self.plotter.start_plot_process(
                self.printer,
                self.output_dir,
                self.frames,
                progress_callback=self.update_progress,
                log_callback=self.update_log
            )
            self.finished_signal.emit(result)
        except Exception as e:
            self.finished_signal.emit(str(e))

    def update_progress(self, value):
        self.progress_changed.emit(value)

    def update_log(self, message):
        self.log_message.emit(message)


class AutoCADPlotterUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoCAD Smart Plotter v4.0")
        self.setMinimumSize(700, 800)

        self.json_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "frames_data.json"
        )

        try:
            import search
            import plotter
            self.search = search
            self.plotter = plotter
        except ImportError as e:
            raise RuntimeError(f"Ошибка импорта модулей: {e}")

        self.init_ui()
        self.fill_printers()

    def init_ui(self):
        self.main_widget = QWidget()
        self.layout = QVBoxLayout()

        # === Поиск ===
        self.btn_scan = QPushButton("🔍 НАЙТИ РАМКИ")
        self.btn_scan.setFixedHeight(50)
        self.btn_scan.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.btn_scan.clicked.connect(self.handle_search)
        self.layout.addWidget(self.btn_scan)

        header_layout = QHBoxLayout()
        h1 = QLabel(f"{'№ Листа':<15}")
        h2 = QLabel(f"{'Формат':<15}")
        for h in (h1, h2):
            h.setFont(QFont("Courier New", 10, QFont.Weight.Bold))
            header_layout.addWidget(h)
        self.layout.addLayout(header_layout)

        self.sheet_list = QListWidget()
        self.sheet_list.setFont(QFont("Courier New", 10))
        self.layout.addWidget(self.sheet_list)

        select_layout = QHBoxLayout()
        self.btn_all = QPushButton("Выделить все")
        self.btn_none = QPushButton("Снять все")
        self.btn_all.clicked.connect(lambda: self.set_all_checks(True))
        self.btn_none.clicked.connect(lambda: self.set_all_checks(False))
        select_layout.addWidget(self.btn_all)
        select_layout.addWidget(self.btn_none)
        self.layout.addLayout(select_layout)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        self.layout.addWidget(line)

        # === Плоттер ===
        self.layout.addWidget(QLabel("<b>Плоттер (PC3):</b>"))
        self.printer_select = QComboBox()
        self.layout.addWidget(self.printer_select)

        # === Папка ===
        self.layout.addWidget(QLabel("<b>Папка для PDF:</b>"))
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.btn_browse = QPushButton("Обзор")
        self.btn_browse.clicked.connect(self.browse_folder)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.btn_browse)
        self.layout.addLayout(path_layout)

        # === Печать ===
        self.btn_plot = QPushButton("🚀 ЗАПУСТИТЬ ПЕЧАТЬ")
        self.btn_plot.setFixedHeight(60)
        self.btn_plot.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.btn_plot.clicked.connect(self.start_plotting)
        self.layout.addWidget(self.btn_plot)

        # === Прогресс ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar)

        # === Лог ===
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output)

        # === Статус ===
        self.status_bar = QLabel("Статус: Готов")
        self.layout.addWidget(self.status_bar)

        self.main_widget.setLayout(self.layout)
        self.setCentralWidget(self.main_widget)

    def fill_printers(self):
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            devices = acad.ActiveDocument.ActiveLayout.GetPlotDeviceNames()
            self.printer_select.addItems(devices)
            idx = self.printer_select.findText("DWG To PDF.pc3", Qt.MatchFlag.MatchContains)
            if idx >= 0:
                self.printer_select.setCurrentIndex(idx)
        except Exception:
            self.status_bar.setText("Ошибка: AutoCAD не запущен")

    def handle_search(self):
        self.sheet_list.clear()
        self.log_output.clear()
        self.status_bar.setText("Поиск рамок...")
        QApplication.processEvents()

        try:
            count = self.search.analyze_to_json(self.json_path)
            if isinstance(count, int):
                with open(self.json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for entry in data:
                    text = f"{str(entry['sheet']):<18} | {entry['format']}"
                    item = QListWidgetItem(text)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(Qt.CheckState.Checked)
                    self.sheet_list.addItem(item)
                self.status_bar.setText(f"Найдено листов: {count}")
            else:
                self.status_bar.setText(f"Ошибка: {count}")
        except Exception as e:
            self.status_bar.setText(f"Ошибка поиска: {e}")

    def set_all_checks(self, state):
        for i in range(self.sheet_list.count()):
            self.sheet_list.item(i).setCheckState(
                Qt.CheckState.Checked if state else Qt.CheckState.Unchecked
            )

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Папка для PDF")
        if folder:
            self.path_input.setText(folder)

    def start_plotting(self):
        output_dir = self.path_input.text()
        if not output_dir or not os.path.exists(output_dir):
            self.status_bar.setText("Выберите корректную папку")
            return

        try:
            with open(self.json_path, "r", encoding="utf-8") as f:
                all_frames = json.load(f)
        except Exception:
            self.status_bar.setText("Сначала выполните поиск")
            return

        selected_frames = []
        for i in range(self.sheet_list.count()):
            if self.sheet_list.item(i).checkState() == Qt.CheckState.Checked:
                selected_frames.append(all_frames[i])

        if not selected_frames:
            self.status_bar.setText("Не выбрано ни одного листа")
            return

        self.btn_plot.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(selected_frames))

        self.worker = PlotWorker(
            self.printer_select.currentText(),
            output_dir,
            selected_frames,
            self.plotter
        )

        self.worker.progress_changed.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.log_output.append)
        self.worker.finished_signal.connect(self.on_plot_finished)

        self.worker.start()
        self.status_bar.setText("Печать...")

    def on_plot_finished(self, result):
        self.btn_plot.setEnabled(True)
        if isinstance(result, int):
            self.status_bar.setText(f"Готово. Напечатано: {result}")
        else:
            self.status_bar.setText(f"Ошибка: {result}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutoCADPlotterUI()
    window.show()
    sys.exit(app.exec())
