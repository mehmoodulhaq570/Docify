
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import os
from . import converters

from typing import Callable
from typing import Any

class ConverterGUI(QtWidgets.QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.init_ui()
        self.setAcceptDrops(True)
        # placeholders for worker thread and progress timer
        self._worker_thread: QtCore.QThread | None = None
        self._progress_timer: QtCore.QTimer | None = None

    def init_ui(self) -> None:
        self.setWindowTitle('Docify - File Converter')
        self.setWindowIcon(QtGui.QIcon())
        self.setGeometry(100, 100, 650, 480)
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #e0eafc, stop:1 #cfdef3);
                border-radius: 18px;
            }
        """)

        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(18)

        title = QtWidgets.QLabel('File Converter')
        title.setFont(QtGui.QFont('Segoe UI', 28, QtGui.QFont.Bold))
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setStyleSheet("color: #1e272e; margin-bottom: 18px;")
        # Use a QGraphicsDropShadowEffect instead of CSS text-shadow (Qt stylesheet does not support text-shadow)
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(8)
        shadow.setOffset(1, 1)
        shadow.setColor(QtGui.QColor('#b2bec3'))
        title.setGraphicsEffect(shadow)
        layout.addWidget(title)

        # Input file row
        input_row = QtWidgets.QHBoxLayout()
        self.input_path = QtWidgets.QLineEdit()
        self.input_path.setPlaceholderText('Select input file...')
        self.input_path.setStyleSheet("""
            QLineEdit {
                padding: 10px; border-radius: 8px; border: 1.5px solid #a5b1c2;
                font-size: 15px; background: #f7f1e3;
            }
        """)
        self.input_path.setAcceptDrops(False)  # We'll handle drop at widget level
        input_row.addWidget(self.input_path)
        btn_browse_in = QtWidgets.QPushButton('Browse')
        btn_browse_in.setStyleSheet(self.button_style(accent=True))
        btn_browse_in.clicked.connect(self.browse_input)
        input_row.addWidget(btn_browse_in)
        layout.addLayout(input_row)

        # Output file row
        output_row = QtWidgets.QHBoxLayout()
        self.output_path = QtWidgets.QLineEdit()
        self.output_path.setPlaceholderText('Select output file...')
        self.output_path.setStyleSheet("""
            QLineEdit {
                padding: 10px; border-radius: 8px; border: 1.5px solid #a5b1c2;
                font-size: 15px; background: #f7f1e3;
            }
        """)
        output_row.addWidget(self.output_path)
        btn_browse_out = QtWidgets.QPushButton('Browse')
        btn_browse_out.setStyleSheet(self.button_style(accent=True))
        btn_browse_out.clicked.connect(self.browse_output)
        output_row.addWidget(btn_browse_out)
        layout.addLayout(output_row)

        self.status = QtWidgets.QLabel('')
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        self.status.setStyleSheet("color: #353b48; margin: 16px; font-size: 15px;")
        layout.addWidget(self.status)

        # Progress bar + numeric percent label (hidden until a conversion starts)
        progress_row = QtWidgets.QHBoxLayout()
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setVisible(False)
        progress_row.addWidget(self.progress_bar)

        self.percent_label = QtWidgets.QLabel('')
        self.percent_label.setFixedWidth(60)
        self.percent_label.setAlignment(QtCore.Qt.AlignCenter)
        self.percent_label.setStyleSheet("color: #353b48; font-size: 14px; margin-left: 8px;")
        self.percent_label.setVisible(False)
        progress_row.addWidget(self.percent_label)

        layout.addLayout(progress_row)

        btns_layout = QtWidgets.QHBoxLayout()
        btns_layout.setSpacing(18)
        btn_word2pdf = QtWidgets.QPushButton('Word → PDF')
        btn_word2pdf.setStyleSheet(self.button_style(color="#00b894"))
        # expected output extension: .pdf
        btn_word2pdf.clicked.connect(lambda: self.run_conversion(converters.word_to_pdf, '.pdf'))
        btns_layout.addWidget(btn_word2pdf)

        btn_pdf2word = QtWidgets.QPushButton('PDF → Word')
        btn_pdf2word.setStyleSheet(self.button_style(color="#0984e3"))
        # Wrap pdf_to_word to pass preserve flags and set output ext to .docx
        btn_pdf2word.clicked.connect(lambda: self.run_conversion(lambda a, b: converters.pdf_to_word(a, b, preserve_images=True, preserve_tables=True), '.docx'))
        btns_layout.addWidget(btn_pdf2word)

        btn_xlsx2csv = QtWidgets.QPushButton('Excel → CSV')
        btn_xlsx2csv.setStyleSheet(self.button_style(color="#fdcb6e"))
        btn_xlsx2csv.clicked.connect(lambda: self.run_conversion(converters.xlsx_to_csv, '.csv'))
        btns_layout.addWidget(btn_xlsx2csv)

        btn_csv2xlsx = QtWidgets.QPushButton('CSV → Excel')
        btn_csv2xlsx.setStyleSheet(self.button_style(color="#e17055"))
        btn_csv2xlsx.clicked.connect(lambda: self.run_conversion(converters.csv_to_xlsx, '.xlsx'))
        btns_layout.addWidget(btn_csv2xlsx)

        layout.addLayout(btns_layout)
        self.setLayout(layout)
    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                self.input_path.setText(file_path)
            event.acceptProposedAction()
        else:
            event.ignore()

    def button_style(self, color: str = "#40739e", accent: bool = False) -> str:
        # accent: True for browse buttons
        if accent:
            return (
                "QPushButton {"
                "background-color: #636e72; color: #fff; border-radius: 8px; padding: 8px 22px;"
                "font-size: 15px; font-weight: 500; margin: 2px;"
                "}"
                "QPushButton:hover {background-color: #2d3436;}"
            )
        return (
            f"QPushButton {{background-color: {color}; color: #fff; border-radius: 12px; padding: 12px 28px;"
            "font-size: 16px; font-weight: 600; margin: 2px; border: none;}"
            f"QPushButton:hover {{background-color: #222f3e; color: #fff;}}"
        )

    def browse_input(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input file')
        if path:
            self.input_path.setText(path)

    def browse_output(self) -> None:
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Select output file')
        if path:
            self.output_path.setText(path)

    def run_conversion(self, func: Callable[[str, str], None], output_ext: str | None = None) -> None:
        inp = self.input_path.text()
        out = self.output_path.text()
        if not inp or not out:
            self.status.setText('Please select both input and output files.')
            self.status.setStyleSheet("color: #e84118;")
            QtWidgets.QMessageBox.warning(self, "Missing File", "Please select both input and output files.")
            return
        # If the user did not specify an extension for output, append the expected one.
        if output_ext:
            root, ext = os.path.splitext(out)
            if not ext:
                out = out + output_ext
            elif ext.lower() != output_ext.lower():
                # replace extension to match selected conversion
                out = root + output_ext
        # Prepare UI for conversion
        self.status.setText('Converting...')
        self.status.setStyleSheet("color: #353b48; font-size: 15px;")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.percent_label.setText('0%')
        self.percent_label.setVisible(True)
        QtCore.QCoreApplication.processEvents()

        # Disable buttons while running
        for btn in self.findChildren(QtWidgets.QPushButton):
            btn.setEnabled(False)

        # Worker thread to run the blocking conversion function
        class ConversionWorker(QtCore.QThread):
            finished_signal = QtCore.pyqtSignal(bool, str)

            def __init__(self, fn: Any, a: str, b: str) -> None:
                super().__init__()
                self.fn = fn
                self.a = a
                self.b = b

            def run(self) -> None:
                try:
                    self.fn(self.a, self.b)
                    self.finished_signal.emit(True, os.path.basename(self.b))
                except Exception as exc:  # noqa: BLE001 - keep broad catch to relay back to UI
                    self.finished_signal.emit(False, str(exc))

        # Create and start worker
        worker = ConversionWorker(func, inp, out)
        self._worker_thread = worker

        # Timer to animate progress while worker is running
        timer = QtCore.QTimer(self)
        self._progress_timer = timer
        timer.setInterval(60)

        def advance_progress() -> None:
            val = self.progress_bar.value()
            if val < 95:
                # gently step forward
                self.progress_bar.setValue(val + 3)
                self.percent_label.setText(f"{self.progress_bar.value()}%")

        timer.timeout.connect(advance_progress)

        def on_finished(success: bool, message: str) -> None:
            # stop timer and set progress to complete
            if self._progress_timer and self._progress_timer.isActive():
                self._progress_timer.stop()
            self.progress_bar.setValue(100)
            self.percent_label.setText('100%')

            # Update status and font size on success or error
            if success:
                self.status.setText(f'Success: {message}')
                # Make success message slightly larger and bold
                self.status.setStyleSheet("color: #44bd32; font-size: 17px; font-weight: 600; margin: 16px;")
                QtWidgets.QMessageBox.information(self, "Conversion Complete", f"File saved as: {message}")
            else:
                self.status.setText(f'Error: {message}')
                self.status.setStyleSheet("color: #e84118; font-size: 15px; margin: 16px;")
                QtWidgets.QMessageBox.critical(self, "Conversion Error", message)

            # Re-enable buttons
            for btn in self.findChildren(QtWidgets.QPushButton):
                btn.setEnabled(True)

            # hide progress after a short delay
            QtCore.QTimer.singleShot(900, lambda: (self.progress_bar.setVisible(False), self.percent_label.setVisible(False)))

            # cleanup
            self._worker_thread = None
            self._progress_timer = None

        worker.finished_signal.connect(on_finished)
        timer.start()
        worker.start()

def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    gui = ConverterGUI()
    gui.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
