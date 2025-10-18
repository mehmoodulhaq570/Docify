
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import os
from . import converters

from typing import Callable

class ConverterGUI(QtWidgets.QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.init_ui()
        self.setAcceptDrops(True)

    def init_ui(self) -> None:
        self.setWindowTitle('File Converter')
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
        title.setStyleSheet("color: #1e272e; margin-bottom: 18px; text-shadow: 1px 1px 2px #b2bec3;")
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

        btns_layout = QtWidgets.QHBoxLayout()
        btns_layout.setSpacing(18)
        btn_word2pdf = QtWidgets.QPushButton('Word → PDF')
        btn_word2pdf.setStyleSheet(self.button_style(color="#00b894"))
        btn_word2pdf.clicked.connect(lambda: self.run_conversion(converters.word_to_pdf))
        btns_layout.addWidget(btn_word2pdf)

        btn_pdf2word = QtWidgets.QPushButton('PDF → Word')
        btn_pdf2word.setStyleSheet(self.button_style(color="#0984e3"))
        btn_pdf2word.clicked.connect(lambda: self.run_conversion(converters.pdf_to_word))
        btns_layout.addWidget(btn_pdf2word)

        btn_xlsx2csv = QtWidgets.QPushButton('Excel → CSV')
        btn_xlsx2csv.setStyleSheet(self.button_style(color="#fdcb6e"))
        btn_xlsx2csv.clicked.connect(lambda: self.run_conversion(converters.xlsx_to_csv))
        btns_layout.addWidget(btn_xlsx2csv)

        btn_csv2xlsx = QtWidgets.QPushButton('CSV → Excel')
        btn_csv2xlsx.setStyleSheet(self.button_style(color="#e17055"))
        btn_csv2xlsx.clicked.connect(lambda: self.run_conversion(converters.csv_to_xlsx))
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
                "box-shadow: 0 2px 8px #b2bec3;"
                "}"
                "QPushButton:hover {background-color: #2d3436;}"
            )
        return (
            f"QPushButton {{background-color: {color}; color: #fff; border-radius: 12px; padding: 12px 28px;"
            "font-size: 16px; font-weight: 600; margin: 2px; box-shadow: 0 2px 8px #b2bec3; border: none;}"
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

    def run_conversion(self, func: Callable[[str, str], None]) -> None:
        inp = self.input_path.text()
        out = self.output_path.text()
        if not inp or not out:
            self.status.setText('Please select both input and output files.')
            self.status.setStyleSheet("color: #e84118;")
            QtWidgets.QMessageBox.warning(self, "Missing File", "Please select both input and output files.")
            return
        self.status.setText('Converting...')
        self.status.setStyleSheet("color: #353b48;")
        QtCore.QCoreApplication.processEvents()
        try:
            func(inp, out)
            self.status.setText(f'Success: {os.path.basename(out)}')
            self.status.setStyleSheet("color: #44bd32;")
            QtWidgets.QMessageBox.information(self, "Conversion Complete", f"File saved as: {os.path.basename(out)}")
        except Exception as e:
            self.status.setText(f'Error: {e}')
            self.status.setStyleSheet("color: #e84118;")
            QtWidgets.QMessageBox.critical(self, "Conversion Error", str(e))

def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    gui = ConverterGUI()
    gui.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
