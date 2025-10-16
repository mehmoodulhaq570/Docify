
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import os
from . import converters

class ConverterGUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('File Converter')
        self.setWindowIcon(QtGui.QIcon())
        self.setGeometry(100, 100, 500, 350)
        self.setStyleSheet("background-color: #f5f6fa;")

        layout = QtWidgets.QVBoxLayout()

        title = QtWidgets.QLabel('File Converter')
        title.setFont(QtGui.QFont('Arial', 20, QtGui.QFont.Bold))
        title.setAlignment(QtCore.Qt.AlignCenter)
        title.setStyleSheet("color: #273c75; margin-bottom: 20px;")
        layout.addWidget(title)

        self.input_path = QtWidgets.QLineEdit()
        self.input_path.setPlaceholderText('Select input file...')
        self.input_path.setStyleSheet("padding: 6px; border-radius: 5px; border: 1px solid #dcdde1;")
        layout.addWidget(self.input_path)
        btn_browse_in = QtWidgets.QPushButton('Browse')
        btn_browse_in.setStyleSheet(self.button_style())
        btn_browse_in.clicked.connect(self.browse_input)
        layout.addWidget(btn_browse_in)

        self.output_path = QtWidgets.QLineEdit()
        self.output_path.setPlaceholderText('Select output file...')
        self.output_path.setStyleSheet("padding: 6px; border-radius: 5px; border: 1px solid #dcdde1;")
        layout.addWidget(self.output_path)
        btn_browse_out = QtWidgets.QPushButton('Browse')
        btn_browse_out.setStyleSheet(self.button_style())
        btn_browse_out.clicked.connect(self.browse_output)
        layout.addWidget(btn_browse_out)

        self.status = QtWidgets.QLabel('')
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        self.status.setStyleSheet("color: #353b48; margin: 10px;")
        layout.addWidget(self.status)

        btns_layout = QtWidgets.QHBoxLayout()
        btn_word2pdf = QtWidgets.QPushButton('Word → PDF')
        btn_word2pdf.setStyleSheet(self.button_style())
        btn_word2pdf.clicked.connect(lambda: self.run_conversion(converters.word_to_pdf))
        btns_layout.addWidget(btn_word2pdf)

        btn_pdf2word = QtWidgets.QPushButton('PDF → Word')
        btn_pdf2word.setStyleSheet(self.button_style())
        btn_pdf2word.clicked.connect(lambda: self.run_conversion(converters.pdf_to_word))
        btns_layout.addWidget(btn_pdf2word)

        btn_xlsx2csv = QtWidgets.QPushButton('Excel → CSV')
        btn_xlsx2csv.setStyleSheet(self.button_style())
        btn_xlsx2csv.clicked.connect(lambda: self.run_conversion(converters.xlsx_to_csv))
        btns_layout.addWidget(btn_xlsx2csv)

        btn_csv2xlsx = QtWidgets.QPushButton('CSV → Excel')
        btn_csv2xlsx.setStyleSheet(self.button_style())
        btn_csv2xlsx.clicked.connect(lambda: self.run_conversion(converters.csv_to_xlsx))
        btns_layout.addWidget(btn_csv2xlsx)

        layout.addLayout(btns_layout)
        self.setLayout(layout)

    def button_style(self):
        return (
            "QPushButton {"
            "background-color: #40739e; color: white; border-radius: 5px; padding: 8px 16px;"
            "font-size: 14px; margin: 4px;"
            "}"
            "QPushButton:hover {background-color: #273c75;}"
        )

    def browse_input(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input file')
        if path:
            self.input_path.setText(path)

    def browse_output(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Select output file')
        if path:
            self.output_path.setText(path)

    def run_conversion(self, func):
        inp = self.input_path.text()
        out = self.output_path.text()
        if not inp or not out:
            self.status.setText('Please select both input and output files.')
            self.status.setStyleSheet("color: #e84118;")
            return
        self.status.setText('Converting...')
        self.status.setStyleSheet("color: #353b48;")
        QtCore.QCoreApplication.processEvents()
        try:
            func(inp, out)
            self.status.setText(f'Success: {os.path.basename(out)}')
            self.status.setStyleSheet("color: #44bd32;")
        except Exception as e:
            self.status.setText(f'Error: {e}')
            self.status.setStyleSheet("color: #e84118;")

def main():
    app = QtWidgets.QApplication(sys.argv)
    gui = ConverterGUI()
    gui.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
