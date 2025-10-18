# p2w_convertor

A simple and efficient tool for converting between Word, PDF, Excel, and CSV files. Supports both a command-line interface (CLI) and a modern PyQt5 GUI.

## Features

- Convert Word (.docx) to PDF and vice versa
- Convert Excel (.xlsx) to CSV and vice versa
- Batch and single file conversion modes
- Progress bar and logging in CLI
- Drag-and-drop and notifications in GUI (coming soon)
- Cross-platform (Windows, macOS, Linux)

## Installation

```bash
pip install -r requirements.txt
python setup.py install
```

## Usage

### CLI

```bash
# Word to PDF
p2w_convertor word2pdf

# PDF to Word
p2w_convertor pdf2word [--no-images] [--no-tables]

# Excel to CSV
p2w_convertor xlsx2csv

# CSV to Excel
p2w_convertor csv2xlsx
```

### GUI

```bash
python -m p2w_convertor.gui
```

## Screenshots

<!-- Add screenshots of the GUI and sample conversions here -->

## License

MIT

## Author

Mehmood Ul Haq (<mehmoodulhaq1040@gmail.com>)
