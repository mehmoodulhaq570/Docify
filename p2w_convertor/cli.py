import argparse
from . import converters

def main():
    parser = argparse.ArgumentParser(
        prog="p2w_convertor",
        description="Convert Word <-> PDF and Excel <-> CSV files"
    )
    subparsers = parser.add_subparsers(dest="command")

    # Word to PDF
    wp = subparsers.add_parser("word2pdf")
    wp.add_argument("input")
    wp.add_argument("output")

    # PDF to Word
    pw = subparsers.add_parser("pdf2word")
    pw.add_argument("input")
    pw.add_argument("output")

    # Excel to CSV
    xc = subparsers.add_parser("xlsx2csv")
    xc.add_argument("input")
    xc.add_argument("output")

    # CSV to Excel
    cx = subparsers.add_parser("csv2xlsx")
    cx.add_argument("input")
    cx.add_argument("output")

    args = parser.parse_args()

    if args.command == "word2pdf":
        converters.word_to_pdf(args.input, args.output)
    elif args.command == "pdf2word":
        converters.pdf_to_word(args.input, args.output)
    elif args.command == "xlsx2csv":
        converters.xlsx_to_csv(args.input, args.output)
    elif args.command == "csv2xlsx":
        converters.csv_to_xlsx(args.input, args.output)
    else:
        parser.print_help()
