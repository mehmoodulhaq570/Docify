import argparse
import os
from . import converters
from tqdm import tqdm
import time


def get_input_output_paths(args, default_ext):
    """Get input/output file paths interactively if not provided."""
    if not args.input:
        args.input = input("📄 Enter input file path: ").strip()

    if not args.output:
        base, _ = os.path.splitext(args.input)
        args.output = base + default_ext
        print(f"💾 Output will be saved as: {args.output}")

    return args.input, args.output


def show_progress(task_name, func, *args, **kwargs):
    """Display a clean progress bar while performing conversion."""
    print(f"\n🔄 Starting {task_name} conversion...\n")

    with tqdm(
        total=100,
        desc=f"{task_name}",
        ncols=70,
        bar_format="{desc}: [{bar}] {percentage:3.0f}%",
        ascii="=>"
    ) as pbar:
        for _ in range(10):
            time.sleep(0.1)  # simulate small steps
            pbar.update(10)
        func(*args, **kwargs)
        pbar.n = 100
        pbar.refresh()

    print(f"✅ Conversion complete! Saved to: {args[1]}\n")


def main():
    parser = argparse.ArgumentParser(
        prog="p2w_convertor",
        description="✨ Convert Word <-> PDF and Excel <-> CSV files easily!"
    )
    subparsers = parser.add_subparsers(dest="command", help="Choose a conversion type")

    # Word to PDF
    wp = subparsers.add_parser("word2pdf", help="Convert Word (.docx) to PDF")
    wp.add_argument("input", nargs="?", help="Path to the Word file")
    wp.add_argument("output", nargs="?", help="Path to save the converted PDF")

    # PDF to Word
    pw = subparsers.add_parser("pdf2word", help="Convert PDF to Word (.docx)")
    pw.add_argument("input", nargs="?", help="Path to the PDF file")
    pw.add_argument("output", nargs="?", help="Path to save the converted Word file")

    # Excel to CSV
    xc = subparsers.add_parser("xlsx2csv", help="Convert Excel (.xlsx) to CSV")
    xc.add_argument("input", nargs="?", help="Path to the Excel file")
    xc.add_argument("output", nargs="?", help="Path to save the converted CSV")

    # CSV to Excel
    cx = subparsers.add_parser("csv2xlsx", help="Convert CSV to Excel (.xlsx)")
    cx.add_argument("input", nargs="?", help="Path to the CSV file")
    cx.add_argument("output", nargs="?", help="Path to save the converted Excel file")

    args = parser.parse_args()

    # Dispatcher
    if args.command == "word2pdf":
        inp, out = get_input_output_paths(args, ".pdf")
        show_progress("Word → PDF", converters.word_to_pdf, inp, out)

    elif args.command == "pdf2word":
        inp, out = get_input_output_paths(args, ".docx")
        show_progress("PDF → Word", converters.pdf_to_word, inp, out)

    elif args.command == "xlsx2csv":
        inp, out = get_input_output_paths(args, ".csv")
        show_progress("Excel → CSV", converters.xlsx_to_csv, inp, out)

    elif args.command == "csv2xlsx":
        inp, out = get_input_output_paths(args, ".xlsx")
        show_progress("CSV → Excel", converters.csv_to_xlsx, inp, out)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
