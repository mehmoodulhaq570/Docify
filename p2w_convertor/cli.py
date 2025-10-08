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
    """Display a progress bar while performing the conversion."""
    with tqdm(total=100, desc=f"🔄 {task_name}", ncols=80, bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}") as pbar:
        # Simulate progress in 4 steps (25%, 50%, 75%, 100%)
        for _ in range(4):
            time.sleep(0.2)
            pbar.update(25)
        func(*args, **kwargs)
        pbar.n = 100
        pbar.refresh()


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

    if args.command == "word2pdf":
        input_file, output_file = get_input_output_paths(args, ".pdf")
        show_progress("Word → PDF", converters.word_to_pdf, input_file, output_file)

    elif args.command == "pdf2word":
        input_file, output_file = get_input_output_paths(args, ".docx")
        show_progress("PDF → Word", converters.pdf_to_word, input_file, output_file)

    elif args.command == "xlsx2csv":
        input_file, output_file = get_input_output_paths(args, ".csv")
        show_progress("Excel → CSV", converters.xlsx_to_csv, input_file, output_file)

    elif args.command == "csv2xlsx":
        input_file, output_file = get_input_output_paths(args, ".xlsx")
        show_progress("CSV → Excel", converters.csv_to_xlsx, input_file, output_file)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
