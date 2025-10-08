import argparse
import os
import sys
import contextlib
import time
from tqdm import tqdm
from . import converters


# =========================================
# 🔇 Utility: Suppress unwanted console output
# =========================================
@contextlib.contextmanager
def suppress_output():
    """Temporarily suppress stdout/stderr to hide library messages."""
    with open(os.devnull, "w") as devnull:
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout, sys.stderr = devnull, devnull
        try:
            yield
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr


# =========================================
# 📁 Utility: Ask for input/output interactively
# =========================================
def get_input_output_paths(args, default_ext):
    """Prompt for input/output file paths if not given."""
    if not args.input:
        args.input = input("📄 Enter input file path: ").strip()

    # Validate input path
    if not os.path.exists(args.input):
        print("❌ Error: Input file not found. Please check the path and try again.\n")
        sys.exit(1)

    input_dir = os.path.dirname(args.input)
    input_name = os.path.splitext(os.path.basename(args.input))[0]

    # Ask for output name if not provided
    if not args.output:
        out_name = input(f"📝 Enter output file name (press Enter to use '{input_name}'): ").strip()
        if not out_name:
            out_name = input_name
        args.output = os.path.join(input_dir, out_name + default_ext)

    print(f"💾 Output will be saved as: {args.output}\n")
    return args.input, args.output


# =========================================
# 📊 Utility: Single clean progress display
# =========================================
def show_progress(task_name, func, *args, **kwargs):
    """Display a single progress bar using '===' style."""
    print(f"\n🔄 Starting {task_name} conversion...\n")

    with tqdm(
        total=100,
        desc=f"{task_name}",
        ncols=70,
        bar_format="{desc}: [{bar}] {percentage:3.0f}%",
        ascii=("=", " "),   # ✅ fixed tuple style — prevents ZeroDivisionError
        leave=False,
    ) as pbar:
        # Simulated smooth progress animation
        for _ in range(20):
            time.sleep(0.05)
            pbar.update(5)

        # Run conversion silently
        with suppress_output():
            func(*args, **kwargs)

        # Mark as complete
        pbar.n = 100
        pbar.refresh()

    print(f"\n✅ Conversion complete! Saved to: {args[1]}\n")


# =========================================
# 🚀 Main CLI logic
# =========================================
def main():
    parser = argparse.ArgumentParser(
        prog="p2w_convertor",
        description="✨ Convert Word ↔ PDF and Excel ↔ CSV files easily!"
    )
    subparsers = parser.add_subparsers(dest="command", help="Choose a conversion type")

    # Word to PDF
    wp = subparsers.add_parser("word2pdf", help="Convert Word (.docx) → PDF")
    wp.add_argument("input", nargs="?", help="Path to the Word file")
    wp.add_argument("output", nargs="?", help="Path to save the converted PDF")

    # PDF to Word
    pw = subparsers.add_parser("pdf2word", help="Convert PDF → Word (.docx)")
    pw.add_argument("input", nargs="?", help="Path to the PDF file")
    pw.add_argument("output", nargs="?", help="Path to save the converted Word file")

    # Excel to CSV
    xc = subparsers.add_parser("xlsx2csv", help="Convert Excel (.xlsx) → CSV")
    xc.add_argument("input", nargs="?", help="Path to the Excel file")
    xc.add_argument("output", nargs="?", help="Path to save the converted CSV")

    # CSV to Excel
    cx = subparsers.add_parser("csv2xlsx", help="Convert CSV → Excel (.xlsx)")
    cx.add_argument("input", nargs="?", help="Path to the CSV file")
    cx.add_argument("output", nargs="?", help="Path to save the converted Excel file")

    args = parser.parse_args()

    # Match conversion mode
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
