
import argparse
import os
import sys
import contextlib
import time
import logging
from tqdm import tqdm
from . import converters


# =========================================
# 🔇 Utility: Suppress unwanted console output
# =========================================
from typing import Generator, Any

@contextlib.contextmanager
def suppress_output() -> contextlib._GeneratorContextManager:
    """Temporarily suppress stdout/stderr to hide unwanted messages."""
    with open(os.devnull, "w") as devnull:
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout, sys.stderr = devnull, devnull
        try:
            yield
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr


# =========================================
# 📊 Utility: Progress bar
# =========================================
def show_progress(task_name: str, func: Any, *args: Any, **kwargs: Any) -> None:
    """Display a single, smooth progress bar."""
    print(f"\n🔄 Starting {task_name} conversion...\n")
    with tqdm(
        total=100,
        desc=f"{task_name}",
        ncols=70,
        bar_format="{desc}: [{bar}] {percentage:3.0f}%",
        ascii=("=", " "),
        leave=False,
    ) as pbar:
        for _ in range(20):
            time.sleep(0.05)
            pbar.update(5)
        try:
            with suppress_output():
                func(*args, **kwargs)
            pbar.n = 100
            pbar.refresh()
            print(f"\n✅ Conversion complete! Saved to: {args[1]}\n")
            logging.info(f"Conversion complete: {args[0]} -> {args[1]}")
        except Exception as e:
            print(f"\n❌ Conversion failed: {e}\n")
            logging.error(f"Conversion failed: {args[0]} -> {args[1]} | Error: {e}")


# =========================================
# 🧭 Utility: Ask for input/output (single)
# =========================================
def get_input_output_paths(default_ext: str, input_prompt: str, output_prompt: str) -> tuple[str, str]:
    """Ask user for input and output file paths."""
    inp = input(input_prompt).strip()
    if not os.path.exists(inp):
        print("❌ Error: Input file not found.")
        sys.exit(1)

    input_dir = os.path.dirname(inp)
    input_name = os.path.splitext(os.path.basename(inp))[0]

    out_name = input(output_prompt.format(input_name=input_name)).strip()
    if not out_name:
        out_name = input_name
    out = os.path.join(input_dir, out_name + default_ext)

    print(f"\n💾 Output will be saved as: {out}\n")
    return inp, out


# =========================================
# 📁 Utility: Get folder for batch conversion
# =========================================
def get_folder_and_files(extension: str) -> tuple[str, list[str]]:
    """Ask for folder path and get all files matching extension."""
    folder = input("📁 Enter folder path containing files: ").strip()
    if not os.path.isdir(folder):
        print("❌ Error: Invalid folder path.")
        sys.exit(1)

    files = [f for f in os.listdir(folder) if f.lower().endswith(extension)]
    if not files:
        print(f"⚠️ No {extension} files found in the specified folder.")
        sys.exit(0)

    print(f"\n📦 Found {len(files)} '{extension}' files in '{folder}'\n")
    return folder, files


# =========================================
# 🧠 Generic batch converter
# =========================================
def batch_convert(task_name: str, extension: str, output_ext: str, func: Any, **kwargs: Any) -> None:
    folder, files = get_folder_and_files(extension)
    print(f"🔄 Starting batch {task_name} conversion...\n")
    errors = []
    for f in files:
        inp = os.path.join(folder, f)
        out = os.path.join(folder, os.path.splitext(f)[0] + output_ext)
        try:
            show_progress(f"{f[:25]} → {output_ext}", func, inp, out, **kwargs)
        except Exception as e:
            errors.append((f, str(e)))
    if errors:
        print("\n❌ Some files failed to convert:")
        for fname, err in errors:
            print(f"  - {fname}: {err}")
    print(f"\n✅ Batch conversion complete! Files saved in: {folder}\n")


# =========================================
# 🚀 Unified conversion handler
# =========================================
def handle_conversion(task_name: str, input_ext: str, output_ext: str, func: Any, **kwargs: Any) -> None:
    """Ask user for mode: single or batch."""
    print("\n🔢 Select conversion mode:")
    print("1️⃣  Single file conversion")
    print("2️⃣  Batch folder conversion\n")

    choice = input("👉 Enter your choice (1 or 2): ").strip()

    if choice == "1":
        inp, out = get_input_output_paths(
            output_ext,
            input_prompt=f"📄 Enter input {input_ext} file path: ",
            output_prompt="📝 Enter output file name (press Enter to use '{input_name}'): "
        )
        show_progress(task_name, func, inp, out, **kwargs)

    elif choice == "2":
        batch_convert(task_name, input_ext, output_ext, func, **kwargs)

    else:
        print("❌ Invalid choice. Exiting.")
        sys.exit(1)


# =========================================
# 🧩 CLI main
# =========================================

def main() -> None:
    logging.basicConfig(
        filename="convertor.log",
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s"
    )

    parser = argparse.ArgumentParser(
        prog="p2w_convertor",
        description="✨ Convert Word ↔ PDF and Excel ↔ CSV files easily!"
    )

    subparsers = parser.add_subparsers(dest="command", help="Choose conversion type")

    # Add format options for pdf2word
    pdf2word_parser = subparsers.add_parser("pdf2word", help="Convert PDF → Word (.docx)")
    pdf2word_parser.add_argument("--no-images", action="store_true", help="Do not preserve images in PDF to Word conversion")
    pdf2word_parser.add_argument("--no-tables", action="store_true", help="Do not preserve tables in PDF to Word conversion")

    subparsers.add_parser("word2pdf", help="Convert Word (.docx) → PDF")
    subparsers.add_parser("xlsx2csv", help="Convert Excel (.xlsx) → CSV")
    subparsers.add_parser("csv2xlsx", help="Convert CSV → Excel (.xlsx)")

    args = parser.parse_args()

    if args.command == "word2pdf":
        handle_conversion("Word → PDF", ".docx", ".pdf", converters.word_to_pdf)

    elif args.command == "pdf2word":
        preserve_images = not getattr(args, "no_images", False)
        preserve_tables = not getattr(args, "no_tables", False)
        handle_conversion(
            "PDF → Word", ".pdf", ".docx",
            lambda inp, out: converters.pdf_to_word(inp, out, preserve_images=preserve_images, preserve_tables=preserve_tables)
        )

    elif args.command == "xlsx2csv":
        handle_conversion("Excel → CSV", ".xlsx", ".csv", converters.xlsx_to_csv)

    elif args.command == "csv2xlsx":
        handle_conversion("CSV → Excel", ".csv", ".xlsx", converters.csv_to_xlsx)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
