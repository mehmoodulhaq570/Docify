

import pandas as pd
from docx2pdf import convert
from pdf2docx import Converter
import pdfplumber
from docx import Document
import tempfile
from copy import deepcopy
import logging
import os

def word_to_pdf(input_file: str, output_file: str) -> None:
    try:
        if not input_file.lower().endswith('.docx'):
            raise ValueError('Input file must be a .docx file')
        convert(input_file, output_file)
        logging.info(f"Converted Word to PDF: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting Word to PDF: {e}")
        raise


def pdf_to_word(
    input_file: str,
    output_file: str,
    preserve_images: bool = True,
    preserve_tables: bool = True
) -> None:
    try:
        if not input_file.lower().endswith('.pdf'):
            raise ValueError('Input file must be a .pdf file')
        cv = Converter(input_file)
        # Try primary conversion using pdf2docx
        try:
            cv.convert(output_file, start=0, end=None)
            cv.close()
            logging.info(f"Converted PDF to Word: {input_file} -> {output_file}")
            return
        except Exception as primary_exc:
            logging.warning(f"Primary pdf2docx conversion failed, attempting per-page conversion: {primary_exc}")
            try:
                cv.close()
            except Exception:
                pass

            # Try per-page conversion and merge to preserve layout where possible.
            try:
                with pdfplumber.open(input_file) as pdf:
                    total_pages = len(pdf.pages)

                main_doc = Document()
                for page_idx in range(total_pages):
                    tmp = None
                    try:
                        # create a temp file for this single page conversion
                        tmpf = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
                        tmp = tmpf.name
                        tmpf.close()

                        cpage = Converter(input_file)
                        cpage.convert(tmp, start=page_idx, end=page_idx)
                        cpage.close()

                        # append converted page docx content into main_doc
                        tmp_doc = Document(tmp)
                        for element in tmp_doc.element.body:
                            main_doc.element.body.append(deepcopy(element))
                    except Exception as page_exc:
                        logging.warning(f"Page {page_idx} conversion failed: {page_exc}; extracting text for this page instead.")
                        # fallback for this page: extract text and add as paragraphs
                        try:
                            with pdfplumber.open(input_file) as pdf:
                                page = pdf.pages[page_idx]
                                text = page.extract_text()
                                if text:
                                    for line in text.split('\n'):
                                        main_doc.add_paragraph(line)
                        except Exception as text_exc:
                            logging.error(f"Failed to extract text for page {page_idx}: {text_exc}")
                    finally:
                        if tmp:
                            try:
                                os.unlink(tmp)
                            except Exception:
                                pass

                # Save merged document
                main_doc.save(output_file)
                logging.info(f"Per-page PDF->Word merge completed: {input_file} -> {output_file}")
                return
            except Exception as per_page_exc:
                logging.warning(f"Per-page conversion failed, falling back to text-only extraction: {per_page_exc}")

            # Final fallback: simple text extraction using pdfplumber and write to a .docx using python-docx
            doc = Document()
            with pdfplumber.open(input_file) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            doc.add_paragraph(line)
                    # don't add an extra page break after the last page
                    if i != len(pdf.pages) - 1:
                        doc.add_page_break()
            doc.save(output_file)
            logging.info(f"Fallback PDF->Word (text-only) completed: {input_file} -> {output_file}")
            return
    except Exception as e:
        logging.error(f"Error converting PDF to Word: {e}")
        raise

def xlsx_to_csv(input_file: str, output_file: str) -> None:
    try:
        if not input_file.lower().endswith('.xlsx'):
            raise ValueError('Input file must be a .xlsx file')
        df = pd.read_excel(input_file)
        df.to_csv(output_file, index=False)
        logging.info(f"Converted Excel to CSV: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting Excel to CSV: {e}")
        raise

def csv_to_xlsx(input_file: str, output_file: str) -> None:
    try:
        if not input_file.lower().endswith('.csv'):
            raise ValueError('Input file must be a .csv file')
        df = pd.read_csv(input_file)
        df.to_excel(output_file, index=False)
        logging.info(f"Converted CSV to Excel: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting CSV to Excel: {e}")
        raise
