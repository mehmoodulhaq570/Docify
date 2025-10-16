
import pandas as pd
from docx2pdf import convert
from pdf2docx import Converter
import logging

def word_to_pdf(input_file, output_file):
    try:
        if not input_file.lower().endswith('.docx'):
            raise ValueError('Input file must be a .docx file')
        convert(input_file, output_file)
        logging.info(f"Converted Word to PDF: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting Word to PDF: {e}")
        raise


def pdf_to_word(input_file, output_file, preserve_images=True, preserve_tables=True):
    try:
        if not input_file.lower().endswith('.pdf'):
            raise ValueError('Input file must be a .pdf file')
        cv = Converter(input_file)
        # pdf2docx preserves images and tables by default, but options can be added if needed
        cv.convert(output_file, start=0, end=None)
        cv.close()
        logging.info(f"Converted PDF to Word: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting PDF to Word: {e}")
        raise

def xlsx_to_csv(input_file, output_file):
    try:
        if not input_file.lower().endswith('.xlsx'):
            raise ValueError('Input file must be a .xlsx file')
        df = pd.read_excel(input_file)
        df.to_csv(output_file, index=False)
        logging.info(f"Converted Excel to CSV: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting Excel to CSV: {e}")
        raise

def csv_to_xlsx(input_file, output_file):
    try:
        if not input_file.lower().endswith('.csv'):
            raise ValueError('Input file must be a .csv file')
        df = pd.read_csv(input_file)
        df.to_excel(output_file, index=False)
        logging.info(f"Converted CSV to Excel: {input_file} -> {output_file}")
    except Exception as e:
        logging.error(f"Error converting CSV to Excel: {e}")
        raise
