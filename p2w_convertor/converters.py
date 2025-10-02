import pandas as pd
from docx import Document
from docx2pdf import convert
import pdfplumber

def word_to_pdf(input_file, output_file):
    convert(input_file, output_file)

def pdf_to_word(input_file, output_file):
    doc = Document()
    with pdfplumber.open(input_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            doc.add_paragraph(text)
    doc.save(output_file)

def xlsx_to_csv(input_file, output_file):
    df = pd.read_excel(input_file)
    df.to_csv(output_file, index=False)

def csv_to_xlsx(input_file, output_file):
    df = pd.read_csv(input_file)
    df.to_excel(output_file, index=False)
