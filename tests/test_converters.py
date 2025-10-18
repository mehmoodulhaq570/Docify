import os
import pytest
from p2w_convertor import converters
import pandas as pd

def test_word_to_pdf(tmp_path):
    # Dummy test: just check error on wrong extension
    with pytest.raises(ValueError):
        converters.word_to_pdf('file.txt', 'out.pdf')

def test_pdf_to_word(tmp_path):
    with pytest.raises(ValueError):
        converters.pdf_to_word('file.txt', 'out.docx')

def test_xlsx_to_csv(tmp_path):
    # Create a sample xlsx
    df = pd.DataFrame({'a': [1,2], 'b': [3,4]})
    xlsx = tmp_path / 'test.xlsx'
    df.to_excel(xlsx, index=False)
    csv = tmp_path / 'test.csv'
    converters.xlsx_to_csv(str(xlsx), str(csv))
    assert os.path.exists(csv)
    df2 = pd.read_csv(csv)
    assert df2.equals(df)

def test_csv_to_xlsx(tmp_path):
    df = pd.DataFrame({'a': [1,2], 'b': [3,4]})
    csv = tmp_path / 'test.csv'
    df.to_csv(csv, index=False)
    xlsx = tmp_path / 'test.xlsx'
    converters.csv_to_xlsx(str(csv), str(xlsx))
    assert os.path.exists(xlsx)
    df2 = pd.read_excel(xlsx)
    assert df2.equals(df)
