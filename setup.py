from setuptools import setup, find_packages

setup(
    name="p2w_convertor",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "python-docx",
        "docx2pdf",
        "pdfplumber",
        "openpyxl"
    ],
    entry_points={
        "console_scripts": [
            "p2w_convertor=p2w_convertor.cli:main",
        ],
    },
)
