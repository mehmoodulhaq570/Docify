from setuptools import setup, find_packages

setup(
    name="docify",
    version="0.1.0",
    description="Convert Word, PDF, Excel, and CSV files via CLI or GUI.",
    author="Mehmood Ul Haq",
    author_email="mehmoodulhaq1040@gmail.com",
    license="MIT",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "python-docx",
        "docx2pdf",
        "pdfplumber",
        "openpyxl",
        "pdf2docx",
        "PyQt5"
    ],
    entry_points={
        "console_scripts": [
            "p2w_convertor=p2w_convertor.cli:main",
        ],
    },
    url="https://github.com/mehmoodulhaq570/docify",
    project_urls={
        "Homepage": "https://github.com/mehmoodulhaq570/docify",
        "Bug Tracker": "https://github.com/mehmoodulhaq570/docify/issues"
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)
