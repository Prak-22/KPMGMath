# Financial Statement PDF to Excel Converter

A Streamlit application that extracts tables from financial statement PDFs (like 10-K reports), converts them to Excel with calculation verification rows, and provides a Q&A interface to ask questions about the financial data.

## Features

- Extract tables from PDF documents using Python-only libraries
- Convert tables to Excel format
- Add calculation verification rows to check totals
- Ask questions about the financial data using OpenAI's language models
- Download the processed Excel file

## System Requirements

- Python 3.7 or higher
- No external dependencies like Java required - 100% Python-based solution

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd <repository-directory>
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Get an OpenAI API key from [https://platform.openai.com/account/api-keys](https://platform.openai.com/account/api-keys)

## Usage

1. Run the Streamlit application:
```bash
streamlit run app.py
```

2. Open your web browser and navigate to the URL shown in the terminal (typically http://localhost:8501)

3. Follow the steps in the application:
   - Upload a financial statement PDF
   - Extract tables from the PDF
   - Generate an Excel file with calculations
   - Download the Excel file
   - Ask questions about the financial data

## Example Questions

- "Were Total Assets matching with Total Liabilities and Stockholders' Equity for both periods?"
- "What is the value of Total Current Assets for the period July 29, 2023?"
- "What is the value of Total Assets as of July 29, 2023?"

## How It Works

1. **Table Extraction**:
   - Uses pdfplumber and PyPDF (pure Python libraries) to extract tables from PDFs
   - Handles duplicate column names and cleans extracted data

2. **Excel Generation**:
   - Converts extracted tables to Excel format
   - Identifies rows containing totals
   - Adds calculation rows with SUM formulas to verify totals
   - Adds difference rows to show discrepancies

3. **Question Answering**:
   - Uses OpenAI's language models to answer questions about the financial data
   - For common questions, attempts direct lookup from the data

## PDF Extraction Notes

- PDF extraction can vary in accuracy depending on the PDF structure
- For best results, use PDFs with well-defined tables
- Some complex PDFs may require additional processing

## Dependencies

- streamlit: Web application framework
- pandas & numpy: Data manipulation
- pdfplumber & PyPDF: PDF table extraction (pure Python)
- openpyxl: Excel file manipulation
- openai: OpenAI API client

## Notes

- PDF extraction can vary in accuracy depending on the PDF structure
- For best results, use PDFs with well-defined tables
- OpenAI API usage may incur costs based on your OpenAI plan
