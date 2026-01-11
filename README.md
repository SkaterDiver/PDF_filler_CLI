# PDF Filler CLI

A Python CLI tool that fills `.docx` templates by replacing `[placeholder]` text with user input, then exports to PDF.

## Features

- Auto-detects `[placeholder]` fields in documents (paragraphs and tables)
- Auto-fills `[Date]` with current date (YYYY-MM-DD)
- Outputs PDFs named `CoverLetter_[Company]_[Date].pdf`
- Loops for batch processing multiple templates

## Requirements

- Python 3.x
- LibreOffice (for PDF conversion)

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python cv-filler.py
```

1. Select a template by number
2. Enter values for each placeholder
3. PDF saves to `Outputs/` folder
4. Repeat or type `exit` to quit

## Template Format

Create `.docx` files in `Templates/` with placeholders in square brackets:

```
Dear [Company Name],

I am applying for the [Position Title] role...
```