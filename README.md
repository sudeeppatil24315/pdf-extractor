# PDF to Excel Data Extractor

AI-powered tool that extracts structured data from unstructured PDF documents and converts them to Excel format.

**Built by:** Sudeep Patil  
**Developed for:** Turerz Sole Proprietorship  
**Project Type:** AI Agent Development Internship Task

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Run Locally

**Command Line:**
```bash
python3 pdf_to_excel_extractor.py
```

**Web Interface:**
```bash
python3 app.py
```
Then open http://localhost:5001

## Deploy to Vercel

```bash
npm install -g vercel
vercel login
vercel --prod
```

Or push to GitHub and import at vercel.com/new

## Features

- Extracts key-value pairs from unstructured text
- 100% data capture with contextual comments
- Generates structured Excel output
- Web interface with drag-and-drop upload
- Demo mode with sample data

## Project Structure

```
.
├── pdf_to_excel_extractor.py  # Core extraction logic
├── app.py                      # Flask web app
├── api/index.py                # Vercel serverless function
├── templates/index.html        # Web interface
├── vercel.json                 # Vercel config
├── requirements.txt            # Dependencies
└── Data Input.pdf              # Sample file
```

## How It Works

1. Extracts text from PDF using PyPDF2
2. Uses regex patterns to identify key-value pairs
3. Captures contextual comments from surrounding text
4. Outputs structured Excel file with columns: #, Key, Value, Comments

## Requirements

- Python 3.7+
- PyPDF2
- pandas
- openpyxl
- Flask (for web interface)

---

© 2024 Turerz Sole Proprietorship | Developed by Sudeep Patil
