# TAMEER Scaffolding Quotation Generator & CRM

An enterprise quotation engine and lead management system for **Tameer Al Mesaha Contracting Est.**, pre-loaded with a catalog of 136 standard scaffolding components and official quotation formatting.

## Features
- **Interactive Quotation Builder**: Dynamic line-item table with real-time autocomplete across 136 scaffolding items, automatic subtotal, 15% VAT, and Grand Total calculations.
- **Sequential Reference Numbering**: Official `TMR-FZ-{SEQ}-55471-{YEAR}` format continuation.
- **Standardized Arial Typography**: Pixel-perfect alignment and formatting matching official company proposals.
- **Multi-Format Export**:
  - Word document (`.docx`) using the official letterhead, watermark, and footer.
  - Formatted Excel sheet (`.xlsx`) with formulas and number-to-words.
  - Native PDF document (`.pdf`).
  - Clean web print view (`window.print()`).
- **Catalog Management**: In-app interface to manage default prices and units for all scaffolding components.

## Getting Started

### Prerequisites
- Python 3.10+
- Microsoft Word (optional, for native `.docx` to `.pdf` export)

### Installation
```bash
pip install -r requirements.txt
```

### Running the App Locally
- Double-click `run.bat`, or run:
```bash
python app.py
```
- Open [http://localhost:5001](http://localhost:5001) in your browser.

## Cloud Deployment (Vercel)
- Fully configured with `vercel.json` and serverless entrypoint in `api/index.py`.
- Uses cross-platform dependencies without platform-specific binary requirements.
- Uses Python 3.12 and lightweight serverless SQLite storage.

