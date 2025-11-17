# Local File Search Tool

An offline desktop application for searching keywords across PDF, Word, and Excel documents. Built with Python and Tkinter.

## Features

- **PDF Search**: Search across all pages in PDF files, with clickable page numbers that open the document at the specific page
- **Word Document Search**: Search through `.docx` files including text in paragraphs and tables
- **Excel Search**: Cell-level search in `.xlsx` files with clickable cell references that open Excel and navigate to specific cells
- **Multi-threaded**: Non-blocking GUI that remains responsive during searches
- **Cross-platform**: Works on Windows, macOS, and Linux

## Supported File Types

- **PDF** (`.pdf`) - Full page-level search with navigation
- **Word** (`.docx`) - Text and table content search
- **Excel** (`.xlsx`) - Cell-level search with sheet and cell navigation
- **Excel Legacy** (`.xls`) - Requires `xlrd` library

## Requirements

### Required Dependencies
```bash
pip install PyPDF2 python-docx openpyxl
```

### Optional Dependencies (for enhanced features)

**For Excel cell navigation on Windows:**
```bash
pip install pywin32
```

**For legacy Excel (.xls) support:**
```bash
pip install xlrd
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/local-file-search.git
cd local-file-search
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the application:
```bash
python file_search.py
```

### How to Use

1. **Select Directory**: Click "Browse" to select the folder containing your documents
2. **Enter Keywords**: Type comma-separated keywords (e.g., "budget, revenue, Q4")
3. **Search**: Click "Search Documents" or press Enter
4. **View Results**: Results show:
   - For PDFs: Clickable page numbers
   - For Word: "Click to Open Document" link
   - For Excel: Clickable cell references grouped by sheet (e.g., "A5", "B10")
5. **Open Documents**: Click any link to open the document at the specific location

## Features by Platform

### Windows
- PDF page navigation (best with Adobe Reader)
- Excel cell navigation with COM automation (requires `pywin32`)
- Opens files with default applications

### macOS
- PDF page navigation
- Excel file opening (manual navigation to cells)
- Native file opening support

### Linux
- PDF page navigation
- Excel file opening (manual navigation to cells)
- Uses `xdg-open` for file opening

## Project Structure

```
local-file-search/
├── file_search.py      # Main application file
├── requirements.txt    # Python dependencies
├── README.md          # This file
└── .gitignore         # Git ignore file
```

## Development

The application uses:
- **Tkinter** for the GUI
- **PyPDF2** for PDF text extraction
- **python-docx** for Word document parsing
- **openpyxl** for Excel workbook reading
- **Threading** for non-blocking searches

## License

This project is open source and available for personal and commercial use.

## Author

Created by Vinay Bhaskarla

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

