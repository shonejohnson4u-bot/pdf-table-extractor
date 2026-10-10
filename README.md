# PDF Table Extractor

A local browser app for extracting tables from scanned PDF manuals. Users
explicitly mark a reference table, set column boundaries and headings, and
assign page ranges to one or more layouts. Pages outside those ranges are
never processed, so a drawing page cannot be mistaken for a table.

## Requirements

- Python 3.10 or newer
- Tesseract OCR installed separately

Install the Python dependencies:

```powershell
py -m pip install -r .\requirements.txt
```

If Tesseract is not on `PATH`, set its executable location in PowerShell:

```powershell
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

Start the app:

```powershell
py .\app.py
```

Open `http://127.0.0.1:8765` in a browser. The app listens on localhost only.
Upload a PDF, navigate to a parts-list page, drag to mark the table body
starting at the border below its heading row (exclude the heading row), click
each internal vertical column boundary, provide headings in left-to-right
order, and assign the inclusive range of pages that share that layout.
Add more layouts for different formats. A page cannot be assigned to
overlapping ranges. Only explicitly assigned pages are OCR'd.

The extraction expects horizontal row rules within the marked table area. If
a page has no usable row boundaries or an OCR dependency is unavailable, the
export reports an error rather than silently creating a misleading workbook.
The workbook has one sheet per layout, with the layout name, column headings,
and OCR'd cell values. OCR results are not human-reviewed before export.

Uploaded PDFs are written to the system temporary directory while the app is
running and removed when the app exits normally. The built-in server is
intended for local, single-user use, not public network deployment.
