# visio2pdf

Converts all Visio files in folder recurisvely to PDF.

#### Requirements:
```
Microsoft Visio
python3
pywin32 (use pip install -r requirements.txt)
```

#### Usage:

```
usage: python3 .\visio2pdf.py [-h] [--input DIR] [--output DIR] [--collapse]

Export Visio Files to PDF

optional arguments:
  -h, --help            show this help message and exit
  --input DIR, -i DIR   The directory to recurisvely search for Visio files.
  --output DIR, -o DIR  The directory to save the PDF files.
  --collapse, -c        Save files to root of output directory.
```
