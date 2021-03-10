# visio2pdf

Converts all Visio files in folder recurisvely to PDF.

#### Requirements:
```
python3
pywin32 (use pip install -r requirements.txt)
```

#### Usage:

```
usage: python3 .\visio2pdf.py [-h] [--input DIR] [--output DIR] [--collapse]

Export Visio Files to PDF

optional arguments:
  -h, --help            show this help message and exit
  --input DIR, -i DIR   The directory to recurisvely search for Visio files (Defaults to current directory).
  --output DIR, -o DIR  The directory to save the exported files. (Defaults to current directory)
  --collapse, -c        Save files to root of output directory / do not preserve folder structure
  --vsd, -v             Convert all .vsd files in input directory to .vsdx format, and save in output directory
  --pdf, -p             Convert all .vsdx files in input directory to PDF, and save in output directory
```