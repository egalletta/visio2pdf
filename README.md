# visio2pdf

Converts all Visio files in folder recurisvely to PDF.

#### Requirements:
```
Microsoft Visio
python3
pywin32 (use pip install -r requirements.txt)
```

#### Command Line Reference:

```
usage: python3 .\visio2pdf.py [-h] [--input DIR] [--output DIR] [--collapse]

Export Visio Files to PDF

optional arguments:
  -h, --help            show this help message and exit
  --input DIR, -i DIR   The directory to recursively search for Visio files (Defaults to current directory).
  --output DIR, -o DIR  The directory to save the exported files. (Defaults to current directory)
  --collapse, -c        Save files to root of output directory / do not preserve folder structure
  --vsd, -v             Convert all .vsd files in input directory to .vsdx format, and save in output directory
  --pdf, -p             Convert all .vsdx files in input directory to PDF, and save in output directory
```

### Typical Usage

In order to recursively convert Visio files within a folder, use the
`--input` flag followed by the full path to that folder. This tool
will then save the converted files to the directory specified by the
`--output` flag, which could be the same as `--input`.

If the `--collapse` flag is used, this will mean that even if there
are multiple nested folder in the input directory, all converted files
will be saved the the root/top-level of the output directory.

Generally, when working with folders that contain older .vsd formatted
Visio project files, the --vsd flag will have to be used first in
order to convert the files to the .vsdx format, which will then be
able to be converted to .pdf using the --pdf flag.
