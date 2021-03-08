import win32com.client
import pathlib
import os
import glob
import time
import argparse
import shutil

def copy_dir(src, dst, *, follow_sym=True):
    if os.path.isdir(dst):
        dst = os.path.join(dst, os.path.basename(src))
    if os.path.isdir(src):
        shutil.copyfile(src, dst, follow_symlinks=follow_sym)
        shutil.copystat(src, dst, follow_symlinks=follow_sym)
    return dst

def convert_files(in_dir: str, out_dir: str, collapse: bool=False) -> None:
    if in_dir == None:
        path = pathlib.Path().absolute()
    else:
        path = pathlib.Path(in_dir).absolute()
    if out_dir == None:
        out_path = pathlib.Path().absolute()
    else:
        out_path = pathlib.Path(out_dir).absolute()
    if collapse and not os.path.exists(str(out_path)):
        os.mkdir(str(out_path))
    print(f"Input directory: {str(path)}")
    print(f"Output directory: {str(out_path)}")
    if not collapse:
        shutil.copytree(path, out_path, dirs_exist_ok=True, copy_function=copy_dir)
    files = glob.glob(str(path) + '/**/*.vsdx', recursive=True)
    visio = win32com.client.Dispatch("Visio.InvisibleApp")
    for f in files:
        try:
            # print(f"Converting {f}")
            doc = visio.Documents.Open(f)
            if collapse:
                prefix = f.split("\\")[-1].split(".")[0]
                fname = str(out_path) + f"\\{prefix}.pdf"
                print(fname)
                doc.ExportAsFixedFormat( 1, fname , 1, 0 )
            else:
                fname = str(out_path) \
                    + "\\" \
                    + str(f).replace(str(os.path.commonprefix([f,out_path])), "").split(".")[0] \
                    + ".pdf"
                doc.ExportAsFixedFormat( 1, fname , 1, 0 )
                print(fname)
            doc.Close()
        except Exception as e:
            doc.Close()
            visio.Quit()
            raise e
    visio.Quit()


def main():
    p = argparse.ArgumentParser(description="Export Visio Files to PDF")
    p.add_argument("--input", "-i", metavar="DIR", type=str, help="The directory to recurisvely search for Visio files.")
    p.add_argument("--output", "-o", metavar="DIR", type=str, help="The directory to save the PDF files.")
    p.add_argument("--collapse", "-c", help="Save files to root of output directory.", action="store_true")
    args = p.parse_args()
    convert_files(in_dir=args.input, out_dir=args.output, collapse=args.collapse)

if __name__ == "__main__":
    main()