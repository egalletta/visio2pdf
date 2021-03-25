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

def convert_to_vsdm(in_dir: str, out_dir: str, collapse: bool=False) -> None:
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
    files = glob.glob(str(path) + '/**/*.vsd', recursive=True)
    visio = win32com.client.gencache.EnsureDispatch("Visio.Application")
    for f in files:
        try:
            print(f"Converting {f}")
            doc = visio.Documents.Open(f)
            if collapse:
                prefix = f.split("\\")[-1].split(".")[0]
                fname = str(out_path) + f"\\{prefix}.vsdm"
                print(f" Saved {fname}")
                doc.SaveAs(fname)
            else:
                fname = str(out_path) \
                    + "\\" \
                    + str(f).replace(str(os.path.commonprefix([f,out_path])), "").split(".")[0] \
                    + ".vsdm"
                doc.SaveAs(fname)
                print(fname)
            doc.Close()
        except Exception as e:
            print(e)
    visio.Quit()

def convert_pdf(in_dir: str, out_dir: str, collapse: bool=False) -> None:
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
    files = glob.glob(str(path) + '/**/*.vsd[m,x]', recursive=True)
    visio = win32com.client.gencache.EnsureDispatch("Visio.Application")
    for f in files:
        try:
            print(f"Converting {f}")
            doc = visio.Documents.Open(f)
            if collapse:
                prefix = f.split("\\")[-1].split(".")[0]
                fname = str(out_path) + f"\\{prefix}-exported.pdf"
                print(fname)
                doc.ExportAsFixedFormat( 1, fname , 1, 0 )
            else:
                fname = str(out_path) \
                    + "\\" \
                    + str(f).replace(str(os.path.commonprefix([f,out_path])), "").split(".")[0] \
                    + "-exported.pdf"
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
    p.add_argument("--input", "-i", metavar="DIR", type=str, help="The directory to recurisvely search for Visio files (Defaults to current directory).")
    p.add_argument("--output", "-o", metavar="DIR", type=str, help="The directory to save the exported files. (Defaults to current directory)")
    p.add_argument("--collapse", "-c", help="Save files to root of output directory / do not preserve folder structure", action="store_true")
    p.add_argument("--vsd", "-v", help="Convert all .vsd files in input directory to .vsdx format, and save in output directory", action="store_true")
    p.add_argument("--pdf", "-p", help="Convert all .vsdx files in input directory to PDF, and save in output directory", action="store_true")
    args = p.parse_args()
    
    if not args.pdf and not args.vsd:
        print("\n\nNeed at least one of the following options:\n")
        print("\t--vsd, -v             Convert all .vsd files to newer .vsdx files\n\t--pdf, -p             Convert all .vsdx files to PDF\n\n")
        p.print_help()
    
    if args.vsd:
        convert_to_vsdm(in_dir=args.input, out_dir=args.output, collapse=args.collapse)
    if args.pdf:
        convert_pdf(in_dir=args.input, out_dir=args.output, collapse=args.collapse)

if __name__ == "__main__":
    main()