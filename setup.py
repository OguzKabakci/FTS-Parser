from cx_Freeze import setup, Executable
import os, sys

if sys.platform == 'win32':
    base = 'Win32GUI'
else:
    base = None

executables = [Executable("Parse_FTS.py", base=base)]

packages = ["docx2txt", "regex", "xml", "os", "tkinter"]
build_exe_options = {
    'packages': packages,
    'include_files': [r'C:\Users\vjkyky\AppData\Local\Programs\Python\Python36-32\DLLs\tcl86t.dll',
                      r'C:\Users\vjkyky\AppData\Local\Programs\Python\Python36-32\DLLs\tk86t.dll'],
    'zip_include_packages': "*",
    'zip_exclude_packages': "",
    'include_msvcr': True,
    'optimize': 2
}

os.environ['TCL_LIBRARY'] = r'C:\Users\vjkyky\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\vjkyky\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'

setup(
    name="<parse_fts>",
    options={"build_exe": build_exe_options},
    version="0.1",
    description='Parses FTS and creates XML and env.',
    executables=executables
)
