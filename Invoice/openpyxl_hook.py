
import os
import sys

def _append_openpyxl_path():
    openpyxl_path = os.path.join(sys._MEIPASS, "openpyxl")
    if openpyxl_path not in sys.path:
        sys.path.append(openpyxl_path)
    et_xmlfile_path = os.path.join(sys._MEIPASS, "et_xmlfile")
    if et_xmlfile_path not in sys.path:
        sys.path.append(et_xmlfile_path)

_append_openpyxl_path()
