import sys
import types

# Stub modules imported by arama23 but absent in the test environment.
modules = {
    "requests": types.ModuleType("requests"),
    "requests.adapters": types.ModuleType("requests.adapters"),
    "urllib3": types.ModuleType("urllib3"),
    "urllib3.util": types.ModuleType("urllib3.util"),
    "urllib3.util.retry": types.ModuleType("urllib3.util.retry"),
    "openpyxl": types.ModuleType("openpyxl"),
    "openpyxl.utils": types.ModuleType("openpyxl.utils"),
    "tkinter": types.ModuleType("tkinter"),
    "tkinter.ttk": types.ModuleType("tkinter.ttk"),
    "tkinter.filedialog": types.ModuleType("tkinter.filedialog"),
    "tkinter.messagebox": types.ModuleType("tkinter.messagebox"),
    "tkinter.scrolledtext": types.ModuleType("tkinter.scrolledtext"),
}

# minimal stand-ins for used classes/functions
modules["requests"].Session = object
modules["requests.adapters"].HTTPAdapter = object
modules["urllib3.util.retry"].Retry = object
modules["tkinter"].Tk = object

openpyxl = modules["openpyxl"]
class Workbook:  # pragma: no cover - simple stub
    pass

def load_workbook(*args, **kwargs):  # pragma: no cover - simple stub
    return Workbook()
openpyxl.Workbook = Workbook
openpyxl.load_workbook = load_workbook

utils = modules["openpyxl.utils"]
def get_column_letter(i):  # pragma: no cover - simple stub
    return str(i)
utils.get_column_letter = get_column_letter

for name, module in modules.items():
    sys.modules.setdefault(name, module)
