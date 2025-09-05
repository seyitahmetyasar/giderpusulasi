import logging
import tkinter as tk

def setup_logger(widget: tk.Text = None, logfile: str = "app.log") -> logging.Logger:
    logger = logging.getLogger("giderpusulasi")
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh = logging.FileHandler(logfile, encoding="utf-8")
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    if widget is not None:
        class TextHandler(logging.Handler):
            def __init__(self, text_widget: tk.Text):
                super().__init__()
                self.widget = text_widget
            def emit(self, record: logging.LogRecord):
                msg = self.format(record)
                self.widget.insert(tk.END, msg + "\n")
                self.widget.see(tk.END)
        th = TextHandler(widget)
        th.setFormatter(formatter)
        logger.addHandler(th)
    return logger
