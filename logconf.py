import logging
import logging.handlers
import queue
from typing import Optional

_gui_queue: "queue.Queue[str]" = queue.Queue()

class QueueHandler(logging.Handler):
    """Logging handler that writes messages to a queue."""
    def __init__(self, q: "queue.Queue[str]") -> None:
        super().__init__()
        self.queue = q

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self.queue.put_nowait(msg)
        except Exception:
            # Best effort; never raise in logging
            pass

def setup_logger(
    name: str = "giderpusulasi",
    log_file: str = "app.log",
    level: int = logging.INFO,
    max_bytes: int = 5 * 1024 * 1024,
    backup_count: int = 3,
    gui_queue: Optional["queue.Queue[str]"] = None,
) -> logging.Logger:
    """Configure and return application logger.

    Adds console, rotating file and queue handlers while avoiding duplicates.
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    def _handler_exists(cls: type) -> bool:
        return any(isinstance(h, cls) for h in logger.handlers)

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    if not _handler_exists(logging.StreamHandler):
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)

    if not _handler_exists(logging.handlers.RotatingFileHandler):
        fh = logging.handlers.RotatingFileHandler(
            log_file, maxBytes=max_bytes, backupCount=backup_count, encoding="utf-8"
        )
        fh.setFormatter(fmt)
        logger.addHandler(fh)

    q = gui_queue or _gui_queue
    if q and not _handler_exists(QueueHandler):
        qh = QueueHandler(q)
        qh.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logger.addHandler(qh)

    return logger

def get_gui_queue() -> "queue.Queue[str]":
    """Return the queue used by QueueHandler for GUI logging."""
    return _gui_queue
