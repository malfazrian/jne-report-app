import sys
from datetime import datetime

class DualLogger:
    def __init__(self, filepath):
        self.log = open(filepath, "a", encoding="utf-8")
        self.log.write(f"\n\n=== Log started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")

        # Cek apakah stdout asli ada (kalau pythonw biasanya None)
        self.terminal = getattr(sys, '__stdout__', None)

    def write(self, message):
        # Tulis ke log file
        try:
            self.log.write(message)
        except Exception:
            pass

        # Kalau stdout ada (mode python biasa), tulis juga ke terminal
        if self.terminal:
            try:
                self.terminal.write(message)
            except Exception:
                pass

    def flush(self):
        try:
            self.log.flush()
        except Exception:
            pass
        if self.terminal:
            try:
                self.terminal.flush()
            except Exception:
                pass
