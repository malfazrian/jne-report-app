import sys
from datetime import datetime

class DualLogger:
    def __init__(self, filepath):
        self.terminal = sys.__stdout__
        self.log = open(filepath, "a", encoding="utf-8")
        self.write(f"\n\n=== Log started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()