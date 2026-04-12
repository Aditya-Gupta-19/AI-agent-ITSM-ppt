import os
import threading
import time
from typing import Callable, Optional

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer


class ExcelFileWatcher:
    """
    Watches a specific Excel file for modification events.
    On change: waits 3 seconds (file write buffer), then calls callback.
    Debounce: minimum 30 seconds between triggers (avoid double-fire on save).
    """

    def __init__(self, file_path: str, callback: Callable[[str], None]):
        self.file_path = file_path
        self.callback = callback
        self.last_triggered = 0.0
        self.debounce_seconds = 30
        self._observer: Optional[Observer] = None
        self._handler: Optional[FileSystemEventHandler] = None

    def start(self):
        """Start watching. Blocks until KeyboardInterrupt."""

        parent_dir = os.path.dirname(os.path.abspath(self.file_path))
        target_abs = os.path.abspath(self.file_path)

        class Handler(FileSystemEventHandler):
            def __init__(self, outer: "ExcelFileWatcher"):
                self.outer = outer

            def on_modified(self, event):
                if event.is_directory:
                    return
                try:
                    src = os.path.abspath(event.src_path)
                except Exception:
                    return
                if src != target_abs:
                    return

                now = time.time()
                if now - self.outer.last_triggered < self.outer.debounce_seconds:
                    return
                self.outer.last_triggered = now

                def worker():
                    # Buffer for Excel save/write.
                    time.sleep(3)
                    self.outer.callback(self.outer.file_path)

                threading.Thread(target=worker, daemon=True).start()

        self._handler = Handler(self)
        self._observer = Observer()
        self._observer.schedule(self._handler, parent_dir, recursive=False)
        self._observer.start()

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()

    def stop(self):
        """Clean shutdown."""
        if self._observer:
            self._observer.stop()
            self._observer.join(timeout=5)

