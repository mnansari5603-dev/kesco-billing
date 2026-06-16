from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import time

WATCH_FILE = r"C:\Users\Admin\Downloads\KESCO\FINAL SUMMARY.xlsx.xlsx"
DASHBOARD_SCRIPT = r"C:\Users\Admin\Downloads\PAYTHON AUTOMATION\dashboard.py"

class MyHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path == WATCH_FILE:
            print("Excel Updated... Refreshing Dashboard")
            subprocess.run(["python", DASHBOARD_SCRIPT])

event_handler = MyHandler()
observer = Observer()

observer.schedule(
    event_handler,
    path=r"C:\Users\Admin\Downloads\KESCO",
    recursive=False
)

observer.start()

print("Auto Refresh Started...")

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()