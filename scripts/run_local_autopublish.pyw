"""
run_local_autopublish.pyw

Runs the publish flow silently (no console window) -- this is what
Windows Task Scheduler calls every minute.

Logs to local_autopublish.log next to this file so you can check
history without a console window ever appearing.
"""

import datetime
import os
import subprocess
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(SCRIPT_DIR, "local_autopublish.log")
PUBLISH_SCRIPT = os.path.join(SCRIPT_DIR, "publish_dashboard_data.py")

# Belt-and-braces: prevent this subprocess from flashing its own console
# window, even though it's normally launched via pythonw.exe already.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def log(message):
    timestamp = datetime.datetime.now().isoformat(timespec="seconds")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def main():
    log("Run started")
    try:
        result = subprocess.run(
            [sys.executable, PUBLISH_SCRIPT],
            capture_output=True,
            text=True,
            timeout=180,
            creationflags=NO_WINDOW,
        )
        if result.stdout:
            log(result.stdout.strip())
        if result.stderr:
            log("STDERR: " + result.stderr.strip())
        log(f"Run finished with exit code {result.returncode}")
    except Exception as exc:
        log(f"Run failed with exception: {exc}")


if __name__ == "__main__":
    main()
