"""
publish_dashboard_data.py

1. Fetches/rebases the repo (fails loudly if the repo is dirty -- keep it
   clean or auto-publish stops, same as the original Operations pipeline)
2. Rebuilds dashboard_data.json via refresh_dashboard_data.py
3. Commits if the JSON changed
4. Pushes to origin/main

Run this directly to test, or let run_local_autopublish.pyw call it on
a schedule.
"""

import subprocess
import sys
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# On Windows, prevent each subprocess call (git, python) from briefly
# flashing its own console window when this script is run via pythonw.exe.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def run(cmd, **kwargs):
    print(f"[publish] $ {' '.join(cmd)}")
    result = subprocess.run(
        cmd, cwd=REPO_ROOT, capture_output=True, text=True, creationflags=NO_WINDOW, **kwargs
    )
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.stderr.strip():
        print(result.stderr.strip(), file=sys.stderr)
    return result


def dirty_files():
    result = run(["git", "status", "--porcelain"])
    files = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        files.append(line[3:].strip())
    return files


def repo_is_dirty():
    return bool(dirty_files())


def main():
    files = dirty_files()
    unexpected = [f for f in files if f != "dashboard_data.json"]
    if unexpected:
        print(
            "[publish] ERROR: repo has uncommitted changes outside dashboard_data.json: "
            f"{unexpected}. Commit or stash them -- auto-publish will not rebase a dirty repo."
        )
        sys.exit(1)

    if "dashboard_data.json" in files:
        print(
            "[publish] dashboard_data.json has pending changes from an earlier manual run "
            "-- committing them before continuing."
        )
        run(["git", "add", "dashboard_data.json"])
        run(["git", "commit", "-m", "Auto-publish: commit pending dashboard_data.json changes"])

    fetch = run(["git", "fetch", "origin"])
    if fetch.returncode != 0:
        print("[publish] ERROR: git fetch failed.")
        sys.exit(1)

    rebase = run(["git", "rebase", "origin/main"])
    if rebase.returncode != 0:
        print("[publish] ERROR: rebase failed (conflicts?). Resolve manually.")
        sys.exit(1)

    refresh = subprocess.run(
        [sys.executable, os.path.join(REPO_ROOT, "scripts", "refresh_dashboard_data.py")],
        cwd=REPO_ROOT,
        creationflags=NO_WINDOW,
    )
    if refresh.returncode != 0:
        print("[publish] ERROR: refresh_dashboard_data.py failed.")
        sys.exit(1)

    if not repo_is_dirty():
        print("[publish] No changes to dashboard_data.json -- nothing to publish.")
        return

    run(["git", "add", "dashboard_data.json"])
    run(["git", "commit", "-m", "Auto-publish: update dashboard_data.json"])
    push = run(["git", "push", "origin", "main"])
    if push.returncode != 0:
        print("[publish] ERROR: push failed.")
        sys.exit(1)

    print("[publish] Published successfully.")


if __name__ == "__main__":
    main()
