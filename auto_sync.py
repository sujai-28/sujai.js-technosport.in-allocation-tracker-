import os
import sys
import time
import subprocess
from datetime import datetime

GIT_EXE = r"C:\Program Files\Git\cmd\git.exe"
if not os.path.exists(GIT_EXE):
    GIT_EXE = "git"

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

def run_git(args):
    cmd = [GIT_EXE] + args
    result = subprocess.run(cmd, cwd=PROJECT_DIR, capture_output=True, text=True)
    return result.returncode, result.stdout, result.stderr

def sync_now():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Checking for file changes...")
    
    # Check git status
    code, stdout, stderr = run_git(["status", "--porcelain"])
    if code != 0:
        print("Git status check failed:", stderr)
        return False
        
    if not stdout.strip():
        print("No changes detected.")
        return True
        
    print("Changes detected! Staging files...")
    run_git(["add", "."])
    
    commit_msg = f"Auto-update data and files [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]"
    print(f"Committing: {commit_msg}")
    run_git(["commit", "-m", commit_msg])
    
    print("Pushing to GitHub (main branch)...")
    code, stdout, stderr = run_git(["push", "origin", "main"])
    if code == 0:
        print("Successfully synced and pushed to GitHub! Render will auto-update your live app.")
        return True
    else:
        print("Push failed:", stderr or stdout)
        return False

def watch_loop(interval=15):
    print("=" * 60)
    print("  Auto-Sync Service Started")
    print(f"  Monitoring '{PROJECT_DIR}' for changes every {interval}s")
    print("  Press Ctrl+C to stop")
    print("=" * 60)
    
    try:
        while True:
            sync_now()
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\nAuto-Sync service stopped.")

if __name__ == "__main__":
    if "--once" in sys.argv:
        sync_now()
    else:
        watch_loop(interval=15)
