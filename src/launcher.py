import subprocess
import webbrowser
import time
import socket
import sys
from pathlib import Path

APP_PATH = Path(__file__).parent / "app.py"

subprocess.Popen([
    sys.executable,
    "-m",
    "streamlit",
    "run",
    str(APP_PATH),
    "--server.headless=true",
    "--browser.gatherUsageStats=false"
])

def wait_for_server(host="localhost", port=8501, timeout=30):
    start = time.time()

    while time.time() - start < timeout:
        try:
            with socket.create_connection((host, port), timeout=1):
                return True
        except:
            time.sleep(1)

    return False

if wait_for_server():
    webbrowser.open("http://localhost:8501")