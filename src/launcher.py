import subprocess
import webbrowser
import time
import sys
from pathlib import Path

APP_PATH = Path(__file__).parent / "app.py"

subprocess.Popen([
    sys.executable,
    "-m",
    "streamlit",
    "run",
    str(APP_PATH),
    "--server.headless=true"
])

time.sleep(3)

webbrowser.open("http://localhost:8501")