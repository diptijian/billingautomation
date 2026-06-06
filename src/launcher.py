import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

import streamlit.web.cli as stcli

# Force PyInstaller to include these runtime dependencies/modules
import pandas  # noqa: F401
import openpyxl  # noqa: F401
import pyxlsb  # noqa: F401
import xlsxwriter  # noqa: F401
import t1_generate_cd  # noqa: F401
import t2_generate_e_base  # noqa: F401
import t3_generate_x  # noqa: F401


def resource_path(relative_path: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent / relative_path


def find_free_port() -> int:
    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    sock.close()
    return port


def wait_and_open_browser(port: int, timeout: int = 60):
    url = f"http://localhost:{port}"

    start = time.time()

    while time.time() - start < timeout:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                webbrowser.open(url)
                return
        except OSError:
            time.sleep(1)


def main():
    app_path = resource_path("app.py")
    port = find_free_port()

    browser_thread = threading.Thread(
        target=wait_and_open_browser,
        args=(port,),
        daemon=True,
    )
    browser_thread.start()

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        f"--server.port={port}",
        "--global.developmentMode=false",
    ]

    sys.exit(stcli.main())


if __name__ == "__main__":
    main()