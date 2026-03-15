"""
Claude Word Assistant — Windows system tray app.
Packaged with PyInstaller for end-user distribution (no Python required).

On first launch:
  - Generates SSL certs (requires openssl.exe from Git for Windows or PATH)
  - Trusts the cert in the Windows CurrentUser Root store (no UAC required)
  - Registers the add-in folder as a Trusted Catalog in the Windows registry

Then starts the FastAPI server in a background thread and keeps it running.
The user clicks the tray icon to start/stop or view logs.

After first-run setup, the user must:
  1. Restart Word
  2. Insert → Add-ins → My Add-ins → Shared Folder → Claude Assistant
"""

import ctypes
import logging
import logging.handlers
import os
import shutil
import subprocess
import sys
import threading
import time
import winreg
from pathlib import Path

import pystray
from PIL import Image

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BUNDLE_DIR = Path(sys._MEIPASS)
    APP_SUPPORT = Path(os.environ["APPDATA"]) / "ClaudeWordAssistant"
else:
    BUNDLE_DIR = Path(__file__).parent.parent
    APP_SUPPORT = BUNDLE_DIR / "_app_support"

CERTS_DIR    = APP_SUPPORT / "certs"
LOGS_DIR     = APP_SUPPORT / "logs"
LOG_FILE     = LOGS_DIR / "server.log"
ADDIN_DIR    = BUNDLE_DIR / "addin"
CERT_FILE    = CERTS_DIR / "localhost.crt"
KEY_FILE     = CERTS_DIR / "localhost.key"
MANIFEST_SRC = ADDIN_DIR / "manifest.xml"

# Stable GUID for this app's Trusted Catalog registry entry
CATALOG_GUID = "{a7c2e4f0-3b1d-4abc-8def-c0ffee000002}"


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging():
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)-8s %(name)s - %(message)s")
    fh = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(fh)


logger = logging.getLogger("launcher")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def find_claude_cli() -> str:
    candidates = [
        Path(os.environ.get("APPDATA", "")) / "Claude" / "claude.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "claude" / "claude.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Claude" / "claude.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Claude" / "claude.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Claude" / "claude.exe",
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    result = subprocess.run(["where", "claude"], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip().splitlines()[0]
    return "claude"


def find_openssl() -> str:
    candidates = [
        str(BUNDLE_DIR / "openssl.exe"),                                      # bundled
        r"C:\Program Files\Git\usr\bin\openssl.exe",                          # Git for Windows
        r"C:\Program Files (x86)\Git\usr\bin\openssl.exe",
        str(Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "OpenSSL" / "bin" / "openssl.exe"),
        r"C:\ProgramData\chocolatey\bin\openssl.exe",
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    result = subprocess.run(["where", "openssl"], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip().splitlines()[0]
    return "openssl"


def _msgbox(message: str, title: str = "Claude Word Assistant", ok_cancel: bool = False) -> bool:
    """Show a native Win32 message box. Returns True if user clicked OK."""
    flags = 1 if ok_cancel else 0  # MB_OKCANCEL or MB_OK
    result = ctypes.windll.user32.MessageBoxW(0, message, title, flags)
    return result == 1


# ---------------------------------------------------------------------------
# First-run setup
# ---------------------------------------------------------------------------
def is_first_run() -> bool:
    return not CERT_FILE.exists()


def run_setup() -> bool:
    CERTS_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Generate SSL certificate
    logger.info("Generating SSL certificate...")
    openssl = find_openssl()
    result = subprocess.run([
        openssl, "req", "-x509",
        "-newkey", "rsa:4096",
        "-keyout", str(KEY_FILE),
        "-out", str(CERT_FILE),
        "-days", "365",
        "-nodes",
        "-subj", "/CN=localhost",
        "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
    ], capture_output=True)

    if result.returncode != 0:
        logger.error("openssl failed: %s", result.stderr.decode(errors="replace"))
        return False
    logger.info("SSL certificate generated.")

    # 2. Trust cert in CurrentUser Root store (no UAC required)
    logger.info("Trusting certificate in Windows certificate store...")
    result = subprocess.run(
        ["certutil", "-addstore", "-user", "Root", str(CERT_FILE)],
        capture_output=True,
    )
    if result.returncode != 0:
        logger.warning("certutil -user failed (%s), trying machine store via UAC...",
                       result.stderr.decode(errors="replace"))
        subprocess.run([
            "powershell", "-Command",
            f'Start-Process certutil -ArgumentList \'-addstore Root "{CERT_FILE}"\' -Verb RunAs -Wait',
        ])
    logger.info("Certificate trusted.")

    # 3. Register add-in folder as Trusted Catalog in registry
    logger.info("Registering Word add-in catalog...")
    _register_catalog()

    logger.info("First-run setup complete.")
    return True


def _register_catalog():
    key_path = rf"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{CATALOG_GUID}"
    try:
        key = winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "Url",   0, winreg.REG_SZ,    str(ADDIN_DIR))
        winreg.SetValueEx(key, "Flags", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        logger.info("Catalog registered: %s → %s", key_path, ADDIN_DIR)
    except OSError as e:
        logger.error("Failed to write registry: %s", e)


def _remove_catalog():
    key_path = rf"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{CATALOG_GUID}"
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
    except FileNotFoundError:
        pass


# ---------------------------------------------------------------------------
# Server management
# ---------------------------------------------------------------------------
server_process = None
server_lock = threading.Lock()


def start_server():
    global server_process
    with server_lock:
        if server_process and server_process.poll() is None:
            return

        if getattr(sys, "frozen", False):
            cmd = [str(BUNDLE_DIR / "server_bin.exe")]
        else:
            venv_uvicorn = Path(__file__).parent / "venv" / "Scripts" / "uvicorn.exe"
            cmd = [
                str(venv_uvicorn), "main:app",
                "--host", "0.0.0.0",
                "--port", "5000",
            ]

        cmd += [
            "--ssl-keyfile", str(KEY_FILE),
            "--ssl-certfile", str(CERT_FILE),
            "--log-level", "info",
        ]

        env = os.environ.copy()
        env["CLAUDE_CLI_PATH"] = find_claude_cli()
        env["LOG_LEVEL"] = "INFO"
        env["CLAUDE_TIMEOUT_SECONDS"] = "120"

        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        log_fd = open(LOG_FILE, "a")
        cwd = str(BUNDLE_DIR) if getattr(sys, "frozen", False) else str(Path(__file__).parent)

        logger.info("Starting server: %s", " ".join(cmd))
        server_process = subprocess.Popen(
            cmd,
            stdout=log_fd,
            stderr=log_fd,
            env=env,
            cwd=cwd,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.info("Server started (PID %d)", server_process.pid)


def stop_server():
    global server_process
    with server_lock:
        if server_process and server_process.poll() is None:
            server_process.terminate()
            try:
                server_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                server_process.kill()
            logger.info("Server stopped.")
        server_process = None


def is_server_running() -> bool:
    with server_lock:
        return server_process is not None and server_process.poll() is None


# ---------------------------------------------------------------------------
# System tray
# ---------------------------------------------------------------------------
def _load_icon() -> Image.Image:
    for candidate in [
        ADDIN_DIR / "assets" / "icon-32.png",
        BUNDLE_DIR / "assets" / "icon-32.png",
    ]:
        if candidate.exists():
            return Image.open(str(candidate)).convert("RGBA")
    # Fallback: plain coloured square
    return Image.new("RGBA", (32, 32), (204, 120, 92, 255))


def _build_menu() -> pystray.Menu:
    running = is_server_running()
    toggle_label = "Stop server" if running else "Start server"
    return pystray.Menu(
        pystray.MenuItem(toggle_label, _on_toggle),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open logs", _on_open_logs),
        pystray.MenuItem("Reinstall certificate", _on_reinstall),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", _on_quit),
    )


def _on_toggle(icon, item):
    if is_server_running():
        stop_server()
    else:
        start_server()
    icon.menu = _build_menu()
    icon.update_menu()


def _on_open_logs(icon, item):
    if LOG_FILE.exists():
        os.startfile(str(LOG_FILE))
    else:
        icon.notify("No logs available yet.", title="Claude Word Assistant")


def _on_reinstall(icon, item):
    CERT_FILE.unlink(missing_ok=True)
    KEY_FILE.unlink(missing_ok=True)
    _remove_catalog()
    ok = run_setup()
    msg = "Certificate reinstalled." if ok else "Reinstall failed — check logs."
    icon.notify(msg, title="Claude Word Assistant")


def _on_quit(icon, item):
    stop_server()
    icon.stop()


def _poll_status(icon: pystray.Icon):
    while icon.visible:
        icon.menu = _build_menu()
        icon.update_menu()
        time.sleep(3)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    setup_logging()
    APP_SUPPORT.mkdir(parents=True, exist_ok=True)

    if is_first_run():
        logger.info("First run — running setup...")
        confirmed = _msgbox(
            "Welcome!\n\n"
            "First-time setup will:\n"
            "• Generate a local SSL certificate\n"
            "• Trust it in your Windows certificate store\n"
            "• Register the add-in with Word\n\n"
            "After setup, restart Word and go to:\n"
            "Insert → Add-ins → My Add-ins → Shared Folder\n\n"
            "Click OK to continue.",
            ok_cancel=True,
        )
        if not confirmed:
            return

        ok = run_setup()
        if not ok:
            _msgbox("Setup failed. Check logs for details.")
            return

        _msgbox(
            "Setup complete!\n\n"
            "1. Restart Word\n"
            "2. Insert → Add-ins → My Add-ins → Shared Folder\n"
            "3. Check 'Show in Menu' next to Claude Assistant → Close\n"
            "4. The add-in will appear in Insert → Add-ins\n\n"
            "The server starts automatically when you launch this app."
        )

    start_server()

    icon = pystray.Icon(
        "ClaudeWordAssistant",
        icon=_load_icon(),
        title="Claude Word Assistant",
        menu=_build_menu(),
    )

    threading.Thread(target=_poll_status, args=(icon,), daemon=True).start()
    icon.run()
    stop_server()


if __name__ == "__main__":
    main()
