"""
Claude Word Assistant — macOS menu bar app.
Packaged with PyInstaller for end-user distribution (no Python required).

On first launch:
  - Generates SSL certs
  - Trusts the cert in the macOS keychain
  - Copies the manifest to the Word sideload directory

Then starts the FastAPI server in a background thread and keeps it running.
The user clicks the menu bar icon to start/stop or view logs.
"""

import os
import sys
import subprocess
import threading
import time
import logging
import webbrowser
from pathlib import Path

import rumps

# ---------------------------------------------------------------------------
# Paths — resolved relative to the .app bundle at runtime
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    # Running inside PyInstaller bundle
    BUNDLE_DIR = Path(sys._MEIPASS)
    # Resources are stored next to the executable in Contents/MacOS
    APP_SUPPORT = Path.home() / "Library" / "Application Support" / "ClaudeWordAssistant"
else:
    # Dev mode — run from repo root
    BUNDLE_DIR = Path(__file__).parent.parent
    APP_SUPPORT = BUNDLE_DIR / "_app_support"

CERTS_DIR   = APP_SUPPORT / "certs"
LOGS_DIR    = APP_SUPPORT / "logs"
LOG_FILE    = LOGS_DIR / "server.log"
PID_FILE    = APP_SUPPORT / "server.pid"
ADDIN_DIR   = BUNDLE_DIR / "addin"
CERT_FILE   = CERTS_DIR / "localhost.crt"
KEY_FILE    = CERTS_DIR / "localhost.key"
WEF_DIR     = Path.home() / "Library" / "Containers" / "com.microsoft.Word" / "Data" / "Documents" / "wef"
MANIFEST_SRC = ADDIN_DIR / "manifest.xml"
MANIFEST_DST = WEF_DIR / "claude-word-manifest.xml"


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
def setup_logging():
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    import logging.handlers
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
# First-run setup
# ---------------------------------------------------------------------------
def is_first_run() -> bool:
    return not CERT_FILE.exists()


def run_setup() -> bool:
    """Generate certs and trust them. Returns True on success."""
    CERTS_DIR.mkdir(parents=True, exist_ok=True)

    logger.info("Generating SSL certificate...")
    result = subprocess.run([
        "openssl", "req", "-x509",
        "-newkey", "rsa:4096",
        "-keyout", str(KEY_FILE),
        "-out", str(CERT_FILE),
        "-days", "365",
        "-nodes",
        "-subj", "/CN=localhost",
        "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
    ], capture_output=True)

    if result.returncode != 0:
        logger.error("openssl failed: %s", result.stderr.decode())
        return False

    logger.info("Trusting certificate in keychain...")
    subprocess.run([
        "security", "add-trusted-cert",
        "-d", "-r", "trustRoot",
        "-k", str(Path.home() / "Library" / "Keychains" / "login.keychain-db"),
        str(CERT_FILE),
    ])

    logger.info("Copying manifest to Word sideload directory...")
    WEF_DIR.mkdir(parents=True, exist_ok=True)
    import shutil
    shutil.copy2(str(MANIFEST_SRC), str(MANIFEST_DST))

    logger.info("First-run setup complete.")
    return True


# ---------------------------------------------------------------------------
# Server thread
# ---------------------------------------------------------------------------
server_process = None
server_lock = threading.Lock()


def start_server():
    global server_process
    with server_lock:
        if server_process and server_process.poll() is None:
            return  # already running

        # Locate the bundled server binary or use uvicorn from venv
        if getattr(sys, "frozen", False):
            server_bin = BUNDLE_DIR / "server_bin"
            cmd = [str(server_bin)]
        else:
            # Dev mode — use venv uvicorn
            venv_uvicorn = Path(__file__).parent / "venv" / "bin" / "uvicorn"
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

        logger.info("Starting server: %s", " ".join(cmd))
        log_fd = open(LOG_FILE, "a")
        server_process = subprocess.Popen(
            cmd,
            stdout=log_fd,
            stderr=log_fd,
            env=env,
            cwd=str(BUNDLE_DIR / "server") if not getattr(sys, "frozen", False) else str(BUNDLE_DIR),
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


def find_claude_cli() -> str:
    candidates = [
        "/opt/homebrew/bin/claude",
        "/usr/local/bin/claude",
        str(Path.home() / ".local" / "bin" / "claude"),
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    # Last resort: which
    result = subprocess.run(["which", "claude"], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip()
    return "claude"


# ---------------------------------------------------------------------------
# Menu bar app
# ---------------------------------------------------------------------------
class ClaudeWordApp(rumps.App):
    def __init__(self):
        super().__init__(
            name="Claude Word Assistant",
            title="✦",           # menu bar icon (text — no .icns needed)
            menu=[
                rumps.MenuItem("Démarrer le serveur", callback=self.toggle_server),
                rumps.separator,
                rumps.MenuItem("Ouvrir les logs", callback=self.open_logs),
                rumps.MenuItem("Réinstaller le certificat", callback=self.reinstall_cert),
                rumps.separator,
                rumps.MenuItem("Quitter", callback=rumps.quit_application),
            ],
            quit_button=None,
        )
        self._toggle_item = self.menu["Démarrer le serveur"]
        # Status polling timer
        self._timer = rumps.Timer(self._poll_status, 3)
        self._timer.start()

    # ------------------------------------------------------------------
    def toggle_server(self, _):
        if is_server_running():
            stop_server()
            self._set_stopped()
        else:
            start_server()
            self._set_running()

    def open_logs(self, _):
        if LOG_FILE.exists():
            subprocess.run(["open", "-a", "Console", str(LOG_FILE)])
        else:
            rumps.alert("Aucun log disponible pour le moment.")

    def reinstall_cert(self, _):
        CERT_FILE.unlink(missing_ok=True)
        KEY_FILE.unlink(missing_ok=True)
        ok = run_setup()
        if ok:
            rumps.alert("Certificat réinstallé avec succès.")
        else:
            rumps.alert("Échec de la réinstallation. Consultez les logs.")

    # ------------------------------------------------------------------
    def _poll_status(self, _):
        if is_server_running():
            self._set_running()
        else:
            self._set_stopped()

    def _set_running(self):
        self.title = "✦"
        self._toggle_item.title = "Arrêter le serveur"

    def _set_stopped(self):
        self.title = "✧"
        self._toggle_item.title = "Démarrer le serveur"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    setup_logging()
    APP_SUPPORT.mkdir(parents=True, exist_ok=True)

    if is_first_run():
        logger.info("First run detected — running setup...")
        response = rumps.alert(
            title="Claude Word Assistant",
            message=(
                "Bienvenue !\n\n"
                "La configuration initiale va :\n"
                "• Générer un certificat SSL local\n"
                "• L'ajouter à votre trousseau macOS (mot de passe requis)\n"
                "• Enregistrer le plugin dans Word\n\n"
                "Cliquez OK pour continuer."
            ),
            ok="OK",
            cancel="Annuler",
        )
        if response == 1:
            ok = run_setup()
            if not ok:
                rumps.alert("La configuration a échoué. Consultez les logs.")
                return
            rumps.alert(
                title="Configuration terminée",
                message=(
                    "Tout est prêt !\n\n"
                    "1. Ouvrez Word\n"
                    "2. Insertion → Compléments → Mes compléments\n"
                    "3. Cliquez sur Claude Assistant\n\n"
                    "Le serveur va démarrer automatiquement."
                ),
            )

    start_server()
    ClaudeWordApp().run()
    stop_server()


if __name__ == "__main__":
    main()
