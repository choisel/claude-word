# Claude Word Assistant

Interact with Claude directly from Microsoft Word — selected text, document analysis, section navigation — via a sidebar panel.

---

## For users (no technical knowledge required)

### Prerequisites

- macOS
- Microsoft Word (desktop)
- [Claude CLI](https://claude.ai/download) installed and logged in

> To check if Claude CLI is installed, open Terminal and type `claude --version`. If you see a version number, you're good.

### Installation

1. Download `Claude Word Assistant.zip` from the [latest release](../../releases/latest)
2. Unzip it — you get `Claude Word Assistant.app`
3. Move it to your **Applications** folder
4. Double-click to launch it

On first launch:
- A setup wizard will run automatically
- macOS will ask for your **password** once (to trust the local SSL certificate)
- The plugin will be registered in Word automatically

### Using the plugin

1. The **✦** icon appears in your menu bar — click it to start/stop the server
2. Open Word
3. Go to **Insert → Add-ins → My Add-ins** → click **Claude Assistant**
4. The sidebar opens on the right

**What you can do:**
- Select text in your document → click "Refresh" → ask Claude about it
- Load the full document → Claude analyses its structure
- Type a section number (e.g. `1.2` or `Article 3`) to ask about a specific section
- The current section is detected automatically as you move through the document

### Troubleshooting

- **Server won't start** — make sure Claude CLI is installed (`claude --version` in Terminal)
- **Plugin not visible in Word** — try Insert → Add-ins → Refresh
- **Logs** — click ✦ in the menu bar → "Open logs"
- **Certificate error** — click ✦ → "Reinstall certificate"

---

## For developers

### Architecture

```
Word (Office.js taskpane)
    ↓ HTTPS POST
FastAPI server (uvicorn + local SSL)
    ↓ stdin/stdout
Claude CLI
```

### Requirements

- Python 3.8+
- [Claude CLI](https://claude.ai/download) installed and authenticated

### Setup

```bash
git clone <repo>
cd claude-word
bash scripts/setup.sh
```

This will:
- Create a Python virtualenv in `server/venv/`
- Install dependencies
- Generate a self-signed SSL certificate
- Trust it in your macOS keychain
- Copy the manifest to Word's sideload directory

### Running

```bash
bash start.sh        # start server in background
bash stop.sh         # stop server
tail -f server/logs/server.log   # follow logs
```

### Project structure

```
claude-word/
├── addin/               # Office.js taskpane (HTML/CSS/JS)
│   ├── taskpane.html
│   ├── taskpane.css
│   ├── taskpane.js
│   ├── manifest.xml     # Word add-in manifest
│   └── assets/          # icons
├── server/              # Python FastAPI server
│   ├── main.py          # API endpoints (/init, /ask, /ask-stream, /health)
│   ├── document.py      # Document parsing and context building
│   ├── session.py       # In-memory session store
│   ├── app_launcher.py  # macOS menu bar app (for distribution)
│   └── requirements.txt
├── scripts/
│   ├── setup.sh         # Developer one-time setup
│   └── build-app.sh     # Build standalone .app for distribution
├── start.sh             # Start server (background)
└── stop.sh              # Stop server
```

### API

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Server health check |
| `/init` | POST | Load document, extract sections, generate summary if needed |
| `/ask` | POST | Ask Claude (non-streaming) |
| `/ask-stream` | POST | Ask Claude (streaming SSE) |

### Building the distributable .app

```bash
bash scripts/build-app.sh
```

Output: `dist/Claude Word Assistant.app`

To publish a release:
```bash
cd dist
zip -r "Claude Word Assistant.zip" "Claude Word Assistant.app"
# Upload the .zip to a GitHub Release
```

### Configuration

Copy `server/.env.example` to `server/.env` to override defaults:

```ini
PORT=5000
CLAUDE_CLI_PATH=/opt/homebrew/bin/claude
LOG_LEVEL=INFO
CLAUDE_TIMEOUT_SECONDS=120
```

### Phase roadmap

- [x] Phase 1 — Selected text → Claude → sidebar response
- [x] Phase 2 — Full document loading, section navigation, streaming
- [ ] Phase 3 — Direct document rewriting via Claude
