# Setup Guide

Step-by-step instructions to deploy the OneNote MCP Server with Docker Compose.

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| **Docker** | Docker Engine with Compose V2 (`docker compose` command) |
| **A Microsoft account** | The account that owns the OneNote notebooks you want to access |
| **A web browser** | Needed once for the authentication step |

> **You do NOT need Node.js, npm, or any other runtime.** Everything runs inside the Docker container.

---

## Step 1 — Get Your Azure Credentials

The server needs an **Application (client) ID** from Microsoft Entra ID (formerly Azure AD) to authenticate. Follow these steps to create one for free.

### 1.1 Open Microsoft Entra Admin Center

Go to: **<https://entra.microsoft.com>**

Sign in with the same Microsoft account that owns your OneNote notebooks.

### 1.2 Register a New Application

1. In the left sidebar, expand **Identity** → **Applications** → click **App registrations**.
   - Direct link: <https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade>
2. Click **+ New registration** (top of the page).
3. Fill in:
   - **Name:** `OneNote MCP Server` (or anything you like)
   - **Supported account types:** Select **"Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"**
   - **Redirect URI:** Leave blank — not needed for device code flow
4. Click **Register**.

### 1.3 Copy the Application (Client) ID

After registration you land on the app's **Overview** page. Copy the value labeled:

```
Application (client) ID:  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

Save this — you'll need it in Step 2.

### 1.4 Enable Public Client Flow

This is required for the device code authentication to work:

1. In the left sidebar of your app, click **Authentication**.
2. Scroll to the bottom to the section **Advanced settings**.
3. Set **"Allow public client flows"** to **Yes**.
4. Click **Save** (top of the page).

### 1.5 Add API Permissions

1. In the left sidebar, click **API permissions**.
2. Click **+ Add a permission**.
3. Select **Microsoft Graph** → **Delegated permissions**.
4. Search for and check each of these:
   - `Notes.Read`
   - `Notes.ReadWrite`
   - `Notes.Create`
   - `User.Read` (usually already added by default)
5. Click **Add permissions**.

> **Admin consent:** If you're using a personal Microsoft account, no admin consent is needed — you grant consent yourself during the device code flow. If you're on a work/school tenant, an admin may need to click "Grant admin consent" on this page.

You're done with Azure. You now have a **Client ID** and the right permissions.

---

## Step 2 — Configure Docker Compose

### 2.1 Clone the Repository

```bash
git clone https://github.com/eshlon/onenotemcp.git
cd onenotemcp
```

### 2.2 Add Your Client ID

Open `docker-compose.yml` and uncomment/set the `AZURE_CLIENT_ID` line:

```yaml
services:
  onenote-mcp:
    build: .
    container_name: onenote-mcp
    ports:
      - "3300:3000"
    volumes:
      - onenote-data:/data
    environment:
      - PORT=3000
      - TOKEN_FILE_PATH=/data/.access-token.txt
      - AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx   # <-- your Client ID from Step 1.3
    restart: unless-stopped

volumes:
  onenote-data:
```

Replace `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` with the actual Application (client) ID you copied.

### 2.3 Build and Start

```bash
docker compose up --build -d
```

Verify it's running:

```bash
docker compose ps
```

You should see:

```
NAME          IMAGE                    STATUS                    PORTS
onenote-mcp   onenotemcp-onenote-mcp   Up X seconds (healthy)   0.0.0.0:3300->3000/tcp
```

Check the health endpoint:

```bash
curl http://localhost:3300/health
```

```json
{"status":"ok","uptime":12,"authenticated":false,"version":"1.0.0","activeSessions":0}
```

> `"authenticated": false` is expected — you haven't logged in yet.

---

## Step 3 — Authenticate with Your Microsoft Account

### Option A: Via the Dashboard (Recommended)

1. Open **http://localhost:3300** in your browser — this is the built-in dashboard.
2. Click the **"Authenticate with Microsoft"** button.
3. A **device code** appears (e.g., `GRLF8MQ3D`) along with a link.
4. Click the link (or go to **https://microsoft.com/devicelogin**).
5. Enter the code, sign in with your Microsoft account, and approve the permissions.
6. The dashboard auto-refreshes — the badge turns **green** when authentication succeeds.

### Option B: Via the API

```bash
# Start the auth flow
curl -s -X POST http://localhost:3300/api/auth/start | python3 -m json.tool
```

Response:

```json
{
    "userCode": "GRLF8MQ3D",
    "verificationUri": "https://microsoft.com/devicelogin"
}
```

Open that URL, enter the code, sign in, and approve.

### Verify Authentication

```bash
curl -s http://localhost:3300/api/auth/status | python3 -m json.tool
```

```json
{
    "authenticated": true,
    "tokenExpiry": "2026-04-05T15:00:00.000Z",
    "clientId": "xxxxxxxx...",
    "pendingAuth": null
}
```

The token is saved inside the Docker volume (`onenote-data`), so it **survives container restarts**. You only need to re-authenticate when the token expires (~1 hour).

---

## Step 4 — Connect an AI Client

The MCP server is now running at `http://localhost:3300/mcp` using the **Streamable HTTP** transport. Configure your AI client to connect to it:

### Claude Desktop

Edit the config file:

- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux:** `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "onenote": {
      "url": "http://localhost:3300/mcp"
    }
  }
}
```

Restart Claude Desktop after saving.

### Cursor

Create or edit `.cursor/mcp.json` in your project root (or global config):

```json
{
  "mcpServers": {
    "onenote": {
      "url": "http://localhost:3300/mcp"
    }
  }
}
```

### VS Code (Copilot Chat)

Add to `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "onenote": {
      "url": "http://localhost:3300/mcp"
    }
  }
}
```

Once connected, just ask your AI naturally:

- *"List my OneNote notebooks"*
- *"Search for pages about meeting notes"*
- *"Read the page called 'Project Plan'"*
- *"Append a summary to my daily log page"*

---

## Step 5 — Example: Read & Write via Script

You can also interact with the MCP server programmatically. Save this as `test-onenote.sh`:

```bash
#!/usr/bin/env bash
# test-onenote.sh — Read and write to OneNote via the MCP server
# Usage: bash test-onenote.sh

MCP_URL="http://localhost:3300/mcp"

# --- Helper: send a JSON-RPC request and extract the session + response ---
SESSION_ID=""

mcp_request() {
  local method="$1"
  local params="$2"
  local id="$3"

  local headers=(-H "Content-Type: application/json" -H "Accept: text/event-stream")
  if [ -n "$SESSION_ID" ]; then
    headers+=(-H "Mcp-Session-Id: $SESSION_ID")
  fi

  local body="{\"jsonrpc\":\"2.0\",\"id\":${id},\"method\":\"${method}\",\"params\":${params}}"

  local response
  response=$(curl -s -D /tmp/mcp_headers -X POST "$MCP_URL" "${headers[@]}" -d "$body")

  # Capture session ID from first response
  if [ -z "$SESSION_ID" ]; then
    SESSION_ID=$(grep -i 'mcp-session-id' /tmp/mcp_headers | tr -d '\r' | awk '{print $2}')
  fi

  # Extract JSON from SSE data lines
  echo "$response" | grep '^data: ' | sed 's/^data: //'
}

echo "=== 1. Initialize MCP session ==="
INIT=$(mcp_request "initialize" '{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test-script","version":"1.0.0"}}' 1)
echo "$INIT" | python3 -m json.tool 2>/dev/null || echo "$INIT"
echo ""

# Send initialized notification (required by MCP protocol)
curl -s -X POST "$MCP_URL" \
  -H "Content-Type: application/json" \
  -H "Mcp-Session-Id: $SESSION_ID" \
  -d '{"jsonrpc":"2.0","method":"notifications/initialized"}' > /dev/null

echo "=== 2. List notebooks ==="
NOTEBOOKS=$(mcp_request "tools/call" '{"name":"listNotebooks","arguments":{}}' 2)
echo "$NOTEBOOKS" | python3 -m json.tool 2>/dev/null || echo "$NOTEBOOKS"
echo ""

echo "=== 3. Search for pages ==="
PAGES=$(mcp_request "tools/call" '{"name":"searchPages","arguments":{"query":"meeting"}}' 3)
echo "$PAGES" | python3 -m json.tool 2>/dev/null || echo "$PAGES"
echo ""

# To read a specific page (replace PAGE_ID with an actual ID from search results):
# echo "=== 4. Read a page ==="
# CONTENT=$(mcp_request "tools/call" '{"name":"getPageContent","arguments":{"pageId":"PAGE_ID","format":"text"}}' 4)
# echo "$CONTENT" | python3 -m json.tool 2>/dev/null || echo "$CONTENT"

# To create a new page:
# echo "=== 5. Create a page ==="
# CREATE=$(mcp_request "tools/call" '{"name":"createPage","arguments":{"title":"Test Page from Script","content":"Hello from bash script!"}}' 5)
# echo "$CREATE" | python3 -m json.tool 2>/dev/null || echo "$CREATE"

echo "Done. Session ID was: $SESSION_ID"
```

Run it:

```bash
chmod +x test-onenote.sh
./test-onenote.sh
```

This will list your notebooks and search for pages with "meeting" in the title. Uncomment the sections at the bottom to read specific pages or create new ones.

---

## Managing the Server

| Command | What it does |
|---------|-------------|
| `docker compose up --build -d` | Build and start (or rebuild after code changes) |
| `docker compose ps` | Check container status |
| `docker compose logs -f` | Stream live logs |
| `docker compose restart` | Restart (token is preserved) |
| `docker compose down` | Stop and remove container (token volume is kept) |
| `docker compose down -v` | Stop and **delete everything** including saved token |

### Dashboard

Open **http://localhost:3300** for a live dashboard showing:

- Server uptime and authentication status
- Per-tool call counts, success rates, and average latency
- A scrollable request log of all MCP tool invocations

### Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `http://localhost:3300` | GET | Dashboard UI |
| `http://localhost:3300/health` | GET | Health check (JSON) |
| `http://localhost:3300/api/stats` | GET | Tool metrics and call counts |
| `http://localhost:3300/api/logs` | GET | Request log (newest first) |
| `http://localhost:3300/api/auth/status` | GET | Authentication status |
| `http://localhost:3300/api/auth/start` | POST | Start device code auth flow |
| `http://localhost:3300/mcp` | POST | MCP protocol endpoint |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Container exits immediately | Run `docker compose logs` to see the error |
| Port 3300 already in use | Change the port in `docker-compose.yml`: `"3301:3000"` |
| `"authenticated": false` after restart | Token expired. Re-authenticate via dashboard or API |
| `AADSTS7000218` error during auth | Go back to Entra → your app → Authentication → set "Allow public client flows" to **Yes** |
| `AADSTS65001` or permission error | Go to Entra → your app → API permissions → make sure all 4 permissions are added |
| Rate limiting / 429 errors | You may be using the default Client ID. Create your own app registration (Step 1) |
| Dashboard loads but no tool data | No MCP client has called any tools yet — the metrics table fills as tools are used |
