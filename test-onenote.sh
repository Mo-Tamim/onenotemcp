#!/usr/bin/env bash
# test-onenote.sh — Example: interact with OneNote via the MCP server
#
# Prerequisites:
#   1. The Docker container is running:  docker compose up -d
#   2. You have authenticated via the dashboard (http://localhost:3300)
#
# Usage:
#   bash test-onenote.sh                     # List notebooks + search pages
#   bash test-onenote.sh read  <PAGE_ID>     # Read a specific page
#   bash test-onenote.sh create "Title" "Content"  # Create a new page

set -euo pipefail

MCP_URL="${MCP_URL:-http://localhost:3300/mcp}"
SESSION_ID=""
REQUEST_ID=0

# ─── Helper: send a JSON-RPC request over MCP Streamable HTTP ───
mcp_request() {
  local method="$1"
  local params="$2"
  REQUEST_ID=$((REQUEST_ID + 1))

  local headers=(-H "Content-Type: application/json" -H "Accept: application/json, text/event-stream")
  if [ -n "$SESSION_ID" ]; then
    headers+=(-H "Mcp-Session-Id: $SESSION_ID")
  fi

  local body="{\"jsonrpc\":\"2.0\",\"id\":${REQUEST_ID},\"method\":\"${method}\",\"params\":${params}}"
  local response
  response=$(curl -s -D /tmp/mcp_headers -X POST "$MCP_URL" "${headers[@]}" -d "$body")

  # Capture session ID from the first response
  if [ -z "$SESSION_ID" ]; then
    SESSION_ID=$(grep -i 'mcp-session-id' /tmp/mcp_headers 2>/dev/null | tr -d '\r' | awk '{print $2}')
  fi

  # Extract JSON from SSE "data:" lines
  echo "$response" | grep '^data: ' | sed 's/^data: //'
}

# ─── Initialize MCP session ───
init_session() {
  echo "Connecting to MCP server at $MCP_URL ..."
  local init
  init=$(mcp_request "initialize" '{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test-onenote-script","version":"1.0.0"}}')

  # Extract session ID from headers (subshell in mcp_request can't set outer vars)
  SESSION_ID=$(grep -i 'mcp-session-id' /tmp/mcp_headers 2>/dev/null | tr -d '\r' | awk '{print $2}')

  if [ -z "$SESSION_ID" ]; then
    echo "ERROR: Failed to connect. Is the server running?" >&2
    echo "  Try: docker compose ps" >&2
    exit 1
  fi

  # Send the required initialized notification
  curl -s -X POST "$MCP_URL" \
    -H "Content-Type: application/json" \
    -H "Accept: application/json, text/event-stream" \
    -H "Mcp-Session-Id: $SESSION_ID" \
    -d '{"jsonrpc":"2.0","method":"notifications/initialized"}' > /dev/null

  echo "Connected. Session: ${SESSION_ID:0:8}..."
  echo ""
}

# ─── Call an MCP tool and pretty-print the result ───
call_tool() {
  local tool_name="$1"
  local arguments="$2"
  local result
  result=$(mcp_request "tools/call" "{\"name\":\"${tool_name}\",\"arguments\":${arguments}}")
  echo "$result" | python3 -m json.tool 2>/dev/null || echo "$result"
}

# ─── Commands ───

cmd_list() {
  echo "=== Your OneNote Notebooks ==="
  call_tool "listNotebooks" "{}"
  echo ""
  echo "=== Recent Pages ==="
  call_tool "searchPages" "{}"
}

cmd_search() {
  local query="$1"
  echo "=== Searching for: $query ==="
  call_tool "searchPages" "{\"query\":\"${query}\"}"
}

cmd_read() {
  local page_id="$1"
  local format="${2:-text}"
  echo "=== Reading page $page_id ==="
  call_tool "getPageContent" "{\"pageId\":\"${page_id}\",\"format\":\"${format}\"}"
}

cmd_create() {
  local title="$1"
  local content="$2"
  echo "=== Creating page: $title ==="
  # Escape quotes in title and content for JSON
  title=$(echo "$title" | sed 's/"/\\"/g')
  content=$(echo "$content" | sed 's/"/\\"/g')
  call_tool "createPage" "{\"title\":\"${title}\",\"content\":\"${content}\"}"
}

cmd_append() {
  local page_id="$1"
  local content="$2"
  echo "=== Appending to page $page_id ==="
  content=$(echo "$content" | sed 's/"/\\"/g')
  call_tool "appendToPage" "{\"pageId\":\"${page_id}\",\"content\":\"${content}\"}"
}

# ─── Main ───

init_session

case "${1:-list}" in
  list)
    cmd_list
    ;;
  search)
    cmd_search "${2:?Usage: $0 search <query>}"
    ;;
  read)
    cmd_read "${2:?Usage: $0 read <page_id> [text|html|summary]}" "${3:-text}"
    ;;
  create)
    cmd_create "${2:?Usage: $0 create <title> <content>}" "${3:?Usage: $0 create <title> <content>}"
    ;;
  append)
    cmd_append "${2:?Usage: $0 append <page_id> <content>}" "${3:?Usage: $0 append <page_id> <content>}"
    ;;
  *)
    echo "Usage: $0 {list|search|read|create|append}"
    echo ""
    echo "  list                          List notebooks and pages"
    echo "  search <query>                Search pages by title"
    echo "  read <page_id> [format]       Read a page (text/html/summary)"
    echo "  create <title> <content>      Create a new page"
    echo "  append <page_id> <content>    Append content to a page"
    exit 1
    ;;
esac

echo ""
echo "Done."
