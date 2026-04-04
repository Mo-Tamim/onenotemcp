# OneNote MCP Server — Documentation

Welcome to the OneNote MCP Server documentation. This guide covers everything you need to understand, set up, and use the server.

## Table of Contents

| Document | Description |
|----------|-------------|
| [Architecture Overview](architecture.md) | System design, component map, and data flow diagrams |
| [Setup Guide](setup-guide.md) | Prerequisites, installation, configuration, and first run |
| [Authentication Guide](authentication.md) | Microsoft identity flow, token lifecycle, and troubleshooting |
| [Tools API Reference](tools-reference.md) | Complete reference for every MCP tool exposed by the server |
| [Content Processing](content-processing.md) | HTML extraction, markdown-to-HTML conversion, and formatting internals |

## What Is This Project?

The OneNote MCP Server is a **Model Context Protocol (MCP)** server that bridges AI language models (such as Claude, Cursor, or any MCP-compatible client) with **Microsoft OneNote** via the **Microsoft Graph API**. It allows an LLM to read, search, create, and edit OneNote notebooks, sections, and pages on behalf of an authenticated user.

### Key Highlights

- **13 MCP tools** spanning authentication, reading, editing, and page creation.
- **Device Code OAuth 2.0** flow — no browser redirect server needed.
- **Markdown & HTML** content support with automatic conversion.
- **Zod-validated** input schemas for every tool.
- **Single-file server** (`onenote-mcp.mjs`) — easy to deploy and extend.

## Quick Start

```bash
# 1. Clone & install
git clone https://github.com/eshlon/onenotemcp.git
cd onenotemcp
npm install

# 2. (Optional) Set your own Azure App Client ID
export AZURE_CLIENT_ID="your-client-id"

# 3. Start the server
node onenote-mcp.mjs
```

Then connect your MCP client (Claude Desktop, Cursor, etc.) and invoke the `authenticate` tool to sign in.

## Technology Stack

| Technology | Purpose |
|------------|---------|
| Node.js ≥ 18 | Runtime |
| `@modelcontextprotocol/sdk` | MCP server framework |
| `@azure/identity` | Azure Device Code OAuth 2.0 |
| `@microsoft/microsoft-graph-client` | OneNote / Graph API client |
| `jsdom` | Server-side HTML parsing |
| `zod` | Input schema validation |
| `node-fetch` | HTTP requests to Graph API |
