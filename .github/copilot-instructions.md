# Copilot Instructions

## Project Context

Before starting any work on this codebase, **read the `documentation/` folder** in the project root. It contains:

- [documentation/README.md](../documentation/README.md) — Project overview, tech stack, and quick start
- [documentation/architecture.md](../documentation/architecture.md) — System design, component map, data flow diagrams
- [documentation/authentication.md](../documentation/authentication.md) — OAuth Device Code flow, token lifecycle, Azure App Registration
- [documentation/tools-reference.md](../documentation/tools-reference.md) — Complete API reference for all 13 MCP tools
- [documentation/setup-guide.md](../documentation/setup-guide.md) — Installation, configuration, and MCP client integration
- [documentation/content-processing.md](../documentation/content-processing.md) — HTML extraction, markdown conversion, formatting internals

## Project Summary

This is a **Model Context Protocol (MCP) server** that bridges AI clients (Claude, Cursor, etc.) with **Microsoft OneNote** via the **Microsoft Graph API**. The entire server is a single file: `onenote-mcp.mjs`.

## Key Conventions

- **Single-file architecture** — all server code lives in `onenote-mcp.mjs`. Do not split into multiple files unless explicitly asked.
- **ES Modules** — the project uses `"type": "module"` with `.mjs` extension and `import` syntax.
- **Zod for validation** — all MCP tool input schemas are defined inline with Zod in `server.tool()` calls.
- **Stdio transport** — the server communicates via stdin/stdout JSON-RPC. Diagnostic output goes to `stderr` (`console.error`).
- **Token storage** — authentication tokens are saved to `.access-token.txt` (JSON format). This file must never be committed.
- **Graph API** — page content operations use raw `fetch()` with Bearer tokens (not the Graph SDK) for reliability.
- **Content conversion** — `textToHtml()` converts markdown input to HTML; `extractReadableText()` converts HTML responses to plain text.
