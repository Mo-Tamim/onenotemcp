# Content Processing

This document covers the internal utilities that convert content between HTML, plain text, and markdown. These functions are the backbone of every read and write operation.

## Processing Pipeline Overview

```mermaid
graph TD
    subgraph "Inbound (OneNote → User)"
        A["OneNote Page HTML"] --> B["extractReadableText()"]
        A --> C["extractTextSummary()"]
        A --> D["Raw HTML<br/>(passthrough)"]
        B --> E["Readable plain text"]
        C --> F["≤300 char summary"]
    end

    subgraph "Outbound (User → OneNote)"
        G["User input<br/>(text/markdown)"] --> H{"Already HTML?"}
        H -- Yes --> I["Pass through"]
        H -- No --> J["textToHtml()"]
        J --> K["HTML for Graph API"]
    end
```

---

## `extractReadableText(html)`

Converts raw OneNote HTML into structured, readable plain text. Used when the `format` parameter is `"text"`.

### Algorithm

```mermaid
graph TD
    A["Input HTML"] --> B["Parse with JSDOM"]
    B --> C["Remove script & style tags"]
    C --> D["Extract headings<br/>h1–h6 → text + underline"]
    D --> E["Extract paragraphs<br/>p → text + blank line"]
    E --> F["Extract lists<br/>ul/ol → bullets/numbers"]
    F --> G["Extract tables<br/>tr/td → pipe-separated rows"]
    G --> H{"Any text extracted?"}
    H -- Yes --> I["Return trimmed text"]
    H -- No --> J["Fallback: body.textContent"]
```

### Formatting Rules

| HTML Element | Output Format |
|-------------|---------------|
| `<h1>` – `<h6>` | Text followed by a line of dashes (`---`) |
| `<p>` | Text followed by a blank line |
| Unordered list (`<ul><li>`) | `- item` |
| Ordered list (`<ol><li>`) | `1. item`, `2. item`, ... |
| Table (`<table>`) | `📊 Table content:` header, then pipe-separated rows |
| Everything else | Falls back to `body.textContent` with whitespace collapsed |

### Example

**Input HTML:**
```html
<h1>Meeting Notes</h1>
<p>Discussed Q2 goals with the team.</p>
<ul>
  <li>Launch feature X</li>
  <li>Hire two engineers</li>
</ul>
```

**Output text:**
```
Meeting Notes
-------------

Discussed Q2 goals with the team.

- Launch feature X
- Hire two engineers
```

---

## `extractTextSummary(html, maxLength = 300)`

Returns a truncated plain-text summary of the HTML body. Used when the `format` parameter is `"summary"`.

### Algorithm

1. Parse HTML with JSDOM.
2. Get `body.textContent`, trim, and collapse whitespace.
3. Truncate to `maxLength` characters.
4. Append `...` if truncated.

### Example

A 500-character body text becomes:

```
First 300 characters of the page content here...
```

---

## `textToHtml(text)`

Converts plain text (with optional markdown syntax) into HTML suitable for the OneNote PATCH API. Used by all write/edit tools.

### Detection Logic

```mermaid
graph TD
    A["Input text"] --> B{"Contains <html> or <!DOCTYPE html>?"}
    B -- Yes --> C["Return as-is<br/>(already HTML)"]
    B -- No --> D["Apply markdown conversions"]
    D --> E["Wrap remaining lines in &lt;p&gt; tags"]
    E --> F["Group &lt;li&gt; into &lt;ul&gt;"]
    F --> G["Group &lt;blockquote&gt; runs"]
    G --> H["Return HTML string"]
```

### Supported Markdown Conversions

| Markdown Syntax | HTML Output |
|----------------|-------------|
| `` ```code``` `` | `<pre><code>code</code></pre>` |
| `` `inline` `` | `<code>inline</code>` |
| `### Heading` | `<h3>Heading</h3>` |
| `## Heading` | `<h2>Heading</h2>` |
| `# Heading` | `<h1>Heading</h1>` |
| `**bold**` or `__bold__` | `<strong>bold</strong>` |
| `*italic*` or `_italic_` | `<em>italic</em>` |
| `[text](url)` | `<a href="url">text</a>` |
| `---` | `<hr>` |
| `> quote` | `<blockquote>quote</blockquote>` |
| `- item` / `* item` / `+ item` | `<li>item</li>` (grouped into `<ul>`) |
| `1. item` | `<li>item</li>` (grouped into `<ul>`) |
| Plain line | `<p>line</p>` |

### Processing Order

The conversion is applied in a specific order to avoid conflicts:

```mermaid
graph TD
    A["1. HTML-escape &amp; &lt; &gt;"] --> B["2. Fenced code blocks"]
    B --> C["3. Inline code"]
    C --> D["4. Headings (h3 → h2 → h1)"]
    D --> E["5. Bold"]
    E --> F["6. Italic"]
    F --> G["7. Links"]
    G --> H["8. Horizontal rules"]
    H --> I["9. Blockquotes"]
    I --> J["10. List items"]
    J --> K["11. Wrap remaining lines in &lt;p&gt;"]
    K --> L["12. Group &lt;li&gt; into &lt;ul&gt;"]
    L --> M["13. Group adjacent &lt;blockquote&gt;"]
```

### Important: HTML Passthrough

If the input already looks like a full HTML document (contains `<html>` or `<!DOCTYPE html>`), the function returns it unchanged. This allows tools like `createPage` to accept raw HTML content directly.

---

## `fetchPageContentAdvanced(pageId, method)`

A utility that fetches the raw HTML content of a OneNote page.

| Method | Implementation | Notes |
|--------|---------------|-------|
| `httpDirect` (default) | `fetch()` with Bearer token to `https://graph.microsoft.com/v1.0/me/onenote/pages/{id}/content` | Preferred — handles binary/HTML reliably |
| `direct` | `graphClient.api(...).get()` | Fallback using the Graph SDK |

---

## `formatPageInfo(page, index)`

Formats a OneNote page object into a display string.

**Output format:**
```
1. **Page Title**
   ID: 0-abc123...
   Created: 4/4/2026
   Modified: 4/4/2026
```

Used by `listNotebooks`, `searchPages`, and other listing tools to present results consistently.

---

## Data Flow Summary

```mermaid
graph LR
    subgraph "Read Path"
        R1["Graph API"] -->|"HTML"| R2["extractReadableText()"]
        R1 -->|"HTML"| R3["extractTextSummary()"]
        R1 -->|"HTML"| R4["Pass-through"]
    end

    subgraph "Write Path"
        W1["User markdown/text"] --> W2["textToHtml()"]
        W2 --> W3["Tool wraps in div/table/note HTML"]
        W3 --> W4["Graph API PATCH"]
    end

    subgraph "Create Path"
        C1["User markdown/text"] --> C2["textToHtml()"]
        C2 --> C3["Full XHTML document"]
        C3 --> C4["Graph API POST"]
    end
```
