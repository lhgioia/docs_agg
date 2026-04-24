# Docs Aggregator

A Google Docs Editor Add-on that lets you tag text excerpts using inline list syntax and automatically sync them to a separate aggregation document.

## Features

- **Inline tag syntax** — write `[Tag Name]` as a list item; indent child items beneath it and they become the synced excerpts
- **Sync** — a single button scans the document and syncs all linked tags to their aggregation docs; header shows live status ("Scanning…", "Syncing…", "Synced at X:XX")
- **Auto-link markers** — on every sync, `[Tag Name]` markers in the source doc are automatically hyperlinked to their aggregation doc
- **Inline formatting preserved** — bold, italic, underline, font size, etc. are carried over to the aggregation doc via `element.copy()`
- **Nested excerpts** — multi-level nesting under a tag marker is preserved; nesting levels are normalized relative to the marker
- **Settings pane** — gear icon opens a settings view with two tabs:
  - *Manage Tags* — view all defined tags, change their aggregation doc link, or delete them
  - *Timestamp Settings* — configure font size, number of parent folders shown in the source link, link color, text color, and date/time format

## How it works

### Tagging

Tags are written directly in the document using list syntax:

```
• [Research Notes]
  • This child item will be synced
  • So will this one
    • Nested items are synced too
• [Bug Fixes]
  • Regression in the login flow
```

Any list item whose full trimmed text matches `[Tag Name]` is treated as a tag marker. All list items at a strictly deeper nesting level beneath it are collected as excerpts. The group ends at the first list item at the same or shallower level, or at any non-list element (paragraph, table, etc.).

The document is the source of truth — no named ranges or hidden metadata are stored in the source document. The sidebar scans the document on every sync.

### Aggregation document format

Each `(tag, source document)` pair gets a delimited section in the aggregation doc:

```
[DOCSAGG:<tagId>:<sourceDocId>]      ← hidden marker (1pt, white)
[Grandparent/Parent/Filename]   [Tag Name]   1/1/2025 2:30 PM
  • excerpt 1
  • excerpt 2
    • nested excerpt
[/DOCSAGG:<tagId>:<sourceDocId>]     ← hidden marker (1pt, white)
```

Re-syncing replaces only the content between the markers — the aggregation doc always reflects the current state of the source document. The timestamp line format is configurable in Settings.

### Data storage

Tag definitions and timestamp settings are stored in `UserProperties` and persist across all documents for the signed-in user. No data is stored in the source document beyond the `[Tag Name]` list items themselves.

## Project structure

```
.
├── appsscript.json   — Apps Script manifest & OAuth scopes
├── Code.gs           — Add-on lifecycle (onOpen, onInstall, showSidebar)
├── Tags.gs           — Tag CRUD, document tag summary, auto-link markers
├── Scan.gs           — Document scanning functions
├── Sync.gs           — Sync to aggregation document
├── Settings.gs       — Timestamp display settings
├── Utils.gs          — Shared helpers (file path, timestamp formatting, doc ID extraction)
└── Sidebar.html      — Client-side sidebar UI
```

## Setup

### Prerequisites

```bash
npm install -g @google/clasp
clasp login
```

### 1. Create an Apps Script project

Go to [script.google.com](https://script.google.com) → **New project**.
Copy the script ID from the URL: `https://script.google.com/d/<SCRIPT_ID>/edit`

### 2. Configure clasp

Edit `.clasp.json` and replace the placeholder with your script ID:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "."
}
```

### 3. Push the code

```bash
clasp push
```

### 4. Test in Google Docs

1. Open the Google Doc the script is bound to.
2. Reload the page — an **Extensions → Docs Aggregator** menu appears.
3. Open the sidebar via **Extensions → Docs Aggregator → Open Docs Aggregator**.

### 5. Publish (optional)

To install the add-on from the Workspace Marketplace, create a deployment via **Deploy → New deployment → Add-on** in the Apps Script editor.
