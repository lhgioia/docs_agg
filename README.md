# Docs Aggregator

A Google Docs Editor Add-on that lets you tag text excerpts using inline list syntax and automatically sync them to a separate aggregation document.

## Features

- **Inline tag syntax** — write `[Tag Name]` as a list item; indent child items beneath it and they become the synced excerpts
- **Document scan** — the sidebar scans the document on open and shows every `[tag]` it finds, with a live excerpt preview
- **Link tags** — associate a tag name with an aggregation Google Doc and a highlight color
- **Sync** — push excerpts to the aggregation doc (idempotent; re-syncing replaces the previous version)
- **Sync All** — sync every linked tag in one click

## Project structure

```
.
├── appsscript.json   — Apps Script manifest & OAuth scopes
├── Code.gs           — Server-side logic (tag CRUD, document scanning, sync)
├── Sidebar.html      — Client-side sidebar UI
└── .clasp.json       — clasp deployment config (fill in your scriptId)
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

Edit `.clasp.json` and replace `YOUR_SCRIPT_ID_HERE` with your script ID:

```json
{
  "scriptId": "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms",
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

## How it works

### Tagging mechanism

Tags are written directly in the document using list syntax:

```
• [Research Notes]
  • This child item will be synced
  • So will this one
• [Bug Fixes]
  • Regression in the login flow
```

Any list item whose full text matches `[Tag Name]` is treated as a tag marker. All list items indented beneath it (at a strictly deeper nesting level) are collected as excerpts. The group ends when a list item at the same or shallower level is encountered, or when a non-list element (paragraph, table, etc.) appears.

The document is the source of truth — no named ranges or text highlighting are used. The sidebar scans the document each time it loads or is refreshed.

### Aggregation document format

Each `(tag, source document)` pair gets a delimited section:

```
[DOCSAGG:<tagId>:<sourceDocId>]    ← hidden marker (6pt gray)
## <TagName> — <Source Doc Name>
Source: <url>   |   Synced: <timestamp>
    excerpt 1 (highlighted, indented)
    excerpt 2
[/DOCSAGG:<tagId>:<sourceDocId>]   ← hidden marker
```

Re-syncing removes the old section and appends a fresh one, so the aggregation doc always reflects the current state of the source document.

### Data storage

- **Tag definitions** (name, aggregation doc ID, color) are stored in `UserProperties` — they persist across all documents for the user.
- **Tag content** lives in the document itself as ordinary list items — no hidden metadata is stored in the source document.
