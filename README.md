# Docs Aggregator

A Google Docs Editor Add-on that lets you tag text excerpts and automatically sync them to a separate aggregation document.

## Features

- **Create tags** — each tag links to a user-specified aggregation Google Doc
- **Apply tags to selections** — highlights tagged text with a chosen color and tracks it via a named range
- **Tag dropdown** — sidebar shows all your tags; selecting one lists every excerpt from the current document
- **Sync** — push excerpts to the aggregation doc (idempotent; re-syncing replaces the previous version)
- **Sync All** — sync every tag in one click

## Project structure

```
.
├── appsscript.json   — Apps Script manifest & OAuth scopes
├── Code.gs           — Server-side logic (tag CRUD, named ranges, sync)
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

1. Open any Google Doc.
2. In Apps Script editor: **Run → onInstall** (grants permissions).
3. Reload the Doc — an **Extensions → Docs Aggregator** menu appears.
4. Open the sidebar via **Extensions → Docs Aggregator → Open Docs Aggregator**.

### 5. Publish (optional)

To install the add-on from the Workspace Marketplace, create a deployment via **Deploy → New deployment → Add-on** in the Apps Script editor.

## How it works

### Tagging mechanism

When you apply a tag to a selection:
1. A **named range** is created in the document with the convention `docsagg_<tagId>_<timestamp>`.
2. The selected text is **highlighted** with the tag's color.

Named ranges survive edits to surrounding text and are used to retrieve the current content of each excerpt at sync time.

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

- **Tags** (name, aggregation doc ID, color) are stored in `UserProperties` — they persist across all documents for the user.
- **Tagged ranges** live as named ranges inside each source document.
