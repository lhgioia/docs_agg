/**
 * Docs Aggregator — Google Docs Editor Add-on
 *
 * Tags are written inline in the document as [tagname] list items.
 * Child list items (deeper nesting) under a [tagname] item are the excerpts.
 *
 * Example document structure:
 *   • [Research Notes]
 *     • This sentence will be synced.
 *     • So will this one.
 *   • [Bug Fixes]
 *     • Fix for login regression
 *
 * Tag metadata (aggregation doc URL, color) is stored in UserProperties.
 * The document itself is the source of truth for which tags exist and what
 * content they contain.
 */

const TAGS_PROP_KEY = 'DOCSAGG_TAGS_V1';

const TAG_COLORS = [
  { name: 'Yellow', value: '#FFF176' },
  { name: 'Blue',   value: '#BBDEFB' },
  { name: 'Green',  value: '#C8E6C9' },
  { name: 'Pink',   value: '#F48FB1' },
  { name: 'Orange', value: '#FFCC80' },
  { name: 'Purple', value: '#CE93D8' },
  { name: 'Teal',   value: '#80DEEA' },
  { name: 'Red',    value: '#EF9A9A' },
];

// ── Add-on lifecycle ──────────────────────────────────────────────────────────

function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Open Docs Aggregator', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Docs Aggregator')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

// ── Tag metadata CRUD ─────────────────────────────────────────────────────────

/** Returns the tags map: { [tagId]: { name, aggregationDocId, color, createdAt } } */
function getTags() {
  const raw = PropertiesService.getUserProperties().getProperty(TAGS_PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveTags_(tags) {
  PropertiesService.getUserProperties().setProperty(TAGS_PROP_KEY, JSON.stringify(tags));
}

/**
 * Creates or updates a tag definition.
 * If a tag with the same name already exists, its aggregation doc and color
 * are updated (upsert by name).
 */
function createTag(name, aggDocUrl, color) {
  name = name.trim();
  if (!name) throw new Error('Tag name cannot be empty.');

  const aggDocId = extractDocId_(aggDocUrl);

  try {
    DocumentApp.openById(aggDocId);
  } catch (e) {
    throw new Error(
      'Cannot access the aggregation document. Check the URL and your edit access.'
    );
  }

  const tags = getTags();

  // Upsert by name (case-insensitive).
  let tagId = null;
  Object.entries(tags).forEach(([id, tag]) => {
    if (tag.name.toLowerCase() === name.toLowerCase()) tagId = id;
  });
  tagId = tagId || Utilities.getUuid();

  tags[tagId] = {
    name,
    aggregationDocId: aggDocId,
    color: color || TAG_COLORS[0].value,
    createdAt: (tags[tagId] && tags[tagId].createdAt) || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  saveTags_(tags);
  return { tagId, ...tags[tagId] };
}

/** Removes a tag definition. Does not modify the document. */
function deleteTag(tagId) {
  const tags = getTags();
  if (!tags[tagId]) throw new Error('Tag not found.');
  delete tags[tagId];
  saveTags_(tags);
}

// ── Document scanning ─────────────────────────────────────────────────────────

/**
 * Scans the active document for [tagname] list items and collects their
 * deeper-nested children as excerpts.
 *
 * Rules:
 *   - A list item whose full trimmed text matches /^\[.+\]$/ is a tag marker.
 *   - Subsequent list items at a strictly deeper nesting level are excerpts.
 *   - A list item at the same or shallower level as the marker ends the group.
 *   - A non-list element (paragraph, table, etc.) also ends the current group.
 *
 * @returns {{ [tagName: string]: string[] }}
 */
function scanDocumentForTags_() {
  const body = DocumentApp.getActiveDocument().getBody();
  const result = {};
  let currentTag = null;
  let parentLevel = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (child.getType() !== DocumentApp.ElementType.LIST_ITEM) {
      currentTag = null;
      parentLevel = -1;
      continue;
    }

    const item = child.asListItem();
    const level = item.getNestingLevel();
    const text = item.getText().trim();
    const tagMatch = text.match(/^\[([^\]]+)\]$/);

    if (tagMatch) {
      currentTag = tagMatch[1];
      parentLevel = level;
      if (!result[currentTag]) result[currentTag] = [];
    } else if (currentTag !== null && level > parentLevel) {
      if (text) result[currentTag].push(text);
    } else {
      currentTag = null;
      parentLevel = -1;
    }
  }

  return result;
}

/**
 * Returns all tags found in the document merged with their definitions.
 * This is the primary data source for the sidebar.
 *
 * @returns {{
 *   inDocument: Array<{
 *     name: string,
 *     excerpts: string[],
 *     defined: boolean,
 *     tagId: string|null,
 *     color: string|null,
 *     aggregationDocId: string|null
 *   }>,
 *   notInDocument: Array<{ tagId: string, name: string, color: string, aggregationDocId: string }>
 * }}
 */
function getDocumentTagSummary() {
  const tags = getTags();
  const scan = scanDocumentForTags_();

  // name (lowercase) → tag definition
  const nameToTag = {};
  Object.entries(tags).forEach(([id, tag]) => {
    nameToTag[tag.name.toLowerCase()] = { id, ...tag };
  });

  const inDocument = Object.entries(scan).map(([name, excerpts]) => {
    const linked = nameToTag[name.toLowerCase()] || null;
    return {
      name,
      excerpts,
      defined: !!linked,
      tagId: linked ? linked.id : null,
      color: linked ? linked.color : null,
      aggregationDocId: linked ? linked.aggregationDocId : null,
    };
  });

  const docTagNames = new Set(Object.keys(scan).map(n => n.toLowerCase()));
  const notInDocument = Object.entries(tags)
    .filter(([, tag]) => !docTagNames.has(tag.name.toLowerCase()))
    .map(([id, tag]) => ({ tagId: id, ...tag }));

  return { inDocument, notInDocument };
}

// ── Content retrieval ─────────────────────────────────────────────────────────

function getTaggedContent(tagId) {
  const tags = getTags();
  const tag = tags[tagId];
  if (!tag) return [];

  const scan = scanDocumentForTags_();
  const nameLower = tag.name.toLowerCase();
  const entry = Object.entries(scan).find(([n]) => n.toLowerCase() === nameLower);
  return entry ? entry[1].map(text => ({ text })) : [];
}

// ── Sync ──────────────────────────────────────────────────────────────────────

/**
 * Syncs excerpts for a defined tag to its aggregation document.
 * The section in the aggregation doc is replaced on every call (idempotent).
 */
function syncTagToAggDoc(tagId) {
  const tags = getTags();
  const tag = tags[tagId];
  if (!tag) throw new Error('Tag not found.');

  const sourceDoc = DocumentApp.getActiveDocument();
  const excerpts = getTaggedContent(tagId);

  let aggDoc;
  try {
    aggDoc = DocumentApp.openById(tag.aggregationDocId);
  } catch (e) {
    throw new Error('Could not open aggregation document: ' + e.message);
  }

  const body = aggDoc.getBody();
  const sourceDocId = sourceDoc.getId();
  const startMarker = `[DOCSAGG:${tagId}:${sourceDocId}]`;
  const endMarker = `[/DOCSAGG:${tagId}:${sourceDocId}]`;

  if (excerpts.length > 0) {
    upsertSection_(body, startMarker, endMarker, tag,
      sourceDoc.getName(), sourceDoc.getUrl(), excerpts);
  } else {
    removeSection_(body, startMarker, endMarker);
  }

  return { synced: excerpts.length };
}

/** Syncs all defined tags that have excerpts in the active document. */
function syncAllTags() {
  const tags = getTags();
  const results = {};
  const errors = {};
  Object.keys(tags).forEach(tagId => {
    try {
      results[tagId] = syncTagToAggDoc(tagId);
    } catch (e) {
      errors[tagId] = e.message;
    }
  });
  return { results, errors };
}

// ── Misc ──────────────────────────────────────────────────────────────────────

function getTagColors() { return TAG_COLORS; }

function getDocumentInfo() {
  const doc = DocumentApp.getActiveDocument();
  return { name: doc.getName(), id: doc.getId(), url: doc.getUrl() };
}

// ── Private helpers ───────────────────────────────────────────────────────────

function extractDocId_(urlOrId) {
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : urlOrId.trim();
}

/**
 * Creates or updates a tagged section in the aggregation document.
 *
 * On update: clears only the content *between* the marker paragraphs (the
 * markers themselves are never removed), then re-inserts fresh content.
 * This avoids the "can't remove last paragraph" error entirely because the
 * markers always remain as surviving paragraphs.
 *
 * On first write: appends markers and content at the end of the body.
 */
function upsertSection_(body, startMarker, endMarker, tag, sourceDocName, sourceDocUrl, excerpts) {
  let startIdx = -1;
  let endIdx = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const t = child.asParagraph().getText();
      if (t === startMarker) startIdx = i;
      else if (t === endMarker) endIdx = i;
    }
  }

  let insertIdx; // index at which to start inserting content paragraphs

  if (startIdx !== -1 && endIdx !== -1 && startIdx < endIdx) {
    // Section exists. Remove everything between the markers (high → low so
    // indices stay valid as elements are removed).
    for (let i = endIdx - 1; i > startIdx; i--) {
      body.removeChild(body.getChild(i));
    }
    // Markers are now adjacent: startIdx = start, startIdx+1 = end.
    // Insert content before the end marker.
    insertIdx = startIdx + 1;
  } else {
    // No existing section — append markers at the end.
    if (body.getText().trim().length > 0) {
      body.appendHorizontalRule();
    }

    const sP = body.appendParagraph(startMarker);
    sP.editAsText().setFontSize(6).setForegroundColor('#CCCCCC');
    sP.setSpacingAfter(0);

    const eP = body.appendParagraph(endMarker);
    eP.editAsText().setFontSize(6).setForegroundColor('#CCCCCC');
    eP.setSpacingBefore(0);

    // End marker is the last child; insert content before it.
    insertIdx = body.getNumChildren() - 1;
  }

  // Insert heading.
  const heading = body.insertParagraph(insertIdx++, '');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  heading.setSpacingBefore(8);
  const hText = heading.editAsText();
  hText.insertText(0, '  ' + tag.name + ' — ' + sourceDocName);
  hText.setBackgroundColor(0, 0, tag.color);
  if (hText.getText().length > 1) {
    hText.setForegroundColor(1, hText.getText().length - 1, '#222222');
  }

  // Insert source + timestamp line.
  const meta = body.insertParagraph(
    insertIdx++,
    'Source: ' + sourceDocUrl + '   |   Synced: ' + new Date().toLocaleString()
  );
  meta.editAsText().setFontSize(8).setForegroundColor('#777777').setItalic(true);
  meta.setSpacingAfter(6);

  // Insert excerpts.
  excerpts.forEach(function ({ text }) {
    const p = body.insertParagraph(insertIdx++, text);
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setIndentStart(18);
    p.setSpacingBefore(4);
    p.setSpacingAfter(4);
    p.editAsText().setBackgroundColor(tag.color).setFontSize(11);
  });
}

/**
 * Removes an entire section (markers + content) when there are no excerpts
 * to sync. Uses high-to-low index removal so indices stay valid after each
 * removal. Appends a blank paragraph first if removing the section would
 * leave the body with no paragraphs at all.
 */
function removeSection_(body, startMarker, endMarker) {
  let startIdx = -1;
  let endIdx = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const t = child.asParagraph().getText();
      if (t === startMarker) startIdx = i;
      else if (t === endMarker) endIdx = i;
    }
  }

  if (startIdx === -1 || endIdx === -1 || startIdx >= endIdx) return;

  // Count paragraphs outside the section.
  let paragraphsOutside = 0;
  for (let i = 0; i < body.getNumChildren(); i++) {
    if (i >= startIdx && i <= endIdx) continue;
    if (body.getChild(i).getType() === DocumentApp.ElementType.PARAGRAPH) {
      paragraphsOutside++;
    }
  }

  // Guard: if nothing outside, ensure a paragraph remains after removal.
  if (paragraphsOutside === 0) {
    body.appendParagraph(''); // appended beyond endIdx, safe from the loop below
  }

  // Remove from high to low so earlier indices stay valid.
  for (let i = endIdx; i >= startIdx; i--) {
    body.removeChild(body.getChild(i));
  }
}
