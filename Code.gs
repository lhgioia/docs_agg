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
 * If a tag with the same name already exists, its aggregation doc is updated (upsert by name).
 */
function createTag(name, aggDocUrl) {
  name = name.trim();
  if (!name) throw new Error('Tag name cannot be empty.');

  const aggDocId = extractDocId_(aggDocUrl);

  let aggDocName;
  try {
    aggDocName = DocumentApp.openById(aggDocId).getName();
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
    aggregationDocName: aggDocName,
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
      aggregationDocId: linked ? linked.aggregationDocId : null,
      aggregationDocName: linked ? linked.aggregationDocName : null,
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

  const scan = scanDocumentForTagElements_();
  const nameLower = tag.name.toLowerCase();
  const entry = Object.entries(scan).find(([n]) => n.toLowerCase() === nameLower);
  return entry
    ? entry[1].map(function (e) { return { text: e.element.getText().trim(), element: e.element, level: e.level }; })
    : [];
}

/**
 * Same scan logic as scanDocumentForTags_ but stores ListItem element
 * references and relative nesting levels instead of plain text.
 * Used server-side only (element references cannot be serialized for the sidebar).
 *
 * "level" is relative to the tag marker: a direct child is level 0,
 * a grandchild is level 1, etc.
 *
 * @returns {{ [tagName: string]: { element: ListItem, level: number }[] }}
 */
function scanDocumentForTagElements_() {
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
      if (text) result[currentTag].push({ element: item, level: level - parentLevel - 1 });
    } else {
      currentTag = null;
      parentLevel = -1;
    }
  }

  return result;
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
      sourceDoc.getName(), sourceDoc.getUrl(), sourceDocId, excerpts);
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

function getDocumentInfo() {
  const doc = DocumentApp.getActiveDocument();
  return { name: doc.getName(), id: doc.getId(), url: doc.getUrl() };
}

// ── Private helpers ───────────────────────────────────────────────────────────

/**
 * Returns the file path up to 2 parent directories deep, e.g.
 * "Grandparent/Parent/Filename". Falls back to just the filename if
 * parent folders are not accessible.
 */
function getFilePath_(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const name = file.getName();
    const p1It = file.getParents();
    if (!p1It.hasNext()) return name;
    const p1 = p1It.next();
    const p2It = p1.getParents();
    if (!p2It.hasNext()) return p1.getName() + '/' + name;
    return p2It.next().getName() + '/' + p1.getName() + '/' + name;
  } catch (e) {
    return null;
  }
}

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
function upsertSection_(body, startMarker, endMarker, tag, sourceDocName, sourceDocUrl, sourceDocId, excerpts) {
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
    sP.editAsText().setFontSize(1).setForegroundColor('#FFFFFF');
    sP.setSpacingAfter(0);

    const eP = body.appendParagraph(endMarker);
    eP.editAsText().setFontSize(1).setForegroundColor('#FFFFFF');
    eP.setSpacingBefore(0);

    // End marker is the last child; insert content before it.
    insertIdx = body.getNumChildren() - 1;
  }

  // Insert source + timestamp line.
  // Link text is the file path (up to 2 parent dirs), followed by sync date and tag name.
  const filePath   = getFilePath_(sourceDocId) || sourceDocName;
  const linkText   = '[' + filePath + ']';
  const tagText    = '[' + tag.name + ']';
  const metaStr    = linkText + '   ' + tagText + '   ' + new Date().toLocaleString();
  const meta       = body.insertParagraph(insertIdx++, metaStr);
  const metaText   = meta.editAsText();
  metaText.setFontSize(9).setForegroundColor('#777777').setItalic(false);
  metaText.setLinkUrl(0, linkText.length - 1, sourceDocUrl);
  metaText.setForegroundColor(0, linkText.length - 1, '#1155CC');
  meta.setSpacingAfter(6);

  // Insert excerpts using element.copy() so all inline formatting, glyph type,
  // and list structure are preserved without manual attribute copying.
  // Nesting level is normalized to be relative to the [tag] marker.
  excerpts.forEach(function ({ element, level }) {
    const glyphType = element.getGlyphType();
    const li = body.insertListItem(insertIdx++, element.copy());
    li.setNestingLevel(level || 0);
    li.setGlyphType(glyphType);
    li.setSpacingBefore(2);
    li.setSpacingAfter(2);
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
