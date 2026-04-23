/**
 * Docs Aggregator — Google Docs Editor Add-on
 *
 * Tags selected text with a named range, highlights it, and syncs excerpts
 * to a user-specified aggregation document.
 *
 * Data storage:
 *   - Tags are stored in UserProperties so they persist across all documents.
 *   - Named ranges in the active document mark tagged text with the convention:
 *       docsagg_<tagId>_<timestamp>
 *   - Background color is applied to tagged text for visual indication.
 *
 * Aggregation document format:
 *   - Each (tag, sourceDoc) pair gets its own section delimited by hidden
 *     marker paragraphs: [DOCSAGG:<tagId>:<sourceDocId>] … [/DOCSAGG:…]
 *   - Syncing replaces the existing section (idempotent).
 */

// ── Constants ─────────────────────────────────────────────────────────────────

const RANGE_PREFIX = 'docsagg_';
const TAGS_PROP_KEY = 'DOCSAGG_TAGS_V1';

const TAG_COLORS = [
  { name: 'Yellow',  value: '#FFF176' },
  { name: 'Blue',    value: '#BBDEFB' },
  { name: 'Green',   value: '#C8E6C9' },
  { name: 'Pink',    value: '#F48FB1' },
  { name: 'Orange',  value: '#FFCC80' },
  { name: 'Purple',  value: '#CE93D8' },
  { name: 'Teal',    value: '#80DEEA' },
  { name: 'Red',     value: '#EF9A9A' },
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

// ── Tag CRUD ──────────────────────────────────────────────────────────────────

/** Returns the tags map: { [tagId]: { name, aggregationDocId, color, createdAt } } */
function getTags() {
  const raw = PropertiesService.getUserProperties().getProperty(TAGS_PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveTags_(tags) {
  PropertiesService.getUserProperties().setProperty(TAGS_PROP_KEY, JSON.stringify(tags));
}

/**
 * Creates a new tag.
 * @param {string} name          Display name for the tag.
 * @param {string} aggDocUrl     Full URL (or ID) of the aggregation document.
 * @param {string} color         Hex background color for highlights.
 * @returns {{ tagId, name, aggregationDocId, color, createdAt }}
 */
function createTag(name, aggDocUrl, color) {
  name = name.trim();
  if (!name) throw new Error('Tag name cannot be empty.');

  const aggDocId = extractDocId_(aggDocUrl);

  // Verify access before saving.
  try {
    DocumentApp.openById(aggDocId);
  } catch (e) {
    throw new Error(
      'Cannot access the aggregation document. Make sure the URL is correct and you have edit access.'
    );
  }

  const tags = getTags();

  // Prevent duplicate names.
  const duplicate = Object.values(tags).find(t => t.name.toLowerCase() === name.toLowerCase());
  if (duplicate) throw new Error(`A tag named "${name}" already exists.`);

  const tagId = Utilities.getUuid();
  tags[tagId] = {
    name,
    aggregationDocId: aggDocId,
    color: color || TAG_COLORS[0].value,
    createdAt: new Date().toISOString(),
  };
  saveTags_(tags);
  return { tagId, ...tags[tagId] };
}

/**
 * Deletes a tag and removes its named ranges + highlights from the active document.
 */
function deleteTag(tagId) {
  const tags = getTags();
  const tag = tags[tagId];
  if (!tag) throw new Error('Tag not found.');

  // Remove named ranges and clear highlights in the active document.
  const doc = DocumentApp.getActiveDocument();
  doc.getNamedRanges().forEach(nr => {
    if (nr.getName().startsWith(RANGE_PREFIX + tagId + '_')) {
      clearRangeHighlight_(nr.getRange(), null);
      nr.remove();
    }
  });

  delete tags[tagId];
  saveTags_(tags);
}

// ── Tagging ───────────────────────────────────────────────────────────────────

/**
 * Applies a tag to the current document selection.
 * Creates a named range and sets the background color.
 * @returns {string} The created named range ID.
 */
function applyTagToSelection(tagId) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) throw new Error('No text selected. Please select some text first.');

  const tags = getTags();
  const tag = tags[tagId];
  if (!tag) throw new Error('Tag not found.');

  const elements = selection.getRangeElements();
  if (!elements.length) throw new Error('Selection is empty.');

  // Build range.
  const rangeBuilder = doc.newRange();
  elements.forEach(re => {
    if (re.isPartial()) {
      rangeBuilder.addElement(re.getElement(), re.getStartOffset(), re.getEndOffsetInclusive());
    } else {
      rangeBuilder.addElement(re.getElement());
    }
  });

  const rangeId = RANGE_PREFIX + tagId + '_' + Date.now();
  doc.addNamedRange(rangeId, rangeBuilder.build());

  // Apply highlight color.
  applyHighlight_(elements, tag.color);

  return rangeId;
}

/**
 * Removes all tag instances from the active document that match a given tagId.
 * Clears named ranges and highlights.
 */
function removeAllTagInstances(tagId) {
  const doc = DocumentApp.getActiveDocument();
  const prefix = RANGE_PREFIX + tagId + '_';
  doc.getNamedRanges().forEach(nr => {
    if (nr.getName().startsWith(prefix)) {
      clearRangeHighlight_(nr.getRange(), null);
      nr.remove();
    }
  });
}

// ── Content retrieval ─────────────────────────────────────────────────────────

/**
 * Returns all tagged excerpts for a given tag in the active document.
 * @returns {{ rangeId: string, text: string }[]}
 */
function getTaggedContent(tagId) {
  const prefix = RANGE_PREFIX + tagId + '_';
  return DocumentApp.getActiveDocument()
    .getNamedRanges()
    .filter(nr => nr.getName().startsWith(prefix))
    .map(nr => ({ rangeId: nr.getName(), text: extractRangeText_(nr.getRange()) }))
    .filter(item => item.text.trim().length > 0);
}

/**
 * Returns tagged content for ALL tags in the active document.
 * @returns {{ [tagId]: { tag, excerpts: { rangeId, text }[] } }}
 */
function getAllTaggedContent() {
  const tags = getTags();
  const result = {};
  Object.keys(tags).forEach(id => {
    result[id] = { tag: tags[id], excerpts: getTaggedContent(id) };
  });
  return result;
}

// ── Sync ──────────────────────────────────────────────────────────────────────

/**
 * Syncs all excerpts for a given tag to its aggregation document.
 * The section in the aggregation doc is replaced on every sync (idempotent).
 * @returns {{ synced: number }}
 */
function syncTagToAggDoc(tagId) {
  const tags = getTags();
  const tag = tags[tagId];
  if (!tag) throw new Error('Tag not found.');

  const sourceDoc = DocumentApp.getActiveDocument();
  const sourceDocId = sourceDoc.getId();
  const sourceDocName = sourceDoc.getName();
  const sourceDocUrl = sourceDoc.getUrl();
  const excerpts = getTaggedContent(tagId);

  let aggDoc;
  try {
    aggDoc = DocumentApp.openById(tag.aggregationDocId);
  } catch (e) {
    throw new Error('Could not open aggregation document: ' + e.message);
  }

  const body = aggDoc.getBody();
  const startMarker = `[DOCSAGG:${tagId}:${sourceDocId}]`;
  const endMarker = `[/DOCSAGG:${tagId}:${sourceDocId}]`;

  // Remove existing section for this (tag, sourceDoc) pair.
  removeSectionByMarkers_(body, startMarker, endMarker);

  // Write new section (even if empty, so removal is recorded).
  if (excerpts.length > 0) {
    writeSection_(body, startMarker, endMarker, tag, sourceDocName, sourceDocUrl, excerpts);
  }

  return { synced: excerpts.length };
}

/** Syncs all tags that have excerpts in the active document. */
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

// ── Document info ─────────────────────────────────────────────────────────────

function getDocumentInfo() {
  const doc = DocumentApp.getActiveDocument();
  return { name: doc.getName(), id: doc.getId(), url: doc.getUrl() };
}

function getTagColors() {
  return TAG_COLORS;
}

// ── Private helpers ───────────────────────────────────────────────────────────

function extractDocId_(urlOrId) {
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : urlOrId.trim();
}

function extractRangeText_(range) {
  let text = '';
  range.getRangeElements().forEach(re => {
    const el = re.getElement();
    if (el.getText) {
      const full = el.getText();
      if (re.isPartial()) {
        text += full.substring(re.getStartOffset(), re.getEndOffsetInclusive() + 1);
      } else {
        text += full;
      }
    }
  });
  return text;
}

function applyHighlight_(rangeElements, color) {
  rangeElements.forEach(re => {
    const el = re.getElement();
    if (!el.editAsText) return;
    const textEl = el.editAsText();
    const len = el.asText ? el.asText().getText().length : 0;
    if (len === 0) return;
    if (re.isPartial()) {
      textEl.setBackgroundColor(re.getStartOffset(), re.getEndOffsetInclusive(), color);
    } else {
      textEl.setBackgroundColor(0, len - 1, color);
    }
  });
}

function clearRangeHighlight_(range) {
  range.getRangeElements().forEach(re => {
    const el = re.getElement();
    if (!el.editAsText) return;
    const textEl = el.editAsText();
    const len = el.asText ? el.asText().getText().length : 0;
    if (len === 0) return;
    if (re.isPartial()) {
      textEl.setBackgroundColor(re.getStartOffset(), re.getEndOffsetInclusive(), null);
    } else {
      textEl.setBackgroundColor(0, len - 1, null);
    }
  });
}

/**
 * Finds and removes all body children between (and including) the start and end
 * marker paragraphs. Handles the edge case where the body has only one child.
 */
function removeSectionByMarkers_(body, startMarker, endMarker) {
  const toRemove = [];
  let inSection = false;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const t = child.asParagraph().getText();
      if (t === startMarker) { inSection = true; toRemove.push(child); continue; }
      if (t === endMarker)   { toRemove.push(child); inSection = false; continue; }
    }
    if (inSection) toRemove.push(child);
  }

  toRemove.forEach(el => {
    if (body.getNumChildren() > 1) {
      body.removeChild(el);
    } else {
      // Last remaining child — clear it rather than remove.
      try { el.asParagraph().setText(''); } catch (e) {}
    }
  });
}

/**
 * Appends a new tagged section to the aggregation document body.
 */
function writeSection_(body, startMarker, endMarker, tag, sourceDocName, sourceDocUrl, excerpts) {
  // Separator (only if body already has content).
  if (body.getText().trim().length > 0) {
    body.appendHorizontalRule();
  }

  // Hidden start marker (tiny gray text).
  const sP = body.appendParagraph(startMarker);
  sP.editAsText().setFontSize(6).setForegroundColor('#CCCCCC');
  sP.setSpacingAfter(0);

  // Section heading: colored square + tag name + source doc.
  const heading = body.appendParagraph('');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  heading.setSpacingBefore(8);
  const hText = heading.editAsText();
  // Color swatch using a full-width block character.
  hText.insertText(0, '  ' + tag.name + ' — ' + sourceDocName);
  hText.setBackgroundColor(0, 0, tag.color);
  hText.setForegroundColor(1, tag.name.length + sourceDocName.length + 5, '#222222');

  // Source link + timestamp (small italic).
  const meta = body.appendParagraph('Source: ' + sourceDocUrl + '   |   Synced: ' + new Date().toLocaleString());
  meta.editAsText().setFontSize(8).setForegroundColor('#777777').setItalic(true);
  meta.setSpacingAfter(6);

  // Excerpts.
  excerpts.forEach(({ text }) => {
    const p = body.appendParagraph(text);
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setIndentStart(18);
    p.setSpacingBefore(4);
    p.setSpacingAfter(4);
    const pText = p.editAsText();
    pText.setBackgroundColor(tag.color);
    pText.setFontSize(11);
  });

  // Hidden end marker.
  const eP = body.appendParagraph(endMarker);
  eP.editAsText().setFontSize(6).setForegroundColor('#CCCCCC');
  eP.setSpacingBefore(0);
}
