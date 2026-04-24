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

/** Syncs all defined tags that appear in the active document. */
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

// ── Private: aggregation doc section helpers ──────────────────────────────────

/**
 * Creates or updates a tagged section in the aggregation document.
 *
 * On update: clears only the content between the marker paragraphs (markers
 * are never removed), then re-inserts fresh content.
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

  let insertIdx;

  if (startIdx !== -1 && endIdx !== -1 && startIdx < endIdx) {
    // Section exists — clear content between markers.
    for (let i = endIdx - 1; i > startIdx; i--) {
      body.removeChild(body.getChild(i));
    }
    insertIdx = startIdx + 1;
  } else {
    // No section — append markers at the end.
    if (body.getText().trim().length > 0) {
      body.appendHorizontalRule();
    }

    const sP = body.appendParagraph(startMarker);
    sP.editAsText().setFontSize(1).setForegroundColor('#FFFFFF');
    sP.setSpacingAfter(0);

    const eP = body.appendParagraph(endMarker);
    eP.editAsText().setFontSize(1).setForegroundColor('#FFFFFF');
    eP.setSpacingBefore(0);

    insertIdx = body.getNumChildren() - 1;
  }

  // Timestamp / source line.
  const ts       = getTimestampSettings();
  const filePath = getFilePath_(sourceDocId, ts.parentFolders) || sourceDocName;
  const linkText = '[' + filePath + ']';
  const tagText  = '[' + tag.name + ']';
  const metaStr  = linkText + '   ' + tagText + '   ' + formatTimestamp_(new Date(), ts.dateFormat);
  const meta     = body.insertParagraph(insertIdx++, metaStr);
  const metaTxt  = meta.editAsText();
  metaTxt.setFontSize(ts.fontSize).setForegroundColor(ts.textColor);
  metaTxt.setLinkUrl(0, linkText.length - 1, sourceDocUrl);
  metaTxt.setForegroundColor(0, linkText.length - 1, ts.linkColor);
  meta.setSpacingAfter(6);

  // Excerpts — element.copy() preserves inline formatting, glyph type, and
  // list structure. Nesting level is normalized relative to the [tag] marker.
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
 * Removes an entire section (markers + content) when there are no excerpts.
 * Uses high-to-low index removal so indices stay valid after each removal.
 * Appends a blank paragraph first if removal would leave the body empty.
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

  let paragraphsOutside = 0;
  for (let i = 0; i < body.getNumChildren(); i++) {
    if (i >= startIdx && i <= endIdx) continue;
    if (body.getChild(i).getType() === DocumentApp.ElementType.PARAGRAPH) paragraphsOutside++;
  }

  if (paragraphsOutside === 0) body.appendParagraph('');

  for (let i = endIdx; i >= startIdx; i--) {
    body.removeChild(body.getChild(i));
  }
}
