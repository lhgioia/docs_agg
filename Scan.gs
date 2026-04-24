// ── Document scanning ─────────────────────────────────────────────────────────

/**
 * Scans the active document for [tagname] list items and collects their
 * deeper-nested children as plain text excerpts.
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
 * Same scan logic as scanDocumentForTags_ but stores ListItem element
 * references and relative nesting levels instead of plain text.
 * Used server-side only — element references cannot be serialized for the sidebar.
 *
 * "level" is relative to the tag marker: a direct child is 0, grandchild is 1, etc.
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

/** Returns the tagged excerpts for a given tagId as { text, element, level }[]. */
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
