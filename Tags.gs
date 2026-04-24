const TAGS_PROP_KEY = 'DOCSAGG_TAGS_V1';

// ── Tag metadata CRUD ─────────────────────────────────────────────────────────

/** Returns the tags map: { [tagId]: { name, aggregationDocId, aggregationDocName, createdAt, updatedAt } } */
function getTags() {
  const raw = PropertiesService.getUserProperties().getProperty(TAGS_PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveTags_(tags) {
  PropertiesService.getUserProperties().setProperty(TAGS_PROP_KEY, JSON.stringify(tags));
}

/**
 * Creates or updates a tag definition.
 * Upserts by name (case-insensitive); updates the aggregation doc on match.
 */
function createTag(name, aggDocUrl) {
  name = name.trim();
  if (!name) throw new Error('Tag name cannot be empty.');

  const aggDocId = extractDocId_(aggDocUrl);

  let aggDocName;
  try {
    aggDocName = DocumentApp.openById(aggDocId).getName();
  } catch (e) {
    throw new Error('Cannot access the aggregation document. Check the URL and your edit access.');
  }

  const tags = getTags();

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

/** Removes a tag definition. Does not modify the source document. */
function deleteTag(tagId) {
  const tags = getTags();
  if (!tags[tagId]) throw new Error('Tag not found.');
  delete tags[tagId];
  saveTags_(tags);
}

// ── Document tag summary ──────────────────────────────────────────────────────

/**
 * Returns all tags found in the document merged with their definitions.
 * Also applies hyperlinks to [tag] markers for defined tags.
 * This is the primary data source for the sidebar.
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

  applyTagMarkerLinks_(nameToTag);

  return { inDocument, notInDocument };
}

/**
 * For each [tagname] list item whose tag has a linked aggregation doc,
 * sets a hyperlink on the marker text pointing to that doc.
 * Skips markers whose link is already correct to avoid unnecessary writes.
 * Safe on every scan — getText() ignores link attributes so the scan regex
 * continues to match unchanged.
 */
function applyTagMarkerLinks_(nameToTag) {
  const body = DocumentApp.getActiveDocument().getBody();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;
    const item = child.asListItem();
    const raw = item.getText();
    const tagMatch = raw.trim().match(/^\[([^\]]+)\]$/);
    if (!tagMatch) continue;
    const linked = nameToTag[tagMatch[1].toLowerCase()];
    if (!linked || !linked.aggregationDocId) continue;
    const url = 'https://docs.google.com/document/d/' + linked.aggregationDocId + '/edit';
    const textEl = item.editAsText();
    if (textEl.getLinkUrl(0) === url) continue;
    textEl.setLinkUrl(0, raw.length - 1, url);
  }
}
