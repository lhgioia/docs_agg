// ── Shared utilities ──────────────────────────────────────────────────────────

/**
 * Returns the file path up to `depth` parent directories deep, e.g.
 * depth=2 → "Grandparent/Parent/Filename".
 * Falls back to just the filename if parent folders are not accessible.
 */
function getFilePath_(fileId, depth) {
  if (depth === undefined) depth = 2;
  try {
    const file = DriveApp.getFileById(fileId);
    const name = file.getName();
    if (depth === 0) return name;
    const p1It = file.getParents();
    if (!p1It.hasNext()) return name;
    const p1 = p1It.next();
    if (depth === 1) return p1.getName() + '/' + name;
    const p2It = p1.getParents();
    if (!p2It.hasNext()) return p1.getName() + '/' + name;
    return p2It.next().getName() + '/' + p1.getName() + '/' + name;
  } catch (e) {
    return null;
  }
}

/** Formats a Date for the aggregation doc timestamp line. */
function formatTimestamp_(date, format) {
  if (format === 'iso') {
    const pad = n => String(n).padStart(2, '0');
    return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) +
           ' ' + pad(date.getHours()) + ':' + pad(date.getMinutes());
  }
  if (format === 'date-only') {
    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  }
  if (format === 'short') {
    let h = date.getHours();
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    const m = String(date.getMinutes()).padStart(2, '0');
    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear() +
           ' ' + h + ':' + m + ' ' + ampm;
  }
  return date.toLocaleString().replace(',', '');
}

/** Extracts a Google Docs file ID from a URL, or returns the value as-is. */
function extractDocId_(urlOrId) {
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : urlOrId.trim();
}
