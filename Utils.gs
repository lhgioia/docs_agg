// ── Shared utilities ──────────────────────────────────────────────────────────

// Per-execution cache — reset automatically between google.script.run calls.
const filePathCache_ = {};

/**
 * Returns the file path up to `depth` parent directories deep, e.g.
 * depth=2 → "Grandparent/Parent/Filename".
 * Falls back to just the filename if parent folders are not accessible.
 * Results are memoized for the duration of the current execution.
 */
function getFilePath_(fileId, depth) {
  if (depth === undefined) depth = 2;
  const key = fileId + ':' + depth;
  if (key in filePathCache_) return filePathCache_[key];
  let result;
  try {
    const file = DriveApp.getFileById(fileId);
    const name = file.getName();
    if (depth === 0) {
      result = name;
    } else {
      const p1It = file.getParents();
      if (!p1It.hasNext()) {
        result = name;
      } else {
        const p1 = p1It.next();
        if (depth === 1) {
          result = p1.getName() + '/' + name;
        } else {
          const p2It = p1.getParents();
          result = p2It.hasNext()
            ? p2It.next().getName() + '/' + p1.getName() + '/' + name
            : p1.getName() + '/' + name;
        }
      }
    }
  } catch (e) {
    result = null;
  }
  filePathCache_[key] = result;
  return result;
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
