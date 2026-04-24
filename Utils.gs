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

/**
 * Formats a Date for the aggregation doc timestamp line.
 * @param {Date}   date
 * @param {string} format   — 'locale' | 'short' | 'date-only' | 'iso'
 * @param {string} timezone — IANA timezone string (e.g. 'America/Los_Angeles')
 */
function formatTimestamp_(date, format, timezone) {
  const tz = timezone || Session.getScriptTimeZone();
  if (format === 'iso')       return Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');
  if (format === 'date-only') return Utilities.formatDate(date, tz, 'M/d/yyyy');
  // 'short' and 'locale' both show date + time
  return Utilities.formatDate(date, tz, 'M/d/yyyy h:mm a');
}

/** Extracts a Google Docs file ID from a URL, or returns the value as-is. */
function extractDocId_(urlOrId) {
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : urlOrId.trim();
}
