const TIMESTAMP_SETTINGS_KEY = 'DOCSAGG_TS_V1';
const TIMESTAMP_DEFAULTS_ = {
  fontSize: 11,
  parentFolders: 2,
  linkColor: '#1155CC',
  textColor: '#777777',
  dateFormat: 'locale',
};

// ── Timestamp settings ────────────────────────────────────────────────────────

function getTimestampSettings() {
  const raw = PropertiesService.getUserProperties().getProperty(TIMESTAMP_SETTINGS_KEY);
  return Object.assign({}, TIMESTAMP_DEFAULTS_, raw ? JSON.parse(raw) : {});
}

function saveTimestampSettings(settings) {
  PropertiesService.getUserProperties().setProperty(TIMESTAMP_SETTINGS_KEY, JSON.stringify(settings));
  return getTimestampSettings();
}
