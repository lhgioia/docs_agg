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
 * Tag metadata is stored in UserProperties; the document is the source of
 * truth for which tags exist and what content they contain.
 *
 * File layout:
 *   Code.gs     — Add-on lifecycle (onOpen, onInstall, showSidebar)
 *   Tags.gs     — Tag CRUD, getDocumentTagSummary, applyTagMarkerLinks_
 *   Scan.gs     — Document scanning functions
 *   Sync.gs     — Sync to aggregation document
 *   Settings.gs — Timestamp display settings
 *   Utils.gs    — Shared helpers (getFilePath_, formatTimestamp_, extractDocId_)
 */

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
