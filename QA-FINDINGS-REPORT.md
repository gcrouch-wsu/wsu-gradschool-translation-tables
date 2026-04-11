# Validator App QA / Code Review Findings

**Scope:** `apps/validate-translation-tables`  
**Date:** 2025-02-06  
**Tests:** `npm run check:validate-translation` — **PASS**

---

## Summary

| Severity | Count |
|----------|-------|
| Critical | 0 |
| High | 0 |
| Medium | 2 |
| Low | 2 |

---

## 1. Campus-Family Rules Upload (JSON/TXT/CSV/XLSX/XLS)

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | `#campus-family-file`, template buttons | `accept=".json,.txt,.csv,.xlsx,.xls"`; Download JSON template, Download CSV template |
| `app.js` | `parseCampusFamilyRulesFile` ~658–720 | Handles JSON, TXT, CSV, XLSX, XLS; normalizes to `{ version, patterns[] }` |
| `app.js` | `normalizeCampusFamilyRule` ~604–623 | `pattern`, `parentKey`, `country`, `state`, `priority`, `enabled` |
| `app.js` | `parseCampusFamilyDelimitedText` ~626–656 | TSV/CSV/pipe-delimited; `=>` pattern; header detection |
| `app.js` | `setupCampusFamilyTemplateDownloads` | Wires template download buttons |
| `validation.js` | `loadFile` | `expectedHeaders` for CSV; `parseCSV`, `sheetToJsonWithHeaderDetection` |

**Finding:** CSV campus-family uses `parseCSV` (returns objects from first row as headers). Files with `pattern`/`parentKey` columns work as expected.

---

## 2. Recommended Review Order Card Removed

**Status:** ⚠️ Minor — Dead code

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | — | `#recommended-review-order` and `#review-order-list` are absent (correct) |
| `app.js` | 2546–2574 | `renderRecommendedReviewOrder` still exists |

**Finding:** `renderRecommendedReviewOrder` is never called. It references `#recommended-review-order` and `#review-order-list`; both return `null`, so it exits early. No runtime errors, but the function is dead.

**Recommendation:** Remove `renderRecommendedReviewOrder` in `app.js` (lines 2546–2574) to avoid confusion.

---

## 3. Error Cards Collapsible, Default Collapsed

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `app.js` | `createErrorCard` ~2618–2712 | Detail panel: `<div id="${cardId}" class="hidden mt-4">` — default collapsed |
| `app.js` | `displayErrorDetails` ~2597–2606 | Toggle buttons with `data-target`; `panel.classList.toggle('hidden')`; chevron rotation |

**Finding:** IDs and toggle wiring are correct. Error cards are collapsed by default.

---

## 4. In-App Review Panel Loading/Progress Indicator

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | 767–773 | `#bulk-load-progress` with text, percent, progress bar |
| `app.js` | 3170–3190, 3218–3219 | `loadProgressWrap.classList.remove('hidden')`; `onProgress` callback updates text/percent/bar |
| `app.js` | `runExportWorkerTask` | Passes `onProgress` to worker |
| `export-worker.js` | 22–29, 1981, 2466, 2510 etc. | `reportProgress(stage, processed)` during `buildValidationExport` |

**Finding:** Progress UI updates when `get_action_queue` runs. `buildValidationExport` with `returnActionQueueOnly: true` still calls `reportProgress` before returning at line 2381.

---

## 5. Progress Behavior — Unresolved Errors Toggle

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | 612 | `#show-unresolved-errors-only` checkbox (checked by default) |
| `app.js` | `buildErrorSamplesFromQueue` ~2124–2151 | Filters by `!row.Decision` when `unresolvedOnly` |
| `app.js` | `buildChartErrorsFromQueue` ~2153–2183 | Same filter |
| `app.js` | `refreshErrorPresentation` ~2199–2210 | Uses queue-derived samples when toggle on + queue exists |
| `app.js` | `updateUnresolvedErrorsToggleState` ~2185–2197 | Disables when no queue; defaults to checked when queue exists |

**Finding:** Making decisions reduces visible unresolved counts via `Decision` filtering. Base validation stats (`stats.errors`) are unchanged.

---

## 6. Duplicate-Target Edit — Effective myWSU Preview

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | 855 | `Effective myWSU (Preview)` column header |
| `app.js` | `formatPreview` ~2969 | Formats preview entry with badge |
| `app.js` | ~3086–3126 | `effectivePreview` from `wsuKeyLookup`, manual key, or selected candidate |

**Finding:** Preview logic uses candidate/manual key + key lookup for full match context.

---

## 7. Edit Only Translation-Table Rows

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | `#bulk-filter-translation-only` | Checkbox in bulk panel |
| `app.js` | 3039 | `filterTranslationOnly?.checked` |
| `app.js` | 3045 | `Source_Sheet === 'Missing_Mappings'` or `Error_Type === 'Missing_Mapping'` → excluded |

**Finding:** Non-translation/inferred rows (e.g., Missing_Mappings) are excluded when the filter is checked.

---

## 8. Reduce Horizontal Scrolling — Bulk Panel Width

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | 723 | `#bulk-edit-panel` uses `lg:w-[calc(100vw-2rem)]` and `lg:left-1/2 lg:-translate-x-1/2` |
| `index.html` | 847 | `table class="w-full text-sm table-fixed"` |

**Finding:** Panel is wide and full-screen oriented on desktop. `table-fixed` helps layout.

---

## 9. Session JSON Resume

**Status:** ✅ Implemented

| File | Lines | Notes |
|------|-------|-------|
| `index.html` | 233–256 | `#session-upload-card`, `#session-upload-file`, `#session-upload-status` |
| `app.js` | 297, 306 | `sessionUploadCard` shown in validate mode |
| `app.js` | 467–493 | `setupSessionUploadCard`; parses JSON; `uploadedSessionRows`; `applySessionDataToActionQueue` when queue exists |
| `app.js` | 3209–3211 | `uploadedSessionRows` applied after queue load if not yet applied |
| `app.js` | `bulk-save-session-btn`, `bulk-load-session-file` | Save/Load in panel |

**Finding:** Top-level Review Session JSON upload; applies immediately if queue exists, or after queue load. Save/Load in panel shares logic.

---

## Additional Checks

### No Broken References from Removed Recommended Review Order

- `renderRecommendedReviewOrder` is never called. No DOM references to removed elements exist elsewhere.

### No Syntax/Runtime Errors

- No syntax errors in `app.js`, `validation.js`, `worker.js`, `export-worker.js`.
- No obvious runtime errors from removed elements.

### Tests

- `npm run check:validate-translation` passes (checks + export tests).

### Coverage Gaps

- No automated tests for UI flows (e.g., campus-family file upload, session JSON upload, bulk filter toggle).
- No tests for `parseCampusFamilyRulesFile` or `parseCampusFamilyDelimitedText` directly.

---

## Recommendations

| Priority | Action |
|----------|--------|
| Low | Remove dead `renderRecommendedReviewOrder` in `app.js` (2546–2574) |
| Low | Add unit tests for `parseCampusFamilyRulesFile` and `parseCampusFamilyDelimitedText` |
| Info | Optional: Add browser smoke test for typical Validate flow (upload → validate → bulk edit) |

---

## Reviewer Workflow Concerns

| Concern | Risk |
|---------|------|
| Session JSON applied before queue load | If user uploads session JSON before queue exists, it applies after queue load. Order is correct. |
| Progress bar during `get_action_queue` | Worker sends progress; UI should update. If build is very fast, bar may flash briefly. |
| `show-unresolved-errors-only` default | When queue exists, it defaults to checked. Reviewers may expect to see all errors initially; consider UX if feedback suggests confusion. |
