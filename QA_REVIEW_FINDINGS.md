# QA / Code Review Findings: Validate Translation Tables App

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
| Info | 2 |

---

## 1. Campus-Family Rules Upload (JSON/TXT/CSV/XLSX/XLS)

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| File input | `index.html` ~L268 | `#campus-family-file` with `accept=".json,.txt,.csv,.xlsx,.xls"` |
| Template buttons | `index.html` ~L273–278 | JSON/TXT/CSV template download buttons |
| Parser | `app.js` `parseCampusFamilyRulesFile` | Handles JSON, TXT, CSV, XLSX, XLS |
| Normalizer | `app.js` `normalizeCampusFamilyRule`, `parseCampusFamilyDelimitedText` | Produces `{ version, patterns[] }` |
| Template setup | `app.js` `setupCampusFamilyTemplateDownloads` | Wires template downloads |

**Finding (Low):** CSV parsing uses `loadFile` + `parseCSV` with first row as headers. CSV with `pattern,parentKey` (or similar) works; ensure template matches expected column names.

---

## 2. Recommended Review Order Card Removed

**Status:** ⚠️ Dead code remains

| Item | Location | Notes |
|------|----------|------|
| Dead function | `app.js` L2546–2574 | `renderRecommendedReviewOrder(errors)` |
| DOM refs | L2547–2548 | `#recommended-review-order`, `#review-order-list` — not in `index.html` |
| Call sites | — | Function is **never called** |

**Finding (Medium):** No runtime errors (function returns early when elements are null). Remove `renderRecommendedReviewOrder` to avoid confusion and future misuse.

---

## 3. Error Cards Collapsible, Default Collapsed

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Toggle wiring | `app.js` `displayErrorDetails`, `createErrorCard` | Detail panels use `class="hidden"` |
| Toggle buttons | `createErrorCard` | `error-card-toggle` with `data-target` |
| Default state | `createErrorCard` | Detail panels start collapsed |

---

## 4. In-App Review Panel Loading/Progress Indicator

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| UI | `index.html` | `#bulk-load-progress` with text, percent, bar |
| Show/hide | `app.js` L3170–3174, L3218–3220 | Shown at load start, hidden in `finally` |
| Integration | `app.js` L3189–3199 | `runExportWorkerTask('get_action_queue', ..., onProgress)` |
| Worker progress | `export-worker.js` L22–29, L1981, L2466, L2510, etc. | `reportProgress` called during build |

---

## 5. Progress Behavior (Unresolved Errors Toggle)

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Toggle | `index.html` | `#show-unresolved-errors-only` |
| Queue sampling | `app.js` L2124–2151 | `buildErrorSamplesFromQueue` |
| Chart from queue | `app.js` L2153–2183 | `buildChartErrorsFromQueue` |
| Refresh flow | `app.js` L2199–2210 | `refreshErrorPresentation` |
| Toggle state | `app.js` L2185–2197 | `updateUnresolvedErrorsToggleState` |

**Finding (Info):** When `showUnresolvedErrorsOnly` is true and queue exists, error counts/charts reflect queue-derived unresolved counts. Base validation stats (`stats.validation`) are unchanged.

---

## 6. Duplicate-Target Edit UX — Effective myWSU Preview

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Column | Bulk edit table | "Effective myWSU (Preview)" |
| Preview logic | `app.js` `formatPreview`, `effectivePreview` | Uses `wsuKeyLookup` or manual/selected candidate |

---

## 7. Edit Only Translation-Table Rows

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Filter | `index.html` | `#bulk-filter-translation-only` |
| Logic | `app.js` `getWorkingRows` / `filterTranslationOnly` | Excludes `Source_Sheet === 'Missing_Mappings'` or `Error_Type === 'Missing_Mapping'` |

---

## 8. Reduce Horizontal Scrolling — Bulk Panel Width

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Panel | `index.html` | `#bulk-edit-panel` with `lg:w-[calc(100vw-2rem)]` |
| Table | Bulk table | `table-fixed` layout |

---

## 9. Session JSON Resume

**Status:** ✅ Implemented

| Item | Location | Notes |
|------|----------|------|
| Upload card | `index.html` L233–256 | `#session-upload-card`, `#session-upload-file` |
| Visibility | `app.js` L304–306 | Shown in validate mode only |
| Apply flow | `app.js` L467–493, L3209–3211 | `setupSessionUploadCard`; applies when queue exists or after load |
| Save/Load | `app.js` | `bulk-save-session-btn`, `bulk-load-session-file` share logic with upload |

---

## Additional Verification

### Broken References from Removed Section

- **Finding (Low):** `renderRecommendedReviewOrder` references removed DOM elements but is never invoked. No broken references at runtime.

### Syntax/Runtime Errors

- No syntax errors observed in reviewed files.
- **Recommendation:** Manually run typical Validate flow (upload outcomes, translate, myWSU → Validate → Bulk edit) and confirm no console errors.

### Test Coverage

- `npm run check:validate-translation` — **PASS** (all checks + export tests).
- **Coverage gaps:** No automated UI/E2E tests for bulk panel, session upload, campus-family upload, or collapsible error cards.

---

## Findings by Severity

### Medium

1. **Dead code: `renderRecommendedReviewOrder`** — `app.js` L2546–2574  
   - Remove function or refactor if Recommended Review Order is reintroduced.

2. **Progress callback timing** — `export-worker.js`  
   - `get_action_queue` runs `buildValidationExport` with `returnActionQueueOnly: true`, which returns early at L2381. Progress is reported at L1981, L2466, L2510 before that. Verify in browser that progress bar updates during queue load; if not, add explicit `reportProgress` calls earlier in the build path.

### Low

3. **CSV campus-family format** — `validation.js` `loadFile`  
   - Ensure CSV template uses headers that `parseCampusFamilyDelimitedText` expects (`pattern`, `parentKey` or equivalent).

4. **No E2E/UI tests** — `apps/validate-translation-tools`  
   - Add Playwright or similar tests for Validate flow, bulk panel, session upload, and campus-family upload.

### Info

5. **Base stats vs queue-derived counts** — `refreshErrorPresentation`  
   - When `showUnresolvedErrorsOnly` is on, error details/chart use queue-derived counts. Base validation stats remain correct; UX is intentional.

6. **Session apply order** — `app.js` L3209–3211  
   - Uploaded session applies after queue load if not yet applied. Order is correct.

---

## Reviewer Workflow Concerns

- **Session upload before queue load:** User can upload session JSON before opening bulk panel. Session applies when queue loads. If user expects immediate feedback, consider a status message like "Session queued; will apply when review rows load."
- **Translation-only filter:** Excludes Missing_Mappings. Confirm this matches reviewer expectations (e.g., some may want to see inferred rows for context).

---

## Recommended Actions

1. Remove `renderRecommendedReviewOrder` (or document as intentionally retained for future use).
2. Manually verify bulk-load progress bar updates during queue load.
3. Add E2E tests for Validate flow and new features.
4. Document expected CSV format for campus-family rules in UI or template.
