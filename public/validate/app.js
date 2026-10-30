let filesUploaded = {
    outcomes: false,
    translate: false,
    wsu_org: false
};

let fileObjects = {
    outcomes: null,
    translate: null,
    wsu_org: null
};
let priorDecisions = null;
let campusFamilyRules = null;
let preEditedActionQueueRows = null;

let loadedData = {
    outcomes: [],
    translate: [],
    wsu_org: []
};

let validatedData = [];
let missingData = [];
let stats = {};
let selectedColumns = {
    outcomes: [],
    wsu_org: []
};
let columnRoles = {
    outcomes: {},
    wsu_org: {}
};
let showAllErrors = false;
let showUnresolvedErrorsOnly = true;
let keyConfig = {
    outcomes: '',
    translateInput: '',
    translateOutput: '',
    wsu: ''
};
let currentMode = 'validate';
let matchMethod = 'key';
let keyLabels = {
    outcomes: '',
    translateInput: '',
    translateOutput: '',
    wsu: ''
};
let matchMethodTouched = false;
let validateNameMode = 'key+name';
let lastNameCompareConfig = {
    enabled: false,
    outcomes: '',
    wsu: '',
    threshold: 0.8,
    ambiguity_gap: 0.03,
    state_outcomes: '',
    state_wsu: '',
    city_outcomes: '',
    city_wsu: '',
    country_outcomes: '',
    country_wsu: ''
};
let debugState = {
    outcomes: null,
    translate: null,
    wsu_org: null
};
let activeWorker = null;
let activeWorkerReject = null;
let activeExportWorker = null;
let activeExportWorkerReject = null;
let pageBusy = false;
let runLocked = false;
let uploadedSessionRows = null;
let uploadedSessionApplied = false;
let actionQueuePrefetchPromise = null;
let actionQueuePrefetchInFlight = false;
// One-way flag for the current page lifecycle; reset by Start Over/full reload.
let bulkEditorOpenedOnce = false;

function beforeUnloadHandler(event) {
    event.preventDefault();
    event.returnValue = '';
}

function setPageBusy(isBusy) {
    if (isBusy && !pageBusy) {
        window.addEventListener('beforeunload', beforeUnloadHandler);
        pageBusy = true;
        return;
    }
    if (!isBusy && pageBusy) {
        window.removeEventListener('beforeunload', beforeUnloadHandler);
        pageBusy = false;
    }
}

function hideLoadingUI() {
    const loading = document.getElementById('loading');
    const progressWrap = document.getElementById('loading-progress');
    if (loading) {
        loading.classList.add('hidden');
    }
    if (progressWrap) {
        progressWrap.classList.add('hidden');
    }
}

function applyPrimaryActionDisabledState(button, disabled) {
    if (!button) return;
    button.disabled = disabled;
    if (disabled) {
        button.classList.add('bg-gray-400', 'cursor-not-allowed');
        button.classList.remove('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
        return;
    }
    button.classList.remove('bg-gray-400', 'cursor-not-allowed');
    button.classList.add('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
}

function setPrimaryActionsBusy(isBusy) {
    applyPrimaryActionDisabledState(document.getElementById('validate-btn'), isBusy);
    applyPrimaryActionDisabledState(document.getElementById('generate-btn'), isBusy);
    applyPrimaryActionDisabledState(document.getElementById('join-preview-btn'), isBusy);
}

function setRunLock(isLocked) {
    runLocked = Boolean(isLocked);
    setPrimaryActionsBusy(runLocked || pageBusy);
}

document.addEventListener('DOMContentLoaded', function() {
    setupModeSelector();
    setupFileUploads();
    setupSessionUploadCard();
    setupCampusFamilyUpload();
    loadCampusFamilyBaseline();
    setupColumnSelection();
    setupValidateButton();
    setupGenerateButton();
    setupJoinPreviewButton();
    setupMatchMethodControls();
    setupValidateNameModeControls();
    setupDebugToggle();
    setupDownloadButton();
    setupResetButton();
    setupBulkEditPanel();
    setupNameCompareControls();
    setupShowAllErrorsToggle();
    setupMappingLogicPreviewToggle();
    setupMatchingRulesToggle();
    updateModeUI();
});

function runWorkerTask(type, payload, onProgress) {
    return new Promise((resolve, reject) => {
        if (activeWorker) {
            activeWorker.terminate();
            if (typeof activeWorkerReject === 'function') {
                activeWorkerReject(new Error('Previous task was cancelled.'));
            }
            activeWorkerReject = null;
            activeWorker = null;
        }
        const worker = new Worker('worker.js');
        activeWorker = worker;
        activeWorkerReject = reject;

        worker.onmessage = (event) => {
            const message = event.data || {};
            if (message.type === 'progress') {
                if (onProgress) onProgress(message.stage, message.processed, message.total);
                return;
            }
            if (message.type === 'result') {
                worker.terminate();
                activeWorker = null;
                activeWorkerReject = null;
                resolve(message.result);
                return;
            }
            if (message.type === 'error') {
                worker.terminate();
                activeWorker = null;
                activeWorkerReject = null;
                reject(new Error(message.message));
            }
        };
        worker.onerror = (event) => {
            worker.terminate();
            activeWorker = null;
            activeWorkerReject = null;
            reject(new Error(event.message || 'Worker error'));
        };
        worker.postMessage({ type, payload });
    });
}

function runExportWorkerTask(type, payload, onProgress) {
    return new Promise((resolve, reject) => {
        if (activeExportWorker) {
            activeExportWorker.terminate();
            if (typeof activeExportWorkerReject === 'function') {
                activeExportWorkerReject(new Error('Previous export task was cancelled.'));
            }
            activeExportWorkerReject = null;
            activeExportWorker = null;
        }
        const worker = new Worker('export-worker.js');
        activeExportWorker = worker;
        activeExportWorkerReject = reject;

        worker.onmessage = (event) => {
            const message = event.data || {};
            if (message.type === 'progress') {
                if (onProgress) onProgress(message.stage, message.processed, message.total);
                return;
            }
            if (message.type === 'result') {
                worker.terminate();
                activeExportWorker = null;
                activeExportWorkerReject = null;
                resolve(message.result);
                return;
            }
            if (message.type === 'error') {
                worker.terminate();
                activeExportWorker = null;
                activeExportWorkerReject = null;
                const err = new Error(message.message);
                if (message.stack) err.exportStack = message.stack;
                reject(err);
            }
        };
        worker.onerror = (event) => {
            worker.terminate();
            activeExportWorker = null;
            activeExportWorkerReject = null;
            reject(new Error(event.message || 'Export worker error'));
        };
        worker.postMessage({ type, payload });
    });
}

function cloneActionQueueRows(rows) {
    return (Array.isArray(rows) ? rows : []).map(row => ({
        ...row,
        _candidates: Array.isArray(row._candidates) ? row._candidates.map(c => ({ ...c })) : []
    }));
}

function normalizeBulkReviewScope(value) {
    const raw = String(value || '').trim().toLowerCase();
    if (raw === 'translation_only' || raw === 'missing_only') return raw;
    return 'all';
}

function getBulkReviewScope() {
    return normalizeBulkReviewScope(document.getElementById('bulk-filter-review-scope')?.value);
}

function buildActionQueuePayload() {
    return {
        validated: validatedData,
        missing: missingData,
        selectedCols: selectedColumns,
        priorDecisions: priorDecisions || null,
        options: {
            includeSuggestions: Boolean(document.getElementById('include-suggestions')?.checked),
            showMappingLogic: Boolean(document.getElementById('show-mapping-logic')?.checked),
            nameCompareConfig: lastNameCompareConfig,
            campusFamilyRules: campusFamilyRules || null
        },
        context: { loadedData, columnRoles, keyConfig, keyLabels }
    };
}

function cancelActionQueuePrefetch() {
    actionQueuePrefetchPromise = null;
    if (actionQueuePrefetchInFlight && activeExportWorker) {
        activeExportWorker.terminate();
        if (typeof activeExportWorkerReject === 'function') {
            activeExportWorkerReject(new Error('Action queue prefetch was cancelled.'));
        }
        activeExportWorker = null;
        activeExportWorkerReject = null;
    }
    actionQueuePrefetchInFlight = false;
}

function startActionQueuePrefetch() {
    if (!Array.isArray(validatedData) || !validatedData.length) return null;
    if (actionQueueRowsCache && actionQueueRowsCache.length) return Promise.resolve(actionQueueRowsCache);
    if (actionQueuePrefetchPromise) return actionQueuePrefetchPromise;

    const prefetchPromise = runExportWorkerTask('get_action_queue', buildActionQueuePayload())
        .then((result) => {
            const queueRows = cloneActionQueueRows(result?.actionQueueRows || []);
            actionQueueRowsCache = queueRows;
            preEditedActionQueueRows = cloneActionQueueRows(queueRows);
            return queueRows;
        })
        .catch((error) => {
            console.warn('Action queue prefetch failed:', error);
            return null;
        })
        .finally(() => {
            if (actionQueuePrefetchPromise === prefetchPromise) {
                actionQueuePrefetchPromise = null;
            }
            actionQueuePrefetchInFlight = false;
        });

    actionQueuePrefetchInFlight = true;
    actionQueuePrefetchPromise = prefetchPromise;
    return prefetchPromise;
}

function downloadArrayBuffer(buffer, filename) {
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

function setupModeSelector() {
    const validateRadio = document.getElementById('mode-validate');
    const createRadio = document.getElementById('mode-create');
    const joinPreviewRadio = document.getElementById('mode-join-preview');
    if (!validateRadio || !createRadio) return;

    const handleModeChange = () => {
        if (validateRadio.checked) currentMode = 'validate';
        else if (createRadio.checked) currentMode = 'create';
        else if (joinPreviewRadio && joinPreviewRadio.checked) currentMode = 'join-preview';
        else currentMode = 'validate';
        updateModeUI();
        processAvailableFiles();
    };

    validateRadio.addEventListener('change', handleModeChange);
    createRadio.addEventListener('change', handleModeChange);
    if (joinPreviewRadio) joinPreviewRadio.addEventListener('change', handleModeChange);
}

function updateModeUI() {
    const translateCard = document.getElementById('translate-upload-card');
    const outcomesCard = document.getElementById('outcomes-upload-card');
    const validateAction = document.getElementById('validate-action');
    const generateAction = document.getElementById('generate-action');
    const joinPreviewAction = document.getElementById('join-preview-action');
    const instructionsValidate = document.getElementById('instructions-validate');
    const instructionsCreate = document.getElementById('instructions-create');
    const instructionsJoinPreview = document.getElementById('instructions-join-preview');
    const columnSelection = document.getElementById('column-selection');
    const nameCompare = document.getElementById('name-compare');
    const matchMethodSection = document.getElementById('match-method');
    const validateNameModeSection = document.getElementById('validate-name-mode');
    const validationOptions = document.getElementById('validation-options');
    const toggleColumns = document.getElementById('toggle-columns');
    const columnCheckboxes = document.getElementById('column-checkboxes');
    const translateInputGroup = document.getElementById('translate-input-group');
    const translateOutputGroup = document.getElementById('translate-output-group');
    const keyMatchFields = document.getElementById('key-match-fields');

    const priorValidateCard = document.getElementById('prior-validate-card');
    const sessionUploadCard = document.getElementById('session-upload-card');
    const campusFamilyCard = document.getElementById('campus-family-card');
    if (translateCard) {
        translateCard.classList.toggle('hidden', currentMode === 'create');
    }
    if (priorValidateCard) {
        priorValidateCard.classList.toggle('hidden', currentMode !== 'validate');
    }
    if (sessionUploadCard) {
        sessionUploadCard.classList.toggle('hidden', currentMode !== 'validate');
    }
    if (campusFamilyCard) {
        campusFamilyCard.classList.toggle('hidden', currentMode !== 'validate');
    }
    if (outcomesCard) {
        outcomesCard.classList.remove('hidden');
    }
    if (validateAction) {
        validateAction.classList.toggle('hidden', currentMode !== 'validate');
    }
    if (generateAction) {
        generateAction.classList.toggle('hidden', currentMode !== 'create');
    }
    if (joinPreviewAction) {
        joinPreviewAction.classList.toggle('hidden', currentMode !== 'join-preview');
    }
    if (instructionsValidate && instructionsCreate && instructionsJoinPreview) {
        instructionsValidate.classList.toggle('hidden', currentMode !== 'validate');
        instructionsCreate.classList.toggle('hidden', currentMode !== 'create');
        instructionsJoinPreview.classList.toggle('hidden', currentMode !== 'join-preview');
    }
    if (validationOptions) {
        validationOptions.classList.toggle('hidden', currentMode === 'join-preview');
    }
    const columnCheckboxesEl = document.getElementById('column-checkboxes');
    if (columnCheckboxesEl) {
        columnCheckboxesEl.classList.toggle('join-preview-mode', currentMode === 'join-preview');
    }
    if (columnSelection) {
        if (currentMode === 'create') {
            columnSelection.classList.remove('hidden');
        }
    }
    if (nameCompare) {
        nameCompare.classList.toggle('hidden', currentMode === 'join-preview');
    }
    if (matchMethodSection) {
        matchMethodSection.classList.toggle('hidden', currentMode !== 'create');
    }
    if (validateNameModeSection) {
        validateNameModeSection.classList.toggle('hidden', currentMode !== 'validate');
    }
    if (translateInputGroup) {
        translateInputGroup.classList.toggle('hidden', currentMode === 'create');
    }
    if (translateOutputGroup) {
        translateOutputGroup.classList.toggle('hidden', currentMode === 'create');
    }
    if (toggleColumns) toggleColumns.classList.remove('hidden');
    if (columnCheckboxes) columnCheckboxes.classList.remove('hidden');
    if (keyMatchFields) {
        keyMatchFields.classList.toggle('hidden', currentMode === 'create');
    }
    updateMatchMethodUI();
    updateValidateNameModeUI();
    renderDebugPanel();
}

const encodingSelectIds = {
    outcomes: 'outcomes-encoding',
    translate: 'translate-encoding',
    wsu_org: 'wsu-org-encoding'
};

const sheetHeaderHintsByFileKey = {
    outcomes: ['name', 'mdb_code', 'state', 'country'],
    translate: ['input', 'output', 'translate_input', 'translate_output'],
    wsu_org: ['org id', 'descr', 'city', 'state', 'country']
};

function getFileEncoding(fileKey) {
    const select = document.getElementById(encodingSelectIds[fileKey]);
    return select ? select.value : 'auto';
}

function getSheetHeaderHints(fileKey) {
    return sheetHeaderHintsByFileKey[fileKey] || [];
}

function parsePriorValidateWorkbook(file) {
    return new Promise((resolve, reject) => {
        const fileName = String(file?.name || '').toLowerCase();
        if (!fileName || (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls'))) {
            resolve(null);
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = typeof XLSX !== 'undefined' ? XLSX.read(data, { type: 'array' }) : null;
                if (!workbook || !workbook.SheetNames) {
                    resolve(null);
                    return;
                }
                const sheetName = workbook.SheetNames.find(n => n === 'Review_Workbench');
                if (!sheetName) {
                    resolve(null);
                    return;
                }
                const sheet = workbook.Sheets[sheetName];
                if (!sheet) {
                    resolve(null);
                    return;
                }
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
                if (!rows || rows.length < 2) {
                    resolve({ priorDecisions: {}, rowCount: 0 });
                    return;
                }
                const headers = rows[0].map(h => String(h || '').trim());
                const findCol = (patterns) => {
                    for (const p of patterns) {
                        const i = headers.findIndex(h => {
                            const str = String(h ?? '');
                            return typeof p === 'string' ? str.toLowerCase().includes(p.toLowerCase()) : (p.test ? p.test(str) : false);
                        });
                        if (i >= 0) return i;
                    }
                    return -1;
                };
                const ridCol = findCol(['Review Row ID', /review\s*row\s*id/i]);
                const decCol = findCol(['Decision', /^decision$/i]);
                const manCol = findCol(['Manual Key', /manual.*key/i]);
                const reasonCol = findCol(['Reason Code', /reason\s*code/i]);
                const curInCol = findCol(['Current Translate Input', 'Current Input', /current.*input/i]);
                const curOutCol = findCol(['Current Translate Output', 'Current Output', /current.*output/i]);
                const sugCol = findCol(['Suggested Key', /suggested\s*key/i]);
                if (ridCol < 0 || decCol < 0) {
                    resolve(null);
                    return;
                }
                const priorDecisionsObj = {};
                for (let r = 1; r < rows.length; r += 1) {
                    const row = rows[r] || [];
                    const reviewRowId = String(row[ridCol] ?? '').trim();
                    if (!reviewRowId) continue;
                    const decision = String(row[decCol] ?? '').trim();
                    const manualKey = manCol >= 0 ? String(row[manCol] ?? '').trim() : '';
                    const reasonCode = reasonCol >= 0 ? String(row[reasonCol] ?? '').trim() : '';
                    const currentInput = curInCol >= 0 ? String(row[curInCol] ?? '').trim() : '';
                    const currentOutput = curOutCol >= 0 ? String(row[curOutCol] ?? '').trim() : '';
                    const suggestedKey = sugCol >= 0 ? String(row[sugCol] ?? '').trim() : '';
                    priorDecisionsObj[reviewRowId] = {
                        Decision: decision,
                        Manual_Suggested_Key: manualKey,
                        Reason_Code: reasonCode,
                        Current_Input: currentInput,
                        Current_Output: currentOutput,
                        Suggested_Key: suggestedKey
                    };
                }
                resolve({ priorDecisions: priorDecisionsObj, rowCount: Object.keys(priorDecisionsObj).length });
            } catch (err) {
                reject(new Error(`Error parsing prior workbook: ${err.message}`));
            }
        };
        reader.onerror = () => reject(new Error('Error reading prior workbook file'));
        reader.readAsArrayBuffer(file);
    });
}

function parseSessionRowsPayload(payload) {
    if (!payload || !Array.isArray(payload.rows)) {
        throw new Error('Invalid session file. Expected JSON with a rows array.');
    }
    return payload.rows;
}

function applyPriorDecisionsToActionQueue(priorDecisionMap) {
    const rows = getCurrentActionQueueRows();
    if (!rows.length) return { applied: 0, unmatched: 0 };
    const decisions = priorDecisionMap && typeof priorDecisionMap === 'object'
        ? priorDecisionMap
        : {};
    const rowMap = new Map(rows.map(row => [String(row.Review_Row_ID || ''), row]));
    let applied = 0;
    Object.keys(decisions).forEach((rid) => {
        const row = rowMap.get(String(rid || ''));
        if (!row) return;
        const prior = decisions[rid] || {};
        const priorDecision = String(prior.Decision || '').trim();
        const priorManual = String(prior.Manual_Suggested_Key || '').trim();
        const priorSuggested = String(prior.Suggested_Key || '').trim();
        const effectiveKey = priorManual || (priorDecision === 'Use Suggestion' ? priorSuggested : '');
        if (priorDecision) row.Decision = priorDecision;
        if (effectiveKey) {
            row.Manual_Suggested_Key = effectiveKey;
            row.Selected_Candidate_ID = '';
        }
        if (prior.Reason_Code !== undefined && prior.Reason_Code !== null) {
            row.Reason_Code = String(prior.Reason_Code).trim();
        }
        applied += 1;
    });
    preEditedActionQueueRows = rows;
    return {
        applied,
        unmatched: Math.max(0, Object.keys(decisions).length - applied)
    };
}

function setupSessionUploadCard() {
    const sessionInput = document.getElementById('session-upload-file');
    const statusDiv = document.getElementById('session-upload-status');
    const filenameSpan = document.getElementById('session-upload-filename');
    const rowsSpan = document.getElementById('session-upload-rows');
    if (!sessionInput) return;

    sessionInput.addEventListener('change', async function(e) {
        const file = e.target.files?.[0];
        if (!file) {
            uploadedSessionRows = null;
            uploadedSessionApplied = false;
            if (statusDiv) statusDiv.classList.add('hidden');
            return;
        }

        try {
            const text = await file.text();
            const payload = JSON.parse(text);
            const rows = parseSessionRowsPayload(payload);
            uploadedSessionRows = rows;
            uploadedSessionApplied = false;
            if (filenameSpan) filenameSpan.textContent = file.name;
            if (rowsSpan) rowsSpan.textContent = `${rows.length} session row(s) ready to apply`;
            if (statusDiv) statusDiv.classList.remove('hidden');
            if (actionQueueRowsCache && actionQueueRowsCache.length) {
                uploadedSessionApplied = applySessionDataToActionQueue(rows, { sourceLabel: 'Upload' });
                document.dispatchEvent(new CustomEvent('session-upload-applied'));
                refreshErrorPresentation();
            }
        } catch (err) {
            uploadedSessionRows = null;
            uploadedSessionApplied = false;
            if (statusDiv) statusDiv.classList.add('hidden');
            sessionInput.value = '';
            alert(`Error parsing session JSON: ${err.message}`);
        }
    });
}

function setupFileUploads() {
    const fileInputs = [
        { id: 'outcomes-file', key: 'outcomes' },
        { id: 'translate-file', key: 'translate' },
        { id: 'wsu-org-file', key: 'wsu_org' }
    ];

    fileInputs.forEach(input => {
        const element = document.getElementById(input.id);
        element.addEventListener('change', async function(e) {
            await handleFileSelect(e, input.key);
        });

        const encodingSelect = document.getElementById(encodingSelectIds[input.key]);
        if (encodingSelect) {
            encodingSelect.addEventListener('change', async function() {
                if (fileObjects[input.key]) {
                    await reparseFile(input.key);
                }
            });
        }
    });

    const priorFileInput = document.getElementById('prior-validate-file');
    if (priorFileInput) {
        priorFileInput.addEventListener('change', async function(e) {
            const file = e.target.files[0];
            const statusDiv = document.getElementById('prior-validate-status');
            const filenameSpan = document.getElementById('prior-validate-filename');
            const rowsSpan = document.getElementById('prior-validate-rows');
            if (!file) {
                priorDecisions = null;
                if (statusDiv) statusDiv.classList.add('hidden');
                return;
            }
            try {
                const parsed = await parsePriorValidateWorkbook(file);
                if (parsed && parsed.priorDecisions && Object.keys(parsed.priorDecisions).length > 0) {
                    priorDecisions = parsed.priorDecisions;
                    if (filenameSpan) filenameSpan.textContent = file.name;
                    if (rowsSpan) rowsSpan.textContent = `${parsed.rowCount} decisions ready to apply`;
                    if (statusDiv) statusDiv.classList.remove('hidden');
                    if (actionQueueRowsCache && actionQueueRowsCache.length) {
                        const summary = applyPriorDecisionsToActionQueue(priorDecisions);
                        if (rowsSpan) {
                            rowsSpan.textContent = `${parsed.rowCount} decisions loaded; ${summary.applied} applied to current review queue${summary.unmatched ? `, ${summary.unmatched} unmatched` : ''}`;
                        }
                        document.dispatchEvent(new CustomEvent('prior-upload-applied'));
                        refreshErrorPresentation();
                    }
                } else if (parsed && parsed.priorDecisions) {
                    priorDecisions = parsed.priorDecisions;
                    if (filenameSpan) filenameSpan.textContent = file.name;
                    if (rowsSpan) rowsSpan.textContent = 'No decisions in prior workbook';
                    if (statusDiv) statusDiv.classList.remove('hidden');
                } else {
                    priorDecisions = null;
                    if (statusDiv) statusDiv.classList.add('hidden');
                    if (priorFileInput) priorFileInput.value = '';
                    alert('Prior workbook has no Review_Workbench sheet or could not be parsed.');
                }
            } catch (err) {
                priorDecisions = null;
                if (statusDiv) statusDiv.classList.add('hidden');
                if (priorFileInput) priorFileInput.value = '';
                alert(`Error parsing prior workbook: ${err.message}`);
            }
        });
    }
}

async function loadCampusFamilyBaseline() {
    try {
        const res = await fetch('campus-family-baseline.json');
        if (res.ok) {
            const data = await res.json();
            if (data && Array.isArray(data.patterns) && data.patterns.length > 0) {
                campusFamilyRules = data;
            }
        }
    } catch (_) {
        campusFamilyRules = null;
    }
}

function downloadTextFile(filename, content, mimeType = 'text/plain') {
    const blob = new Blob([content], { type: mimeType });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

function parseCampusFamilyEnabled(value) {
    if (typeof value === 'boolean') return value;
    const normalized = String(value ?? '').trim().toLowerCase();
    if (!normalized) return true;
    if (['true', 'yes', '1', 'y'].includes(normalized)) return true;
    if (['false', 'no', '0', 'n'].includes(normalized)) return false;
    return true;
}

function normalizeCampusFamilyRule(raw, index) {
    const pattern = String(
        raw.pattern ?? raw.Pattern ?? raw.name ?? raw.match ?? ''
    ).trim();
    const parentKey = String(
        raw.parentKey ?? raw.ParentKey ?? raw.parent_key ?? raw.key ?? ''
    ).trim();
    if (!pattern || !parentKey) return null;
    const country = String(raw.country ?? raw.Country ?? '').trim();
    const state = String(raw.state ?? raw.State ?? '').trim();
    const priorityRaw = raw.priority ?? raw.Priority ?? (index + 1);
    const priorityNum = Number.parseInt(priorityRaw, 10);
    return {
        pattern,
        parentKey,
        country,
        state,
        priority: Number.isFinite(priorityNum) ? priorityNum : (index + 1),
        enabled: parseCampusFamilyEnabled(raw.enabled ?? raw.Enabled)
    };
}

function parseCampusFamilyDelimitedText(text) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line && !line.startsWith('#'));
    if (!lines.length) return [];

    const firstLine = lines[0];
    const hasHeader = /pattern/i.test(firstLine) && /parent/i.test(firstLine);
    const delimiter = firstLine.includes('\t')
        ? '\t'
        : firstLine.includes('|')
            ? '|'
            : ',';
    const dataLines = hasHeader ? lines.slice(1) : lines;

    return dataLines.map(line => {
        if (line.includes('=>')) {
            const [pattern, parentKey] = line.split('=>').map(v => String(v || '').trim());
            return { pattern, parentKey };
        }
        const parts = line.split(delimiter).map(v => String(v || '').trim());
        return {
            pattern: parts[0] || '',
            parentKey: parts[1] || '',
            country: parts[2] || '',
            state: parts[3] || '',
            priority: parts[4] || '',
            enabled: parts[5] || ''
        };
    });
}

async function parseCampusFamilyRulesFile(file) {
    const fileName = String(file?.name || '').toLowerCase();
    if (!fileName) throw new Error('Missing campus-family file name.');

    let rows = [];
    if (fileName.endsWith('.json')) {
        const text = await file.text();
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed?.patterns)) {
            rows = parsed.patterns;
        } else if (Array.isArray(parsed)) {
            rows = parsed;
        } else {
            throw new Error('JSON must contain a "patterns" array or an array of rules.');
        }
    } else if (fileName.endsWith('.txt')) {
        rows = parseCampusFamilyDelimitedText(await file.text());
    } else if (fileName.endsWith('.csv')) {
        try {
            rows = await loadFile(file, { expectedHeaders: ['pattern', 'parentKey'] });
        } catch (_) {
            rows = parseCampusFamilyDelimitedText(await file.text());
        }
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        rows = await loadFile(file, { expectedHeaders: ['pattern', 'parentKey'] });
    } else {
        throw new Error('Unsupported campus-family format. Use JSON, text, CSV, or Excel.');
    }

    const normalized = rows
        .map((row, idx) => normalizeCampusFamilyRule(row, idx))
        .filter(Boolean);
    if (!normalized.length) {
        throw new Error('No valid rules found. Each row needs pattern and parentKey.');
    }
    return { version: 1, patterns: normalized };
}

function setupCampusFamilyTemplateDownloads() {
    const jsonBtn = document.getElementById('campus-family-template-json-btn');
    const csvBtn = document.getElementById('campus-family-template-csv-btn');
    const defaultRules = {
        version: 1,
        patterns: [
            {
                pattern: 'Texas A&M*',
                parentKey: 'TAMU-MAIN',
                country: '',
                state: '',
                priority: 1,
                enabled: true
            }
        ]
    };
    if (jsonBtn) {
        jsonBtn.addEventListener('click', function() {
            downloadTextFile(
                'campus-family-template.json',
                `${JSON.stringify(defaultRules, null, 2)}\n`,
                'application/json'
            );
        });
    }
    if (csvBtn) {
        csvBtn.addEventListener('click', function() {
            const csv = [
                'pattern,parentKey,country,state,priority,enabled',
                'Texas A&M*,TAMU-MAIN,,,1,true'
            ].join('\n');
            downloadTextFile('campus-family-template.csv', `${csv}\n`, 'text/csv');
        });
    }
}

function setupCampusFamilyUpload() {
    const fileInput = document.getElementById('campus-family-file');
    const statusDiv = document.getElementById('campus-family-status');
    const filenameSpan = document.getElementById('campus-family-filename');
    const rowsSpan = document.getElementById('campus-family-rows');
    setupCampusFamilyTemplateDownloads();
    if (!fileInput) return;
    fileInput.addEventListener('change', async function(e) {
        const file = e.target.files[0];
        if (!file) {
            campusFamilyRules = null;
            if (statusDiv) statusDiv.classList.add('hidden');
            await loadCampusFamilyBaseline();
            return;
        }
        try {
            const parsed = await parseCampusFamilyRulesFile(file);
            campusFamilyRules = parsed;
            if (filenameSpan) filenameSpan.textContent = file.name;
            if (rowsSpan) rowsSpan.textContent = `${parsed.patterns.length} pattern(s)`;
            if (statusDiv) statusDiv.classList.remove('hidden');
        } catch (err) {
            campusFamilyRules = null;
            if (statusDiv) statusDiv.classList.add('hidden');
            fileInput.value = '';
            alert(`Error parsing campus-family rules: ${err.message}`);
        }
    });
}

function setupDebugToggle() {
    const toggle = document.getElementById('debug-toggle');
    const previewToggle = document.getElementById('debug-preview-toggle');
    if (!toggle) return;
    toggle.addEventListener('change', renderDebugPanel);
    if (previewToggle) {
        previewToggle.addEventListener('change', renderDebugPanel);
    }
}

function renderDebugPanel() {
    const toggle = document.getElementById('debug-toggle');
    const panel = document.getElementById('debug-panel');
    const previewToggle = document.getElementById('debug-preview-toggle');
    if (!toggle || !panel) return;
    panel.classList.toggle('hidden', !toggle.checked);
    if (!toggle.checked) return;
    const showPreview = previewToggle ? previewToggle.checked : true;

    const format = (entry) => {
        if (!entry) return 'No data loaded.';
        const lines = [
            `File: ${entry.filename || 'Unknown'}`,
            `Rows: ${entry.rows}`,
            `Columns: ${entry.columns.join(', ') || 'None'}`
        ];
        if (showPreview) {
            lines.push('Preview:', JSON.stringify(entry.preview, null, 2));
        }
        return lines.join('\n');
    };

    const outcomesEl = document.getElementById('debug-outcomes');
    const translateEl = document.getElementById('debug-translate');
    const wsuEl = document.getElementById('debug-wsu');
    if (outcomesEl) outcomesEl.textContent = format(debugState.outcomes);
    if (translateEl) translateEl.textContent = format(debugState.translate);
    if (wsuEl) wsuEl.textContent = format(debugState.wsu_org);
}

function updateDebugState(fileKey, file, data) {
    const columns = Object.keys(data[0] || {}).filter(col => !col.startsWith('Unnamed'));
    const preview = data.slice(0, 5);
    debugState[fileKey] = {
        filename: file?.name || '',
        rows: data.length,
        columns,
        preview
    };
    renderDebugPanel();
}

async function handleFileSelect(event, fileKey) {
    const file = event.target.files[0];
    if (!file) return;

    const statusDiv = document.getElementById(`${fileKey.replace('_', '-')}-status`);
    const filenameSpan = document.getElementById(`${fileKey.replace('_', '-')}-filename`);
    const rowsSpan = document.getElementById(`${fileKey.replace('_', '-')}-rows`);

    try {
        filenameSpan.textContent = file.name;
        statusDiv.classList.remove('hidden');

        const data = await loadFile(file, {
            encoding: getFileEncoding(fileKey),
            sheetHeaderHints: getSheetHeaderHints(fileKey)
        });

        fileObjects[fileKey] = file;
        loadedData[fileKey] = data;
        filesUploaded[fileKey] = true;
        updateDebugState(fileKey, file, data);

        rowsSpan.textContent = `${data.length} rows`;

        processAvailableFiles();

    } catch (error) {
        console.error(`Error loading ${fileKey}:`, error);
        alert(`Error loading file: ${error.message}`);
        statusDiv.classList.add('hidden');
        filesUploaded[fileKey] = false;
    }
}

async function reparseFile(fileKey) {
    const file = fileObjects[fileKey];
    if (!file) return;

    const statusDiv = document.getElementById(`${fileKey.replace('_', '-')}-status`);
    const rowsSpan = document.getElementById(`${fileKey.replace('_', '-')}-rows`);

    try {
        const data = await loadFile(file, {
            encoding: getFileEncoding(fileKey),
            sheetHeaderHints: getSheetHeaderHints(fileKey)
        });
        loadedData[fileKey] = data;
        filesUploaded[fileKey] = true;
        updateDebugState(fileKey, file, data);
        rowsSpan.textContent = `${data.length} rows`;
        processAvailableFiles();
    } catch (error) {
        console.error(`Error re-parsing ${fileKey}:`, error);
        alert(`Error re-parsing file with selected encoding: ${error.message}`);
        statusDiv.classList.add('hidden');
        filesUploaded[fileKey] = false;
    }
}

function processAvailableFiles() {
    try {
        const outcomesReady = filesUploaded.outcomes;
        const translateReady = filesUploaded.translate;
        const wsuReady = filesUploaded.wsu_org;

        if (!outcomesReady && !translateReady && !wsuReady) {
            return;
        }

        const filterSourceColumns = (columns) => {
            const excluded = new Set([
                'Error_Type',
                'Error_Description',
                'Duplicate_Group',
                'translate_input',
                'translate_output',
                'translate_input_norm',
                'translate_output_norm',
                'normalized_key',
                'missing_in',
                'match_similarity'
            ]);
            return columns.filter(col => !excluded.has(col));
        };

        const outcomesColumns = outcomesReady
            ? filterSourceColumns(
                Object.keys(loadedData.outcomes[0] || {}).filter(col => !col.startsWith('Unnamed'))
            )
            : [];
        const translateColumns = translateReady
            ? Object.keys(loadedData.translate[0] || {}).filter(col => !col.startsWith('Unnamed'))
            : [];
        const wsuOrgColumns = wsuReady
            ? filterSourceColumns(
                Object.keys(loadedData.wsu_org[0] || {}).filter(col => !col.startsWith('Unnamed'))
            )
            : [];

        if (outcomesReady || translateReady || wsuReady) {
            populateKeySelection(outcomesColumns, translateColumns, wsuOrgColumns);
        }
        if (outcomesReady || wsuReady) {
            populateColumnSelection(outcomesColumns, wsuOrgColumns);
            populateNameCompareOptions(outcomesColumns, wsuOrgColumns);
            document.getElementById('column-selection').classList.remove('hidden');
            applyCreateDefaults(outcomesColumns, wsuOrgColumns);
        } else if (currentMode === 'create') {
            document.getElementById('column-selection').classList.remove('hidden');
        }

        const validateBtn = document.getElementById('validate-btn');
        const validateMessage = document.getElementById('validation-message');
        if (currentMode === 'validate' && outcomesReady && translateReady && wsuReady) {
            validateBtn.disabled = false;
            validateBtn.classList.remove('bg-gray-400', 'cursor-not-allowed');
            validateBtn.classList.add('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
            validateMessage.textContent = 'Ready to validate!';
        } else {
            validateBtn.disabled = true;
            validateBtn.classList.add('bg-gray-400', 'cursor-not-allowed');
            validateBtn.classList.remove('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
            validateMessage.textContent = currentMode === 'validate'
                ? 'Upload Outcomes, Translation, and myWSU to validate.'
                : 'Switch to Validate mode to run validation.';
        }

        const generateBtn = document.getElementById('generate-btn');
        const generateMessage = document.getElementById('generate-message');
        if (currentMode === 'create' && outcomesReady && wsuReady) {
            generateBtn.disabled = false;
            generateBtn.classList.remove('bg-gray-400', 'cursor-not-allowed');
            generateBtn.classList.add('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
            generateMessage.textContent = matchMethod === 'name'
                ? 'Name matching mode: key selections are optional and ignored.'
                : 'Choose match method and columns, then generate.';
        } else {
            generateBtn.disabled = true;
            generateBtn.classList.add('bg-gray-400', 'cursor-not-allowed');
            generateBtn.classList.remove('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
            generateMessage.textContent = currentMode === 'create'
                ? 'Upload Outcomes + myWSU to populate match options.'
                : 'Switch to Create mode to generate a translation table.';
        }

        const joinPreviewBtn = document.getElementById('join-preview-btn');
        const joinPreviewMessage = document.getElementById('join-preview-message');
        if (joinPreviewBtn && joinPreviewMessage) {
            if (currentMode === 'join-preview' && outcomesReady && translateReady && wsuReady) {
                joinPreviewBtn.disabled = false;
                joinPreviewBtn.classList.remove('bg-gray-400', 'cursor-not-allowed');
                joinPreviewBtn.classList.add('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
                joinPreviewMessage.textContent = 'Ready to generate join preview.';
            } else {
                joinPreviewBtn.disabled = true;
                joinPreviewBtn.classList.add('bg-gray-400', 'cursor-not-allowed');
                joinPreviewBtn.classList.remove('bg-wsu-crimson', 'hover:bg-red-800', 'cursor-pointer');
                joinPreviewMessage.textContent = currentMode === 'join-preview'
                    ? 'Upload Outcomes, Translation table, and myWSU to generate join preview.'
                    : 'Switch to Join Preview mode.';
            }
        }

        if (pageBusy || runLocked) {
            setPrimaryActionsBusy(true);
            if (validateMessage && currentMode === 'validate') {
                validateMessage.textContent = 'Validation is running. Please wait...';
            }
            if (generateMessage && currentMode === 'create') {
                generateMessage.textContent = 'Generation is running. Please wait...';
            }
            if (joinPreviewMessage && currentMode === 'join-preview') {
                joinPreviewMessage.textContent = 'Join preview is running. Please wait...';
            }
        }

        if (currentMode === 'create') {
            syncCreateKeyControlsForMatchMethod();
        }

    } catch (error) {
        console.error('Error processing files:', error);
        alert('Error processing files. Please try again.');
    }
}

function applyCreateDefaults(outcomesColumns, wsuOrgColumns) {
    if (currentMode !== 'create') {
        return;
    }
    if (matchMethodTouched) {
        return;
    }
    const keyRadio = document.getElementById('match-method-key');
    const nameRadio = document.getElementById('match-method-name');
    if (!keyRadio || !nameRadio) {
        return;
    }
    const hasKeyDefaults = Boolean(keyConfig.outcomes && keyConfig.wsu);
    const hasNameDefaults = Boolean(
        outcomesColumns.includes('name') &&
        wsuOrgColumns.includes('Descr')
    );
    if (!hasKeyDefaults && hasNameDefaults) {
        nameRadio.checked = true;
        keyRadio.checked = false;
        matchMethod = 'name';
        updateNameCompareState();
        updateMatchMethodUI();
    }
}

function findColumn(columns, candidates) {
    const lowerMap = new Map(columns.map(col => [col.toLowerCase(), col]));
    for (const candidate of candidates) {
        const match = lowerMap.get(candidate.toLowerCase());
        if (match) return match;
    }
    return '';
}

function populateColumnSelection(outcomesColumns, wsuOrgColumns) {
    const defaultOutcomes = ['name', 'mdb_code', 'state', 'country'];
    const defaultWsuOrg = ['Org ID', 'Descr', 'City', 'State', 'Country'];
    const roleOptions = [
        { value: '', label: 'None' },
        { value: 'School', label: 'School' },
        { value: 'City', label: 'City' },
        { value: 'State', label: 'State' },
        { value: 'Country', label: 'Country' },
        { value: 'Other', label: 'Other' }
    ];

    const guessRole = (col) => {
        const normalized = String(col || '').toLowerCase();
        // Do not auto-assign school type columns to Other; leave for explicit reviewer choice.
        if (normalized.includes('school type')) return '';
        if (normalized.includes('name')) return 'School';
        if (normalized.includes('city')) return 'City';
        if (normalized.includes('state')) return 'State';
        if (normalized.includes('country')) return 'Country';
        return '';
    };

    const defaultOutcomesKey = keyConfig.outcomes || findColumn(outcomesColumns, [
        'mdb_code',
        'mdb code',
        'outcomes_state',
        'outcomes state',
        'state'
    ]) || (outcomesColumns[0] || '');

    const defaultWsuKey = keyConfig.wsu || findColumn(wsuOrgColumns, [
        'org id',
        'state',
        'mywsu_state',
        'mywsu state'
    ]) || (wsuOrgColumns[0] || '');

    const outcomesDiv = document.getElementById('outcomes-columns');
    outcomesDiv.innerHTML = '';
    selectedColumns.outcomes = [];
    columnRoles.outcomes = {};
    if (!outcomesColumns.length) {
        outcomesDiv.innerHTML = '<p class="text-xs text-gray-500">Upload Outcomes to see columns.</p>';
    } else {
        outcomesColumns.forEach(col => {
            const isChecked = defaultOutcomes.includes(col);
            const row = document.createElement('div');
            row.className = 'grid grid-cols-4 gap-2 items-center';

            const nameSpan = document.createElement('span');
            nameSpan.className = 'text-sm text-gray-700 truncate';
            nameSpan.title = col;
            nameSpan.textContent = col;

            const includeInput = document.createElement('input');
            includeInput.type = 'checkbox';
            includeInput.name = 'outcomes-include';
            includeInput.value = col;
            includeInput.checked = isChecked;
            includeInput.className = 'rounded border-gray-300 text-wsu-crimson focus:ring-wsu-crimson';

            const keyInput = document.createElement('input');
            keyInput.type = 'radio';
            keyInput.name = 'outcomes-key';
            keyInput.value = col;
            keyInput.checked = col === defaultOutcomesKey;
            keyInput.className = 'text-wsu-crimson focus:ring-wsu-crimson';

            const roleSelect = document.createElement('select');
            roleSelect.name = 'outcomes-role';
            roleSelect.dataset.col = col;
            roleSelect.className = 'w-full border border-gray-300 rounded-md p-1 text-xs role-col';
            roleOptions.forEach(optionData => {
                const option = document.createElement('option');
                option.value = optionData.value;
                option.textContent = optionData.label;
                roleSelect.appendChild(option);
            });
            const guessedRole = guessRole(col);
            roleSelect.value = guessedRole;
            if (guessedRole) {
                roleSelect.dataset.prevRole = guessedRole;
            }
            columnRoles.outcomes[col] = guessedRole;
            if (col === defaultOutcomesKey) {
                roleSelect.value = '';
                roleSelect.disabled = true;
            }

            row.appendChild(nameSpan);
            row.appendChild(includeInput);
            row.appendChild(keyInput);
            row.appendChild(roleSelect);
            outcomesDiv.appendChild(row);

            if (isChecked) {
                selectedColumns.outcomes.push(col);
            }
        });
    }

    const wsuOrgDiv = document.getElementById('wsu-org-columns');
    wsuOrgDiv.innerHTML = '';
    selectedColumns.wsu_org = [];
    columnRoles.wsu_org = {};
    if (!wsuOrgColumns.length) {
        wsuOrgDiv.innerHTML = '<p class="text-xs text-gray-500">Upload myWSU to see columns.</p>';
    } else {
        wsuOrgColumns.forEach(col => {
            const isChecked = defaultWsuOrg.includes(col);
            const row = document.createElement('div');
            row.className = 'grid grid-cols-4 gap-2 items-center';

            const nameSpan = document.createElement('span');
            nameSpan.className = 'text-sm text-gray-700 truncate';
            nameSpan.title = col;
            nameSpan.textContent = col;

            const includeInput = document.createElement('input');
            includeInput.type = 'checkbox';
            includeInput.name = 'wsu-include';
            includeInput.value = col;
            includeInput.checked = isChecked;
            includeInput.className = 'rounded border-gray-300 text-wsu-crimson focus:ring-wsu-crimson';

            const keyInput = document.createElement('input');
            keyInput.type = 'radio';
            keyInput.name = 'wsu-key';
            keyInput.value = col;
            keyInput.checked = col === defaultWsuKey;
            keyInput.className = 'text-wsu-crimson focus:ring-wsu-crimson';

            const roleSelect = document.createElement('select');
            roleSelect.name = 'wsu-role';
            roleSelect.dataset.col = col;
            roleSelect.className = 'w-full border border-gray-300 rounded-md p-1 text-xs role-col';
            roleOptions.forEach(optionData => {
                const option = document.createElement('option');
                option.value = optionData.value;
                option.textContent = optionData.label;
                roleSelect.appendChild(option);
            });
            const guessedRole = guessRole(col);
            roleSelect.value = guessedRole;
            if (guessedRole) {
                roleSelect.dataset.prevRole = guessedRole;
            }
            columnRoles.wsu_org[col] = guessedRole;
            if (col === defaultWsuKey) {
                roleSelect.value = '';
                roleSelect.disabled = true;
            }

            row.appendChild(nameSpan);
            row.appendChild(includeInput);
            row.appendChild(keyInput);
            row.appendChild(roleSelect);
            wsuOrgDiv.appendChild(row);

            if (isChecked) {
                selectedColumns.wsu_org.push(col);
            }
        });
    }

    document.querySelectorAll('input[name="outcomes-include"], input[name="wsu-include"], input[name="outcomes-key"], input[name="wsu-key"]').forEach(input => {
        input.addEventListener('change', updateSelectedColumns);
    });
    document.querySelectorAll('select[name="outcomes-role"], select[name="wsu-role"]').forEach(select => {
        select.addEventListener('change', updateSelectedColumns);
    });

    updateSelectedColumns();
}

function populateKeySelection(outcomesColumns, translateColumns, wsuOrgColumns) {
    const translateInputSelect = document.getElementById('key-translate-input');
    const translateOutputSelect = document.getElementById('key-translate-output');

    if (!translateInputSelect || !translateOutputSelect) {
        return;
    }

    const previousTranslateInput = keyConfig.translateInput || translateInputSelect.value || '';
    const previousTranslateOutput = keyConfig.translateOutput || translateOutputSelect.value || '';

    translateInputSelect.innerHTML = '<option value="">Select column</option>';
    translateOutputSelect.innerHTML = '<option value="">Select column</option>';
    if (translateColumns.length) {
        translateColumns.forEach(col => {
            const inputOption = document.createElement('option');
            inputOption.value = col;
            inputOption.textContent = col;
            translateInputSelect.appendChild(inputOption);
            const outputOption = document.createElement('option');
            outputOption.value = col;
            outputOption.textContent = col;
            translateOutputSelect.appendChild(outputOption);
        });
        translateInputSelect.disabled = false;
        translateOutputSelect.disabled = false;
    } else {
        const inputOption = document.createElement('option');
        inputOption.value = '';
        inputOption.textContent = 'Upload translation table to select';
        translateInputSelect.appendChild(inputOption);
        const outputOption = document.createElement('option');
        outputOption.value = '';
        outputOption.textContent = 'Upload translation table to select';
        translateOutputSelect.appendChild(outputOption);
        translateInputSelect.disabled = true;
        translateOutputSelect.disabled = true;
    }
    const defaultTranslateInput = translateColumns.length
        ? (findColumn(translateColumns, [
            'input',
            'mdb_code',
            'outcomes_state',
            'outcomes state',
            'state'
        ]) || (translateColumns[0] || ''))
        : '';

    const defaultTranslateOutput = translateColumns.length
        ? (findColumn(translateColumns, [
            'output',
            'org id',
            'mywsu_state',
            'mywsu state',
            'state'
        ]) || (translateColumns[1] || translateColumns[0] || ''))
        : '';

    const resolvedTranslateInput = translateColumns.includes(previousTranslateInput)
        ? previousTranslateInput
        : defaultTranslateInput;
    const resolvedTranslateOutput = translateColumns.includes(previousTranslateOutput)
        ? previousTranslateOutput
        : defaultTranslateOutput;

    translateInputSelect.value = resolvedTranslateInput;
    translateOutputSelect.value = resolvedTranslateOutput;

    keyConfig = {
        outcomes: keyConfig.outcomes || '',
        translateInput: resolvedTranslateInput,
        translateOutput: resolvedTranslateOutput,
        wsu: keyConfig.wsu || ''
    };

    keyLabels = {
        outcomes: keyConfig.outcomes,
        translateInput: resolvedTranslateInput,
        translateOutput: resolvedTranslateOutput,
        wsu: keyConfig.wsu
    };

    // Avoid duplicate listeners as this function runs each time files are (re)processed.
    translateInputSelect.onchange = updateKeyConfig;
    translateOutputSelect.onchange = updateKeyConfig;
}

function updateKeyConfig() {
    const translateInputSelect = document.getElementById('key-translate-input');
    const translateOutputSelect = document.getElementById('key-translate-output');
    const outcomesKey = document.querySelector('input[name="outcomes-key"]:checked')?.value || '';
    const wsuKey = document.querySelector('input[name="wsu-key"]:checked')?.value || '';

    keyConfig = {
        outcomes: outcomesKey,
        translateInput: translateInputSelect?.value || '',
        translateOutput: translateOutputSelect?.value || '',
        wsu: wsuKey
    };

    keyLabels = {
        outcomes: keyConfig.outcomes,
        translateInput: keyConfig.translateInput,
        translateOutput: keyConfig.translateOutput,
        wsu: keyConfig.wsu
    };
}

function updateSelectedColumns() {
    selectedColumns.outcomes = Array.from(document.querySelectorAll('input[name="outcomes-include"]:checked'))
        .map(cb => cb.value);
    selectedColumns.wsu_org = Array.from(document.querySelectorAll('input[name="wsu-include"]:checked'))
        .map(cb => cb.value);

    columnRoles.outcomes = {};
    columnRoles.wsu_org = {};

    const ignoreKeysForCreateNameMode = currentMode === 'create' && matchMethod === 'name';
    const outcomesKey = ignoreKeysForCreateNameMode
        ? ''
        : (document.querySelector('input[name="outcomes-key"]:checked')?.value || '');
    const wsuKey = ignoreKeysForCreateNameMode
        ? ''
        : (document.querySelector('input[name="wsu-key"]:checked')?.value || '');

    document.querySelectorAll('select[name="outcomes-role"]').forEach(select => {
        const col = select.dataset.col;
        const wasDisabled = select.disabled;
        if (col === outcomesKey) {
            if (select.value && select.value !== '') {
                select.dataset.prevRole = select.value;
            }
            select.value = '';
            select.disabled = true;
        } else {
            select.disabled = false;
            // Restore previous role only when coming back from key-lock state.
            if (wasDisabled && !select.value && select.dataset.prevRole) {
                select.value = select.dataset.prevRole;
            }
        }
        if (col) {
            columnRoles.outcomes[col] = select.value || '';
        }
    });
    document.querySelectorAll('select[name="wsu-role"]').forEach(select => {
        const col = select.dataset.col;
        const wasDisabled = select.disabled;
        if (col === wsuKey) {
            if (select.value && select.value !== '') {
                select.dataset.prevRole = select.value;
            }
            select.value = '';
            select.disabled = true;
        } else {
            select.disabled = false;
            // Restore previous role only when coming back from key-lock state.
            if (wasDisabled && !select.value && select.dataset.prevRole) {
                select.value = select.dataset.prevRole;
            }
        }
        if (col) {
            columnRoles.wsu_org[col] = select.value || '';
        }
    });

    updateKeyConfig();
}

function setupNameCompareControls() {
    const fields = document.getElementById('name-compare-fields');
    if (!fields) return;
    updateNameCompareState();
}

function syncCreateKeyControlsForMatchMethod() {
    if (currentMode !== 'create') {
        return;
    }
    const useNameOnly = matchMethod === 'name';
    const keyHelp = document.getElementById('create-key-help');
    if (keyHelp) {
        keyHelp.classList.toggle('hidden', !useNameOnly);
    }

    const toggleGroup = (groupName) => {
        const inputs = Array.from(document.querySelectorAll(`input[name="${groupName}"]`));
        if (!inputs.length) return;
        if (useNameOnly) {
            inputs.forEach(input => {
                input.checked = false;
                input.disabled = true;
            });
            return;
        }
        inputs.forEach(input => {
            input.disabled = false;
        });
        const hasChecked = inputs.some(input => input.checked);
        if (!hasChecked) {
            inputs[0].checked = true;
        }
    };

    toggleGroup('outcomes-key');
    toggleGroup('wsu-key');
}

function updateNameCompareState() {
    const fields = document.getElementById('name-compare-fields');
    if (!fields) return;

    const controls = fields.querySelectorAll('select, input');
    controls.forEach(control => {
        control.disabled = false;
    });
    fields.classList.remove('opacity-50');
}

function populateNameCompareOptions(outcomesColumns, wsuOrgColumns) {
    const outcomesSelect = document.getElementById('name-compare-outcomes');
    const wsuSelect = document.getElementById('name-compare-wsu');
    const thresholdInput = document.getElementById('name-compare-threshold');
    const ambiguityInput = document.getElementById('name-compare-ambiguity-gap');

    if (!outcomesSelect || !wsuSelect || !thresholdInput || !ambiguityInput) return;

    const previousOutcomes = outcomesSelect.value;
    const previousWsu = wsuSelect.value;
    const previousThreshold = thresholdInput.value;
    const previousGap = ambiguityInput.value;

    outcomesSelect.innerHTML = '<option value="">Select column</option>';
    wsuSelect.innerHTML = '<option value="">Select column</option>';

    if (outcomesColumns.length) {
        outcomesColumns.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            outcomesSelect.appendChild(option);
        });
        outcomesSelect.disabled = false;
    } else {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = 'Upload Outcomes to select';
        outcomesSelect.appendChild(option);
        outcomesSelect.disabled = true;
    }
    if (wsuOrgColumns.length) {
        wsuOrgColumns.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            wsuSelect.appendChild(option);
        });
        wsuSelect.disabled = false;
    } else {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = 'Upload myWSU to select';
        wsuSelect.appendChild(option);
        wsuSelect.disabled = true;
    }

    const defaultOutcomes = outcomesColumns.includes('name') ? 'name' : '';
    const defaultWsu = wsuOrgColumns.includes('Descr') ? 'Descr' : '';
    const resolvedOutcomes = outcomesColumns.includes(previousOutcomes)
        ? previousOutcomes
        : defaultOutcomes;
    const resolvedWsu = wsuOrgColumns.includes(previousWsu)
        ? previousWsu
        : defaultWsu;
    if (resolvedOutcomes) outcomesSelect.value = resolvedOutcomes;
    if (resolvedWsu) wsuSelect.value = resolvedWsu;

    thresholdInput.value = previousThreshold || '0.8';
    ambiguityInput.value = previousGap || '0.03';
    updateNameCompareState();
}

function setupColumnSelection() {
    const toggleBtn = document.getElementById('toggle-columns');
    const checkboxesDiv = document.getElementById('column-checkboxes');
    if (!toggleBtn || !checkboxesDiv) return;

    checkboxesDiv.classList.remove('hidden');
    const svg = toggleBtn.querySelector('svg');
    if (svg) {
        svg.classList.add('rotate-180');
    }

    toggleBtn.addEventListener('click', function() {
        checkboxesDiv.classList.toggle('hidden');
        const svg = toggleBtn.querySelector('svg');
        if (svg) {
            svg.classList.toggle('rotate-180');
        }
    });
}

function setupShowAllErrorsToggle() {
    const showAllCheckbox = document.getElementById('show-all-errors');
    const unresolvedCheckbox = document.getElementById('show-unresolved-errors-only');
    if (showAllCheckbox) {
        showAllCheckbox.addEventListener('change', function() {
            showAllErrors = showAllCheckbox.checked;
            if (validatedData.length > 0) {
                refreshErrorPresentation();
                renderMappingLogicPreview();
            }
        });
    }
    if (unresolvedCheckbox) {
        unresolvedCheckbox.addEventListener('change', function() {
            showUnresolvedErrorsOnly = unresolvedCheckbox.checked;
            if (validatedData.length > 0) {
                refreshErrorPresentation();
            }
        });
    }
    updateUnresolvedErrorsToggleState();
}

function setupMappingLogicPreviewToggle() {
    const checkbox = document.getElementById('show-logic-preview');
    if (!checkbox) return;
    checkbox.addEventListener('change', function() {
        renderMappingLogicPreview();
    });
}

function setupMatchingRulesToggle() {
    const checkbox = document.getElementById('show-matching-rules');
    if (!checkbox) return;
    checkbox.addEventListener('change', function() {
        renderMatchingRulesExamples();
    });
}

function setupValidateButton() {
    const validateBtn = document.getElementById('validate-btn');
    validateBtn.addEventListener('click', runValidation);
}

function setupGenerateButton() {
    const generateBtn = document.getElementById('generate-btn');
    if (!generateBtn) return;
    generateBtn.addEventListener('click', runGeneration);
}

function setupJoinPreviewButton() {
    const joinPreviewBtn = document.getElementById('join-preview-btn');
    if (!joinPreviewBtn) return;
    joinPreviewBtn.addEventListener('click', runJoinPreview);
}

async function runJoinPreview() {
    if (runLocked) return;
    if (currentMode !== 'join-preview') {
        alert('Switch to Join Preview mode to run.');
        return;
    }
    updateKeyConfig();
    if (!keyConfig.outcomes || !keyConfig.translateInput || !keyConfig.translateOutput || !keyConfig.wsu) {
        alert('Select Outcomes key, Translation input/output columns, and myWSU key before generating.');
        return;
    }
    setRunLock(true);
    try {
        document.getElementById('loading').classList.remove('hidden');
        document.getElementById('results').classList.add('hidden');
        document.getElementById('loading-message').textContent = 'Building join preview...';
        const progressWrap = document.getElementById('loading-progress');
        if (progressWrap) progressWrap.classList.add('hidden');

        const result = await runExportWorkerTask(
            'build_join_preview_export',
            {
                selectedCols: {
                    outcomes: selectedColumns.outcomes || [],
                    wsu_org: selectedColumns.wsu_org || []
                },
                options: { fileName: '' },
                context: {
                    loadedData,
                    columnRoles,
                    keyConfig,
                    keyLabels
                }
            }
        );
        downloadArrayBuffer(result.buffer, result.filename || 'WSU_Join_Preview.xlsx');
        document.getElementById('loading').classList.add('hidden');
    } catch (error) {
        console.error('Join preview error:', error);
        alert('Join preview failed: ' + (error?.message || String(error)));
        document.getElementById('loading').classList.add('hidden');
    } finally {
        setRunLock(false);
    }
}

function setupMatchMethodControls() {
    const keyRadio = document.getElementById('match-method-key');
    const nameRadio = document.getElementById('match-method-name');
    if (!keyRadio || !nameRadio) return;

    const syncMethod = () => {
        matchMethodTouched = true;
        matchMethod = nameRadio.checked ? 'name' : 'key';
        updateMatchMethodUI();
    };

    keyRadio.addEventListener('change', syncMethod);
    nameRadio.addEventListener('change', syncMethod);
}

function setupValidateNameModeControls() {
    const keyOnlyRadio = document.getElementById('validate-key-only');
    const keyNameRadio = document.getElementById('validate-key-name');
    if (!keyOnlyRadio || !keyNameRadio) return;

    const syncMode = () => {
        validateNameMode = keyNameRadio.checked ? 'key+name' : 'key';
        updateValidateNameModeUI();
    };

    keyOnlyRadio.addEventListener('change', syncMode);
    keyNameRadio.addEventListener('change', syncMode);
    syncMode();
}

function updateMatchMethodUI() {
    if (currentMode !== 'create') {
        return;
    }
    const keyMatchFields = document.getElementById('key-match-fields');
    const nameCompare = document.getElementById('name-compare');
    if (keyMatchFields) {
        keyMatchFields.classList.toggle('hidden', matchMethod === 'name');
    }
    if (nameCompare) {
        nameCompare.classList.toggle('hidden', matchMethod === 'key');
    }
    syncCreateKeyControlsForMatchMethod();
    updateSelectedColumns();
}

function updateValidateNameModeUI() {
    if (currentMode !== 'validate') {
        return;
    }
    const nameCompare = document.getElementById('name-compare');
    if (nameCompare) {
        nameCompare.classList.toggle('hidden', validateNameMode !== 'key+name');
    }
}

async function runValidation() {
    if (runLocked) {
        return;
    }
    setRunLock(true);
    try {
        if (currentMode !== 'validate') {
            alert('Switch to Validate mode to run validation.');
            return;
        }
        cancelActionQueuePrefetch();
        // Clear pre-export bulk edits so a new validation run cannot reuse stale queue state.
        actionQueueRowsCache = null;
        preEditedActionQueueRows = null;
        uploadedSessionApplied = false;
        updateUnresolvedErrorsToggleState();
        const bulkPanel = document.getElementById('bulk-edit-panel');
        if (bulkPanel) bulkPanel.classList.add('hidden');
        document.getElementById('loading').classList.remove('hidden');
        document.getElementById('results').classList.add('hidden');
        document.getElementById('loading-message').textContent = 'Analyzing mappings...';
        const progressWrap = document.getElementById('loading-progress');
        const progressStage = document.getElementById('loading-stage');
        const progressPercent = document.getElementById('loading-percent');
        const progressBar = document.getElementById('loading-bar');
        if (progressWrap && progressStage && progressPercent && progressBar) {
            progressWrap.classList.remove('hidden');
            progressStage.textContent = 'Preparing...';
            progressPercent.textContent = '0%';
            progressBar.style.width = '0%';
        }

        await new Promise(resolve => setTimeout(resolve, 100));

        updateKeyConfig();
        if (!keyConfig.outcomes || !keyConfig.translateInput || !keyConfig.translateOutput || !keyConfig.wsu) {
            alert('Select all key columns before validating.');
            document.getElementById('loading').classList.add('hidden');
            return;
        }

        const nameCompareOutcomes = document.getElementById('name-compare-outcomes')?.value || '';
        const nameCompareWsu = document.getElementById('name-compare-wsu')?.value || '';
        const outcomesStateRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'State'
        ) || '';
        const wsuStateRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'State'
        ) || '';
        const outcomesCityRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'City'
        ) || '';
        const wsuCityRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'City'
        ) || '';
        const outcomesCountryRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'Country'
        ) || '';
        const wsuCountryRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'Country'
        ) || '';
        const findFallbackColumn = (columns, token) => (
            columns.find(col => String(col).toLowerCase().includes(token)) || ''
        );
        const outcomesStateFallback = outcomesStateRole
            || findFallbackColumn(selectedColumns.outcomes, 'state');
        const wsuStateFallback = wsuStateRole
            || findFallbackColumn(selectedColumns.wsu_org, 'state');
        const outcomesCityFallback = outcomesCityRole
            || findFallbackColumn(selectedColumns.outcomes, 'city');
        const wsuCityFallback = wsuCityRole
            || findFallbackColumn(selectedColumns.wsu_org, 'city');
        const outcomesCountryFallback = outcomesCountryRole
            || findFallbackColumn(selectedColumns.outcomes, 'country');
        const wsuCountryFallback = wsuCountryRole
            || findFallbackColumn(selectedColumns.wsu_org, 'country');
        const nameCompareThreshold = parseFloat(
            document.getElementById('name-compare-threshold')?.value || '0.8'
        );
        const nameCompareGap = parseFloat(
            document.getElementById('name-compare-ambiguity-gap')?.value || '0.03'
        );
        const resolvedThreshold = Number.isNaN(nameCompareThreshold)
            ? 0.8
            : Math.max(0, Math.min(1, nameCompareThreshold));
        const resolvedGap = Number.isNaN(nameCompareGap)
            ? 0.03
            : Math.max(0, Math.min(0.2, nameCompareGap));

        const wantsNameCompare = validateNameMode === 'key+name';
        const nameCompareEnabled = wantsNameCompare && Boolean(nameCompareOutcomes && nameCompareWsu);
        if (wantsNameCompare && !nameCompareEnabled) {
            if (nameCompareOutcomes || nameCompareWsu) {
                alert('Select both name columns or disable name comparison.');
            } else {
                alert('Select name columns or switch to key-only validation.');
            }
            document.getElementById('loading').classList.add('hidden');
            return;
        }

        lastNameCompareConfig = {
            enabled: Boolean(nameCompareEnabled),
            outcomes: nameCompareOutcomes,
            wsu: nameCompareWsu,
            threshold: resolvedThreshold,
            ambiguity_gap: resolvedGap,
            state_outcomes: outcomesStateFallback,
            state_wsu: wsuStateFallback,
            city_outcomes: outcomesCityFallback,
            city_wsu: wsuCityFallback,
            country_outcomes: outcomesCountryFallback,
            country_wsu: wsuCountryFallback
        };

        const translateRows = loadedData.translate || [];
        const translateRowCount = translateRows.length;
        if (!translateRowCount) {
            alert('System check failed: Translate table has 0 rows. Re-upload and try again.');
            document.getElementById('loading').classList.add('hidden');
            if (progressWrap) {
                progressWrap.classList.add('hidden');
            }
            return;
        }

        const getValueText = (value) => String(value ?? '').trim();
        let missingInputs = 0;
        let missingOutputs = 0;
        translateRows.forEach(row => {
            if (!getValueText(row[keyConfig.translateInput])) {
                missingInputs += 1;
            }
            if (!getValueText(row[keyConfig.translateOutput])) {
                missingOutputs += 1;
            }
        });
        if (missingInputs > 0 || missingOutputs > 0) {
            alert(
                `System check failed: Translate table has ${missingInputs} blank input key cells and ` +
                `${missingOutputs} blank output key cells. Fix the file or key selection and try again.`
            );
            document.getElementById('loading').classList.add('hidden');
            if (progressWrap) {
                progressWrap.classList.add('hidden');
            }
            return;
        }

        if (progressStage && progressPercent && progressBar) {
            progressStage.textContent = `System check passed: ${translateRowCount.toLocaleString()} rows, no blank key cells`;
            progressPercent.textContent = '5%';
            progressBar.style.width = '5%';
        }
        document.getElementById('loading-message').textContent =
            `System check passed: ${translateRowCount.toLocaleString()} rows, no blank key cells`;

        let lastValidationPercent = 5;
        setPageBusy(true);
        const result = await runWorkerTask(
            'validate',
            {
                outcomes: loadedData.outcomes,
                translate: loadedData.translate,
                wsu_org: loadedData.wsu_org,
                keyConfig,
                nameCompare: {
                    enabled: Boolean(nameCompareEnabled),
                    outcomes_column: nameCompareOutcomes,
                    wsu_column: nameCompareWsu,
                    threshold: resolvedThreshold,
                    ambiguity_gap: resolvedGap,
                    state_outcomes: outcomesStateFallback,
                    state_wsu: wsuStateFallback,
                    city_outcomes: outcomesCityFallback,
                    city_wsu: wsuCityFallback,
                    country_outcomes: outcomesCountryFallback,
                    country_wsu: wsuCountryFallback
                }
            },
            (stage, processed, total) => {
                if (progressStage && progressPercent && progressBar) {
                    let percent = lastValidationPercent;
                    if (stage === 'merge') {
                        progressStage.textContent = 'Merging data...';
                        percent = Math.max(percent, 10);
                    } else if (stage === 'validate') {
                        progressStage.textContent = 'Validating mappings...';
                        const validatePercent = total
                            ? 10 + Math.round((processed / total) * 85)
                            : 10;
                        percent = Math.max(percent, validatePercent);
                    } else {
                        progressStage.textContent = 'Analyzing mappings...';
                        percent = Math.max(percent, 10);
                    }
                    lastValidationPercent = Math.min(percent, 100);
                    progressPercent.textContent = `${lastValidationPercent}%`;
                    progressBar.style.width = `${lastValidationPercent}%`;
                }
                const message = stage === 'merge'
                    ? 'Merging data...'
                    : stage === 'validate'
                        ? 'Validating mappings...'
                        : 'Analyzing mappings...';
                document.getElementById('loading-message').textContent = message;
            }
        );

        validatedData = result.validatedData;
        missingData = result.missingData;
        stats = result.stats;

        const limit = showAllErrors ? 0 : 10;
        const errorSamples = getErrorSamples(validatedData, limit);

        if (progressStage && progressPercent && progressBar) {
            progressStage.textContent = 'Complete';
            progressPercent.textContent = '100%';
            progressBar.style.width = '100%';
        }

        displayResults(stats, errorSamples);
        startActionQueuePrefetch();
        document.dispatchEvent(new CustomEvent('validation-results-ready'));

    } catch (error) {
        console.error('Validation error:', error);
        alert(`Error running validation: ${error.message}`);
    } finally {
        hideLoadingUI();
        setPageBusy(false);
        setRunLock(false);
        processAvailableFiles();
    }
}

async function runGeneration() {
    if (runLocked) {
        return;
    }
    setRunLock(true);
    try {
        if (currentMode !== 'create') {
            alert('Switch to Create mode to generate a translation table.');
            return;
        }
        document.getElementById('loading').classList.remove('hidden');
        document.getElementById('results').classList.add('hidden');
        document.getElementById('loading-message').textContent = 'Generating translation table...';
        const progressWrap = document.getElementById('loading-progress');
        const progressStage = document.getElementById('loading-stage');
        const progressPercent = document.getElementById('loading-percent');
        const progressBar = document.getElementById('loading-bar');
        if (progressWrap && progressStage && progressPercent && progressBar) {
            progressWrap.classList.remove('hidden');
            progressStage.textContent = 'Preparing...';
            progressPercent.textContent = '0%';
            progressBar.style.width = '0%';
        }

        await new Promise(resolve => setTimeout(resolve, 100));

        updateKeyConfig();

        const nameCompareOutcomes = document.getElementById('name-compare-outcomes')?.value || '';
        const nameCompareWsu = document.getElementById('name-compare-wsu')?.value || '';
        const outcomesStateRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'State'
        ) || '';
        const wsuStateRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'State'
        ) || '';
        const outcomesCityRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'City'
        ) || '';
        const wsuCityRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'City'
        ) || '';
        const outcomesCountryRole = Object.keys(columnRoles.outcomes || {}).find(
            col => columnRoles.outcomes[col] === 'Country'
        ) || '';
        const wsuCountryRole = Object.keys(columnRoles.wsu_org || {}).find(
            col => columnRoles.wsu_org[col] === 'Country'
        ) || '';
        const findFallbackColumn = (columns, token) => (
            columns.find(col => String(col).toLowerCase().includes(token)) || ''
        );
        const outcomesStateFallback = outcomesStateRole
            || findFallbackColumn(selectedColumns.outcomes, 'state');
        const wsuStateFallback = wsuStateRole
            || findFallbackColumn(selectedColumns.wsu_org, 'state');
        const outcomesCityFallback = outcomesCityRole
            || findFallbackColumn(selectedColumns.outcomes, 'city');
        const wsuCityFallback = wsuCityRole
            || findFallbackColumn(selectedColumns.wsu_org, 'city');
        const outcomesCountryFallback = outcomesCountryRole
            || findFallbackColumn(selectedColumns.outcomes, 'country');
        const wsuCountryFallback = wsuCountryRole
            || findFallbackColumn(selectedColumns.wsu_org, 'country');
        const nameCompareThreshold = parseFloat(
            document.getElementById('name-compare-threshold')?.value || '0.8'
        );
        const nameCompareGap = parseFloat(
            document.getElementById('name-compare-ambiguity-gap')?.value || '0.03'
        );
        const resolvedThreshold = Number.isNaN(nameCompareThreshold)
            ? 0.8
            : Math.max(0, Math.min(1, nameCompareThreshold));
        const resolvedGap = Number.isNaN(nameCompareGap)
            ? 0.03
            : Math.max(0, Math.min(0.2, nameCompareGap));

        const nameCompareEnabled = Boolean(nameCompareOutcomes && nameCompareWsu);
        const hasKeyConfig = Boolean(keyConfig.outcomes && keyConfig.wsu);
        const canNameMatch = Boolean(nameCompareEnabled && nameCompareOutcomes && nameCompareWsu);
        const forceNameMatch = matchMethod === 'name';

        if ((forceNameMatch && !canNameMatch) || (!forceNameMatch && !hasKeyConfig && !canNameMatch)) {
            alert('Select key columns or enable name comparison to generate a table.');
            return;
        }

        if (nameCompareEnabled && (!nameCompareOutcomes || !nameCompareWsu)) {
            alert('Select both name columns or disable name comparison.');
            return;
        }

        setPageBusy(true);
        const generated = await runWorkerTask('generate', {
            outcomes: loadedData.outcomes,
            wsu_org: loadedData.wsu_org,
            keyConfig,
            nameCompare: {
                enabled: Boolean(nameCompareEnabled),
                outcomes_column: nameCompareOutcomes,
                wsu_column: nameCompareWsu,
                threshold: resolvedThreshold,
                ambiguity_gap: resolvedGap,
                state_outcomes: outcomesStateFallback,
                state_wsu: wsuStateFallback,
                city_outcomes: outcomesCityFallback,
                city_wsu: wsuCityFallback,
                country_outcomes: outcomesCountryFallback,
                country_wsu: wsuCountryFallback
            },
            options: {
                forceNameMatch
            },
            selectedColumns,
            keyLabels
        }, (stage, processed, total) => {
            if (!progressStage || !progressPercent || !progressBar) {
                return;
            }
            const percent = total ? Math.round((processed / total) * 100) : 0;
            if (stage === 'match_candidates') {
                progressStage.textContent = 'Scoring name matches...';
                document.getElementById('loading-message').textContent =
                    `Scoring name matches... ${processed.toLocaleString()} / ${total.toLocaleString()} (${percent}%)`;
            } else if (stage === 'build_rows') {
                progressStage.textContent = 'Building output rows...';
                document.getElementById('loading-message').textContent =
                    `Building output rows... ${processed.toLocaleString()} / ${total.toLocaleString()} (${percent}%)`;
            } else {
                progressStage.textContent = 'Generating translation table...';
                document.getElementById('loading-message').textContent = 'Generating translation table...';
            }
            progressPercent.textContent = `${percent}%`;
            progressBar.style.width = `${percent}%`;
        });

        await createGeneratedTranslationExcel(
            generated.cleanRows,
            generated.errorRows,
            generated.selectedColumns,
            generated.headerLabels,
            generated.generationConfig,
            {
                onProgress: (stage, percent) => {
                    if (progressStage && progressPercent && progressBar) {
                        progressStage.textContent = stage;
                        progressPercent.textContent = `${percent}%`;
                        progressBar.style.width = `${percent}%`;
                    }
                    const loadingMessage = document.getElementById('loading-message');
                    if (loadingMessage) {
                        loadingMessage.textContent = stage;
                    }
                }
            }
        );

    } catch (error) {
        console.error('Generation error:', error);
        alert(`Error generating translation table: ${error.message}`);
    } finally {
        hideLoadingUI();
        setPageBusy(false);
        setRunLock(false);
        processAvailableFiles();
    }
}

async function createGeneratedTranslationExcel(
    cleanRows,
    errorRows,
    selectedCols,
    headerLabels,
    generationConfig = {},
    options = {}
) {
    const onProgress = typeof options.onProgress === 'function'
        ? options.onProgress
        : null;
    const result = await runExportWorkerTask(
        'build_generation_export',
        {
            cleanRows,
            errorRows,
            selectedCols,
            headerLabels,
            generationConfig
        },
        (stage, processed, total) => {
            if (!onProgress) return;
            const percent = total ? Math.round((processed / total) * 100) : 0;
            onProgress(stage, percent);
        }
    );
    downloadArrayBuffer(result.buffer, result.filename || 'Generated_Translation_Table.xlsx');
}

function getCurrentActionQueueRows() {
    return preEditedActionQueueRows || actionQueueRowsCache || [];
}

function isRowReviewed(row) {
    const value = row?.Reviewed;
    if (value === true) return true;
    const normalized = String(value ?? '').trim().toLowerCase();
    return ['1', 'true', 'yes', 'y'].includes(normalized);
}

function isRowUnresolvedForReview(row) {
    const hasDecision = String(row?.Decision || '').trim() !== '';
    return !isRowReviewed(row) || !hasDecision;
}

function getQueueErrorKey(row) {
    const type = String(row?.Error_Type || '').trim();
    const subtype = String(row?.Error_Subtype || '').trim();
    if (type === 'Output_Not_Found' && subtype === 'Output_Not_Found_Likely_Stale_Key') {
        return 'Output_Not_Found_Likely_Stale_Key';
    }
    return type;
}

function createEmptyErrorSamples() {
    const keys = [
        'Input_Not_Found',
        'Output_Not_Found',
        'Output_Not_Found_Likely_Stale_Key',
        'Duplicate_Target',
        'Duplicate_Source',
        'Name_Mismatch',
        'Ambiguous_Match'
    ];
    return keys.reduce((acc, key) => {
        acc[key] = { count: 0, showing: 0, rows: [] };
        return acc;
    }, {});
}

function buildErrorSamplesFromQueue(rows, limit = 10, unresolvedOnly = true) {
    const samples = createEmptyErrorSamples();
    const resolvedLimit = limit && limit > 0 ? limit : null;
    const reviewRows = (rows || []).filter(row => {
        if (!unresolvedOnly) return true;
        return isRowUnresolvedForReview(row);
    });

    reviewRows.forEach(row => {
        const key = getQueueErrorKey(row);
        const bucket = samples[key];
        if (!bucket) return;
        bucket.count += 1;
        if (!resolvedLimit || bucket.rows.length < resolvedLimit) {
            bucket.rows.push({
                translate_input: row.translate_input || '',
                translate_output: row.translate_output || '',
                Error_Description: row.Mapping_Logic || row.Recommended_Action || row.Error_Description || ''
            });
        }
    });

    Object.keys(samples).forEach(key => {
        samples[key].showing = samples[key].rows.length;
    });

    return samples;
}

function buildChartErrorsFromQueue(rows, unresolvedOnly = true) {
    const counts = {
        input_not_found: 0,
        output_not_found: 0,
        output_not_found_likely_stale_key: 0,
        output_not_found_ambiguous_replacement: 0,
        output_not_found_no_replacement: 0,
        duplicate_targets: 0,
        duplicate_sources: 0,
        name_mismatches: 0,
        ambiguous_matches: 0,
        high_confidence_matches: 0
    };
    (rows || []).forEach(row => {
        if (unresolvedOnly && !isRowUnresolvedForReview(row)) return;
        const type = String(row?.Error_Type || '').trim();
        const subtype = String(row?.Error_Subtype || '').trim();
        if (type === 'Input_Not_Found') counts.input_not_found += 1;
        if (type === 'Output_Not_Found') {
            counts.output_not_found += 1;
            if (subtype === 'Output_Not_Found_Likely_Stale_Key') counts.output_not_found_likely_stale_key += 1;
            if (subtype === 'Output_Not_Found_Ambiguous_Replacement') counts.output_not_found_ambiguous_replacement += 1;
            if (subtype === 'Output_Not_Found_No_Replacement') counts.output_not_found_no_replacement += 1;
        }
        if (type === 'Duplicate_Target') counts.duplicate_targets += 1;
        if (type === 'Duplicate_Source') counts.duplicate_sources += 1;
        if (type === 'Name_Mismatch') counts.name_mismatches += 1;
        if (type === 'Ambiguous_Match') counts.ambiguous_matches += 1;
    });
    return counts;
}

function updateUnresolvedErrorsToggleState() {
    const unresolvedCheckbox = document.getElementById('show-unresolved-errors-only');
    if (!unresolvedCheckbox) return;
    const hasQueue = getCurrentActionQueueRows().length > 0;
    unresolvedCheckbox.disabled = !hasQueue;
    if (!hasQueue) {
        showUnresolvedErrorsOnly = false;
        unresolvedCheckbox.checked = false;
        return;
    }
    if (!unresolvedCheckbox.checked) unresolvedCheckbox.checked = true;
    showUnresolvedErrorsOnly = unresolvedCheckbox.checked;
}

function refreshErrorPresentation() {
    if (!validatedData.length) return;
    const limit = showAllErrors ? 0 : 10;
    const queueRows = getCurrentActionQueueRows();
    if (showUnresolvedErrorsOnly && queueRows.length) {
        displayErrorDetails(buildErrorSamplesFromQueue(queueRows, limit, true));
        createErrorChart(buildChartErrorsFromQueue(queueRows, true));
        return;
    }
    displayErrorDetails(getErrorSamples(validatedData, limit));
    createErrorChart(stats.errors || {});
}

function displayResults(stats, errorSamples) {
    document.getElementById('total-mappings').textContent = stats.validation.total_mappings.toLocaleString();
    document.getElementById('valid-count').textContent = stats.validation.valid_count.toLocaleString();
    document.getElementById('valid-percentage').textContent = stats.validation.valid_percentage + '%';
    const duplicateCount = (stats.errors.duplicate_targets || 0) + (stats.errors.duplicate_sources || 0);
    const displayErrorCount = Math.max(0, stats.validation.error_count - duplicateCount);
    const displayErrorPercentage = stats.validation.total_mappings
        ? Math.round((displayErrorCount / stats.validation.total_mappings) * 1000) / 10
        : 0;
    document.getElementById('error-count').textContent = displayErrorCount.toLocaleString();
    document.getElementById('error-percentage').textContent = `${displayErrorPercentage}%`;
    document.getElementById('quality-score').textContent = stats.validation.valid_percentage + '%';

    updateUnresolvedErrorsToggleState();
    if (showUnresolvedErrorsOnly && getCurrentActionQueueRows().length) {
        refreshErrorPresentation();
    } else {
        createErrorChart(stats.errors);
        displayErrorDetails(errorSamples);
    }
    renderMappingLogicPreview();
    renderMatchingRulesExamples();

    document.getElementById('results').classList.remove('hidden');

    document.getElementById('results').scrollIntoView({ behavior: 'smooth' });
}

function formatScorePercent(score) {
    if (!Number.isFinite(score)) return '';
    return `${Math.round(Math.max(0, Math.min(1, score)) * 100)}%`;
}

function getConfiguredValue(row, prefix, fieldName) {
    if (!fieldName) return '';
    const key = `${prefix}_${fieldName}`;
    return row?.[key] ?? '';
}

function buildMatchingExample(row) {
    const threshold = Number.isFinite(lastNameCompareConfig?.threshold)
        ? lastNameCompareConfig.threshold
        : 0.8;
    const outcomesNameKey = lastNameCompareConfig?.outcomes
        ? `outcomes_${lastNameCompareConfig.outcomes}`
        : '';
    const wsuNameKey = lastNameCompareConfig?.wsu
        ? `wsu_${lastNameCompareConfig.wsu}`
        : '';
    const outcomesName = outcomesNameKey ? (row[outcomesNameKey] || row.outcomes_name || '') : (row.outcomes_name || '');
    const wsuName = wsuNameKey ? (row[wsuNameKey] || row.wsu_Descr || '') : (row.wsu_Descr || '');
    const normalizedOutcomes = typeof normalizeNameForCompare === 'function'
        ? normalizeNameForCompare(outcomesName)
        : outcomesName;
    const normalizedWsu = typeof normalizeNameForCompare === 'function'
        ? normalizeNameForCompare(wsuName)
        : wsuName;
    const similarity = (outcomesName && wsuName && typeof calculateNameSimilarity === 'function')
        ? calculateNameSimilarity(outcomesName, wsuName)
        : null;
    const similarityText = formatScorePercent(similarity);

    const outcomesState = getConfiguredValue(row, 'outcomes', lastNameCompareConfig?.state_outcomes);
    const wsuState = getConfiguredValue(row, 'wsu', lastNameCompareConfig?.state_wsu);
    const outcomesCity = getConfiguredValue(row, 'outcomes', lastNameCompareConfig?.city_outcomes);
    const wsuCity = getConfiguredValue(row, 'wsu', lastNameCompareConfig?.city_wsu);
    const outcomesCountry = getConfiguredValue(row, 'outcomes', lastNameCompareConfig?.country_outcomes);
    const wsuCountry = getConfiguredValue(row, 'wsu', lastNameCompareConfig?.country_wsu);

    const evidence = [];
    if (similarityText) {
        evidence.push(`Similarity ${similarityText} (threshold ${formatScorePercent(threshold)})`);
    }
    if (outcomesState && wsuState && typeof statesMatch === 'function' && statesMatch(outcomesState, wsuState)) {
        evidence.push(`State match: ${outcomesState} = ${wsuState}`);
    }
    if (outcomesCountry && wsuCountry && typeof countriesMatch === 'function' && countriesMatch(outcomesCountry, wsuCountry)) {
        evidence.push(`Country match: ${outcomesCountry} = ${wsuCountry}`);
    }
    if (typeof cityInName === 'function' && (cityInName(outcomesName, wsuCity) || cityInName(wsuName, outcomesCity))) {
        evidence.push('City name appears in the other side');
    }
    if (typeof locationInNameMatches === 'function' && (
        locationInNameMatches(outcomesName, wsuCity, wsuState) ||
        locationInNameMatches(wsuName, outcomesCity, outcomesState)
    )) {
        evidence.push('Parenthetical/hyphen location token matched');
    }
    if (row.Error_Type === 'High_Confidence_Match') {
        evidence.push('High-confidence override applied');
    }

    return `
        <div class="border border-gray-200 rounded p-3">
            <p class="text-sm"><strong>Outcomes:</strong> ${escapeHtml(outcomesName || '—')}</p>
            <p class="text-sm"><strong>myWSU:</strong> ${escapeHtml(wsuName || '—')}</p>
            <p class="text-xs text-gray-500 mt-2"><strong>Normalized:</strong> ${escapeHtml(normalizedOutcomes || '—')} ↔ ${escapeHtml(normalizedWsu || '—')}</p>
            <p class="text-xs text-gray-500 mt-1"><strong>Evidence:</strong> ${escapeHtml(evidence.join(' | ') || 'No evidence available')}</p>
            <p class="text-xs text-gray-500 mt-1"><strong>Decision:</strong> ${escapeHtml(normalizeErrorTypeForPreview(row.Error_Type) || row.Error_Type || '—')}</p>
        </div>
    `;
}

function renderMatchingRulesExamples() {
    const toggle = document.getElementById('show-matching-rules');
    const panel = document.getElementById('matching-rules-panel');
    const container = document.getElementById('matching-examples');
    if (!container || !panel || !toggle) return;
    if (!toggle.checked) {
        panel.classList.add('hidden');
        return;
    }
    panel.classList.remove('hidden');
    if (!validatedData.length) {
        container.innerHTML = '<p class="text-xs text-gray-500">Examples from this run will appear after validation.</p>';
        return;
    }

    const priority = [
        'High_Confidence_Match',
        'Name_Mismatch',
        'Ambiguous_Match',
        'Output_Not_Found',
        'Valid'
    ];
    const examples = [];
    priority.forEach(type => {
        if (examples.length >= 2) return;
        const match = validatedData.find(row => row.Error_Type === type && !examples.includes(row));
        if (match) examples.push(match);
    });
    if (examples.length < 2) {
        validatedData.slice(0, 2 - examples.length).forEach(row => examples.push(row));
    }

    const html = examples.map(example => buildMatchingExample(example)).join('');
    container.innerHTML = `
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            ${html || '<p class="text-xs text-gray-500">No examples available.</p>'}
        </div>
    `;
}

function normalizeErrorTypeForPreview(errorType) {
    if (errorType === 'Input_Not_Found') return 'Input key not found in Outcomes';
    if (errorType === 'Output_Not_Found') return 'Output key not found in myWSU';
    if (errorType === 'Missing_Input') return 'Input key is blank in Translate';
    if (errorType === 'Missing_Output') return 'Output key is blank in Translate';
    if (errorType === 'Name_Mismatch') return 'Name mismatch';
    if (errorType === 'Ambiguous_Match') return 'Ambiguous name match';
    return errorType || '';
}

function buildMappingLogicPreviewText(row) {
    const threshold = Number.isFinite(lastNameCompareConfig?.threshold)
        ? lastNameCompareConfig.threshold
        : 0.8;
    const thresholdText = formatScorePercent(threshold);
    const outcomesNameKey = lastNameCompareConfig?.outcomes
        ? `outcomes_${lastNameCompareConfig.outcomes}`
        : '';
    const wsuNameKey = lastNameCompareConfig?.wsu
        ? `wsu_${lastNameCompareConfig.wsu}`
        : '';
    const outcomesName = outcomesNameKey ? (row[outcomesNameKey] || row.outcomes_name || '') : (row.outcomes_name || '');
    const wsuName = wsuNameKey ? (row[wsuNameKey] || row.wsu_Descr || '') : (row.wsu_Descr || '');
    const similarity = (outcomesName && wsuName && typeof calculateNameSimilarity === 'function')
        ? calculateNameSimilarity(outcomesName, wsuName)
        : null;
    const similarityText = formatScorePercent(similarity);

    switch (row.Error_Type) {
    case 'Valid':
        if (similarityText) {
            return `Valid: key checks passed and name similarity ${similarityText} met/exceeded threshold ${thresholdText}.`;
        }
        return 'Valid: key checks passed and no blocking validation rules fired.';
    case 'High_Confidence_Match':
        return similarityText
            ? `High confidence override: similarity ${similarityText} was below threshold ${thresholdText}, but alias/location/token rules confirmed this mapping.`
            : `High confidence override: alias/location/token rules confirmed this mapping below threshold ${thresholdText}.`;
    case 'Duplicate_Target':
        return 'Many-to-one duplicate: multiple source keys map to the same target key.';
    case 'Duplicate_Source':
        return 'One-to-many duplicate: one source key maps to multiple target keys.';
    case 'Input_Not_Found':
        return 'Key lookup failed: translate input key value is present but was not found in Outcomes keys.';
    case 'Output_Not_Found':
        if (row.Error_Subtype === 'Output_Not_Found_Likely_Stale_Key') {
            const suggested = row.Suggested_Key ? ` Suggested replacement key: ${row.Suggested_Key}.` : '';
            return `Key lookup failed: translate output key was not found in myWSU keys. Likely stale key.${suggested}`;
        }
        if (row.Error_Subtype === 'Output_Not_Found_Ambiguous_Replacement') {
            return 'Key lookup failed: translate output key was not found in myWSU keys. Multiple high-confidence replacement candidates were found.';
        }
        if (row.Error_Subtype === 'Output_Not_Found_No_Replacement') {
            return 'Key lookup failed: translate output key was not found in myWSU keys. No high-confidence replacement candidate was found.';
        }
        return 'Key lookup failed: translate output key value is present but was not found in myWSU keys.';
    case 'Missing_Input':
        return 'Translate row has a blank input key cell.';
    case 'Missing_Output':
        return 'Translate row has a blank output key cell.';
    case 'Name_Mismatch':
        return similarityText
            ? `Name comparison failed: similarity ${similarityText} is below threshold ${thresholdText}.`
            : 'Name comparison failed: below configured threshold.';
    case 'Ambiguous_Match':
        return 'Name comparison ambiguous: another candidate scored within the ambiguity gap.';
    default:
        return row.Error_Description || 'Classified by validation rules.';
    }
}

function renderMappingLogicPreview() {
    const toggle = document.getElementById('show-logic-preview');
    const panel = document.getElementById('logic-preview-panel');
    const body = document.getElementById('logic-preview-body');
    const summary = document.getElementById('logic-preview-summary');
    if (!toggle || !panel || !body || !summary) return;

    if (!toggle.checked || !validatedData.length) {
        panel.classList.add('hidden');
        body.innerHTML = '';
        summary.textContent = '';
        return;
    }

    const maxRows = 200;
    const rows = validatedData.slice(0, maxRows);
    const rowsHtml = rows.map((row, index) => {
        const subtype = row.Error_Subtype ? ` (${row.Error_Subtype})` : '';
        const classification = `${normalizeErrorTypeForPreview(row.Error_Type)}${subtype}`;
        const logicText = buildMappingLogicPreviewText(row);
        return `
            <tr class="border-b align-top">
                <td class="py-2 px-3 text-xs text-gray-500">${index + 1}</td>
                <td class="py-2 px-3 text-sm">${escapeHtml(row.translate_input)}</td>
                <td class="py-2 px-3 text-sm">${escapeHtml(row.translate_output)}</td>
                <td class="py-2 px-3 text-sm">${escapeHtml(classification)}</td>
                <td class="py-2 px-3 text-sm">${escapeHtml(logicText)}</td>
            </tr>
        `;
    }).join('');

    body.innerHTML = rowsHtml;
    summary.textContent = validatedData.length > maxRows
        ? `Showing first ${maxRows.toLocaleString()} of ${validatedData.length.toLocaleString()} rows.`
        : `Showing all ${validatedData.length.toLocaleString()} rows.`;
    panel.classList.remove('hidden');
}

function createErrorChart(errors) {
    const ctx = document.getElementById('error-chart').getContext('2d');

    if (window.errorChart) {
        window.errorChart.destroy();
    }

    const data = {
        labels: [
            'Input Keys Not Found in Outcomes',
            'Output Keys Not Found in myWSU',
            'Duplicate Target Keys',
            'Duplicate Source Keys',
            'Name Mismatches',
            'Ambiguous Matches',
            'High Confidence Matches'
        ],
        datasets: [{
            label: 'Error Count',
            data: [
                errors.input_not_found,
                errors.output_not_found,
                errors.duplicate_targets,
                errors.duplicate_sources,
                errors.name_mismatches,
                errors.ambiguous_matches,
                errors.high_confidence_matches || 0
            ],
            backgroundColor: [
                'rgba(239, 68, 68, 0.8)',   // Red
                'rgba(249, 115, 22, 0.8)',  // Orange
                'rgba(251, 191, 36, 0.8)',  // Yellow
                'rgba(59, 130, 246, 0.8)',  // Blue
                'rgba(14, 116, 144, 0.8)',  // Teal
                'rgba(234, 179, 8, 0.8)',   // Amber
                'rgba(168, 85, 247, 0.8)',  // Purple
                'rgba(74, 222, 128, 0.8)'   // Green
            ],
            borderColor: [
                'rgb(239, 68, 68)',
                'rgb(249, 115, 22)',
                'rgb(251, 191, 36)',
                'rgb(59, 130, 246)',
                'rgb(14, 116, 144)',
                'rgb(202, 138, 4)',
                'rgb(147, 51, 234)',
                'rgb(22, 163, 74)'
            ],
            borderWidth: 2
        }]
    };

    window.errorChart = new Chart(ctx, {
        type: 'bar',
        data: data,
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.parsed.y + ' errors';
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        }
    });
}

function displayErrorDetails(errorSamples) {
    const detailsDiv = document.getElementById('error-details');
    detailsDiv.innerHTML = '';

    const errorTypes = [
        { key: 'Input_Not_Found', title: 'Input Keys Not Found in Outcomes', color: 'orange' },
        { key: 'Output_Not_Found', title: 'Output Keys Not Found in myWSU', color: 'orange' },
        { key: 'Output_Not_Found_Likely_Stale_Key', title: 'Likely Stale Output Keys (Suggested Replacements)', color: 'orange' },
        { key: 'Duplicate_Target', title: 'Duplicate Target Keys (Many-to-One Errors)', color: 'yellow' },
        { key: 'Duplicate_Source', title: 'Duplicate Source Keys', color: 'yellow' },
        { key: 'Name_Mismatch', title: 'Name Mismatches (Possible Wrong Mappings)', color: 'yellow' },
        { key: 'Ambiguous_Match', title: 'Ambiguous Matches (Check Alternatives)', color: 'yellow' },
    ];

    const cardsHtml = [];
    errorTypes.forEach(errorType => {
        const sample = errorSamples[errorType.key];
        if (sample && sample.count > 0) {
            cardsHtml.push(createErrorCard(errorType.title, sample, errorType.color));
        }
    });
    detailsDiv.innerHTML = cardsHtml.join('');
    detailsDiv.querySelectorAll('.error-card-toggle').forEach(button => {
        button.addEventListener('click', function() {
            const targetId = button.getAttribute('data-target');
            const panel = targetId ? document.getElementById(targetId) : null;
            const chevron = button.querySelector('[data-chevron]');
            if (!panel) return;
            panel.classList.toggle('hidden');
            if (chevron) chevron.classList.toggle('rotate-180');
        });
    });
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function createErrorCard(title, sample, color) {
    const colorClasses = {
        red: 'border-red-500 bg-red-50',
        orange: 'border-orange-500 bg-orange-50',
        yellow: 'border-yellow-500 bg-yellow-50'
    };

    const explanations = {
        'Input Keys Not Found in Outcomes': {
            icon: '🔴',
            text: 'Translate input key value is present, but it does not exist in Outcomes keys.',
            impact: 'Critical - Data entry/config issue; mapping will fail'
        },
        'Output Keys Not Found in myWSU': {
            icon: '🔴',
            text: 'Translate output key value is present, but it does not exist in myWSU keys.',
            impact: 'Critical - Data entry/config issue; mapping will fail'
        },
        'Likely Stale Output Keys (Suggested Replacements)': {
            icon: '[!]',
            text: 'Output key value is not found in myWSU, but one high-confidence replacement key was found using name and location evidence.',
            impact: 'Critical - Likely stale key; verify then update translate output key'
        },
        'Duplicate Target Keys (Many-to-One Errors)': {
            icon: '🟠',
            text: 'Multiple different source keys map to the SAME target key. Multiple Outcomes records are pointing to one myWSU record.',
            impact: 'Critical - Multiple Outcomes records will be merged into one target record'
        },
        'Duplicate Source Keys': {
            icon: '',
            text: 'The same source key maps to multiple target keys. This creates conflicting mappings for a single Outcomes record.',
            impact: 'Critical - Fix conflicting mappings for the same source record'
        },
        'Name Mismatches (Possible Wrong Mappings)': {
            icon: '⚠️',
            text: 'Names do not match between Outcomes and myWSU (below similarity threshold). These may be incorrect mappings that need review.',
            impact: 'Warning - Review these mappings to ensure correct records'
        },
        'Ambiguous Matches (Check Alternatives)': {
            icon: '⚠️',
            text: 'Name match is ambiguous (another candidate is within the ambiguity gap). Review alternatives before accepting.',
            impact: 'Warning - Review ambiguous matches to confirm correctness'
        },
    };

    const explanation = explanations[title] || { icon: '', text: '', impact: '' };

    const rowsHtml = sample.rows.map(row => `
        <tr class="border-b">
            <td class="py-2 px-4 text-sm">${escapeHtml(row.translate_input)}</td>
            <td class="py-2 px-4 text-sm">${escapeHtml(row.translate_output)}</td>
            <td class="py-2 px-4 text-sm">${escapeHtml(row.Error_Description)}</td>
        </tr>
    `).join('');

    const showingLine = sample.showing < sample.count
        ? `Showing first ${sample.showing} of ${sample.count} errors - download Excel for complete list`
        : `Showing all ${sample.count} errors`;
    const cardId = `error-card-${String(title).toLowerCase().replace(/[^a-z0-9]+/g, '-')}`;

    return `
        <div class="bg-white rounded-lg shadow-md p-6 border-l-4 ${colorClasses[color]}">
            <button type="button" data-target="${cardId}" class="error-card-toggle w-full flex items-start justify-between text-left">
                <div class="pr-2">
                    <h3 class="text-lg font-semibold text-gray-800 mb-1">${explanation.icon} ${title}</h3>
                    <p class="text-sm text-gray-700 mb-2">${explanation.text}</p>
                    <p class="text-xs font-semibold text-${color}-700 uppercase">${explanation.impact}</p>
                </div>
                <div class="flex items-center gap-2">
                    <div class="bg-${color}-100 rounded-full px-4 py-2">
                        <span class="text-2xl font-bold text-${color}-700">${sample.count}</span>
                    </div>
                    <svg data-chevron class="h-5 w-5 text-gray-500 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"/>
                    </svg>
                </div>
            </button>
            <div id="${cardId}" class="hidden mt-4">
                <p class="text-xs text-gray-500 mb-4">${showingLine}</p>
                <div class="overflow-x-auto">
                    <table class="min-w-full divide-y divide-gray-200">
                        <thead class="bg-gray-100">
                            <tr>
                                <th class="py-2 px-4 text-left text-xs font-medium text-gray-700 uppercase">${keyLabels.translateInput || 'Source key'}</th>
                                <th class="py-2 px-4 text-left text-xs font-medium text-gray-700 uppercase">${keyLabels.translateOutput || 'Target key'}</th>
                                <th class="py-2 px-4 text-left text-xs font-medium text-gray-700 uppercase">Description</th>
                            </tr>
                        </thead>
                        <tbody class="bg-white divide-y divide-gray-200">
                            ${rowsHtml}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

function setupDownloadButton() {
    const downloadBtn = document.getElementById('download-btn');
    const progressWrap = document.getElementById('download-progress');
    const progressText = document.getElementById('download-progress-text');
    const progressBar = document.getElementById('download-progress-bar');
    downloadBtn.addEventListener('click', async function() {
        if (validatedData.length === 0) {
            alert('Please run validation first.');
            return;
        }

        try {
            downloadBtn.disabled = true;
            downloadBtn.innerHTML = '<span class="inline-block animate-spin mr-2">⏳</span> Generating...';
            if (progressWrap && progressText && progressBar) {
                progressWrap.classList.remove('hidden');
                progressText.textContent = 'Preparing export...';
                progressBar.style.width = '0%';
            }
            setPageBusy(true);

            const includeSuggestions = Boolean(
                document.getElementById('include-suggestions')?.checked
            );
            const showMappingLogic = Boolean(
                document.getElementById('show-mapping-logic')?.checked
            );
            const reviewScope = getBulkReviewScope();
            const result = await createExcelOutput(
                validatedData,
                missingData,
                selectedColumns,
                {
                    includeSuggestions,
                    showMappingLogic,
                    reviewScope,
                    nameCompareConfig: lastNameCompareConfig,
                    priorDecisions: priorDecisions || undefined,
                    campusFamilyRules: campusFamilyRules || undefined,
                    preEditedActionQueueRows: preEditedActionQueueRows || undefined,
                    onProgress: (stage, percent) => {
                        if (progressText && progressBar) {
                            progressText.textContent = stage;
                            progressBar.style.width = `${percent}%`;
                        }
                    }
                }
            );

            downloadBtn.disabled = false;
            downloadBtn.innerHTML = `
                <svg class="h-6 w-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                </svg>
                Download Full Report
            `;
            if (progressWrap && progressText && progressBar) {
                const s = result?.reimportSummary;
                const summary = s
                    ? `Re-import: ${s.applied ?? 0} applied, ${s.conflicts ?? 0} conflicts, ${s.newRows ?? 0} new rows, ${s.orphaned ?? 0} orphaned.`
                    : 'Download ready.';
                progressText.textContent = summary;
                progressBar.style.width = '100%';
                setTimeout(() => {
                    progressWrap.classList.add('hidden');
                }, 1500);
            }

        } catch (error) {
            console.error('Download error:', error);
            if (error.exportStack) {
                console.error('Export worker stack:', error.exportStack);
            }
            alert(`Error generating Excel: ${error.message}`);
            downloadBtn.disabled = false;
            downloadBtn.innerHTML = `
                <svg class="h-6 w-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                </svg>
                Download Full Report
            `;
            if (progressWrap) {
                progressWrap.classList.add('hidden');
            }
        } finally {
            setPageBusy(false);
        }
    });
}

async function createExcelOutput(validated, missing, selectedCols, options = {}) {
    const onProgressCb = typeof options.onProgress === 'function'
        ? options.onProgress
        : null;
    const reviewScope = normalizeBulkReviewScope(
        options.reviewScope || (options.translationOnlyExport ? 'translation_only' : 'all')
    );
    const payload = {
        validated,
        missing,
        selectedCols,
        priorDecisions: options.priorDecisions || null,
        options: {
            includeSuggestions: Boolean(options.includeSuggestions),
            showMappingLogic: Boolean(options.showMappingLogic),
            reviewScope,
            // Legacy flag kept for compatibility with older consumers/tests.
            translationOnlyExport: reviewScope === 'translation_only',
            nameCompareConfig: options.nameCompareConfig || {},
            campusFamilyRules: options.campusFamilyRules || null
        },
        preEditedActionQueueRows: options.preEditedActionQueueRows || null,
        context: {
            loadedData,
            columnRoles,
            keyConfig,
            keyLabels
        }
    };
    const result = await runExportWorkerTask(
        'build_validation_export',
        payload,
        (stage, processed, total) => {
            if (!onProgressCb) return;
            const percent = total ? Math.round((processed / total) * 100) : 0;
            onProgressCb(stage, percent);
        }
    );
    downloadArrayBuffer(result.buffer, result.filename || 'WSU_Mapping_Validation_Report.xlsx');
    return result;
}

function applySessionDataToActionQueue(sessionRows, options = {}) {
    const rows = getCurrentActionQueueRows();
    if (!rows.length) return false;
    const sourceLabel = options.sourceLabel || 'Session';
    const silent = Boolean(options.silent);
    const toDisplay = (value) => String(value ?? '').trim();
    const rowMap = new Map(rows.map(row => [String(row.Review_Row_ID || ''), row]));
    let applied = 0;
    let missing = 0;

    (sessionRows || []).forEach(entry => {
        const rid = String(entry?.Review_Row_ID || '').trim();
        if (!rid) return;
        const row = rowMap.get(rid);
        if (!row) {
            missing += 1;
            return;
        }
        row.Decision = toDisplay(entry.Decision);
        row.Reason_Code = toDisplay(entry.Reason_Code);
        row.Manual_Suggested_Key = toDisplay(entry.Manual_Suggested_Key);
        row.Selected_Candidate_ID = toDisplay(entry.Selected_Candidate_ID);
        if (entry && Object.prototype.hasOwnProperty.call(entry, 'Reviewed')) {
            row.Reviewed = isRowReviewed(entry);
        }
        applied += 1;
    });

    preEditedActionQueueRows = rows;
    if (!silent) {
        alert(`${sourceLabel} loaded: ${applied} rows applied${missing ? `, ${missing} unmatched Review_Row_ID` : ''}.`);
    }
    return applied > 0;
}

let actionQueueRowsCache = null;

function setupBulkEditPanel() {
    const toggleBtn = document.getElementById('bulk-edit-toggle-btn');
    const panel = document.getElementById('bulk-edit-panel');
    const filterSelect = document.getElementById('bulk-filter-error-type');
    const filterDecisionSelect = document.getElementById('bulk-filter-decision');
    const filterOutcomesNameInput = document.getElementById('bulk-filter-outcomes-name');
    const filterWsuNameInput = document.getElementById('bulk-filter-wsu-name');
    const filterReviewScope = document.getElementById('bulk-filter-review-scope');
    const pageSizeSelect = document.getElementById('bulk-page-size');
    const pageSizeBottomSelect = document.getElementById('bulk-page-size-bottom');
    const pageStatusEl = document.getElementById('bulk-page-status');
    const pageStatusBottomEl = document.getElementById('bulk-page-status-bottom');
    const pageRangeEl = document.getElementById('bulk-page-range');
    const pageRangeBottomEl = document.getElementById('bulk-page-range-bottom');
    const pagePrevBtn = document.getElementById('bulk-page-prev-btn');
    const pagePrevBottomBtn = document.getElementById('bulk-page-prev-btn-bottom');
    const pageNextBtn = document.getElementById('bulk-page-next-btn');
    const pageNextBottomBtn = document.getElementById('bulk-page-next-btn-bottom');
    const applyDecisionSelect = document.getElementById('bulk-apply-decision');
    const applyReasonCodeSelect = document.getElementById('bulk-apply-reason-code');
    const applyCandidateIdSelect = document.getElementById('bulk-apply-candidate-id');
    const applyManualInput = document.getElementById('bulk-apply-manual-key');
    const applyScopeSelect = document.getElementById('bulk-apply-scope');
    const applySelectedOnly = document.getElementById('bulk-apply-selected-only');
    const applyBtn = document.getElementById('bulk-apply-btn');
    const clearSelectedBtn = document.getElementById('bulk-clear-selected-btn');
    const selectFilteredBtn = document.getElementById('bulk-select-filtered-btn');
    const deselectFilteredBtn = document.getElementById('bulk-deselect-filtered-btn');
    const markReviewedBtn = document.getElementById('bulk-mark-reviewed-btn');
    const clearReviewedBtn = document.getElementById('bulk-clear-reviewed-btn');
    const saveSessionBtn = document.getElementById('bulk-save-session-btn');
    const loadSessionInput = document.getElementById('bulk-load-session-file');
    const quickChipButtons = panel.querySelectorAll('.bulk-quick-chip');
    const quickClearBtn = document.getElementById('bulk-quick-clear-btn');
    const tbody = document.getElementById('bulk-edit-tbody');
    const rowCountEl = document.getElementById('bulk-edit-row-count');
    const filterCountEl = document.getElementById('bulk-edit-filter-count');
    const selectedCountEl = document.getElementById('bulk-edit-selected-count');
    const reviewedCountEl = document.getElementById('bulk-edit-reviewed-count');
    const loadProgressWrap = document.getElementById('bulk-load-progress');
    const loadProgressText = document.getElementById('bulk-load-progress-text');
    const loadProgressPercent = document.getElementById('bulk-load-progress-percent');
    const loadProgressBar = document.getElementById('bulk-load-progress-bar');
    if (!toggleBtn || !panel || !tbody) return;
    const syncToggleButtonVisibility = () => {
        toggleBtn.classList.toggle('hidden', bulkEditorOpenedOnce);
    };
    syncToggleButtonVisibility();

    const DEFAULT_PAGE_SIZE = 200;
    const DECISION_OPTIONS = ['', 'Keep As-Is', 'Use Suggestion', 'Allow One-to-Many', 'Ignore'];
    const REASON_CODE_OPTIONS = [
        '',
        'Campus consolidation',
        'Data steward approved',
        'Manual correction',
        'Name match',
        'Other'
    ];
    let filteredRowsCache = [];
    let selectedRowIds = new Set();
    let currentPage = 1;
    let currentPageSize = DEFAULT_PAGE_SIZE;
    const pageSizeControls = [pageSizeSelect, pageSizeBottomSelect].filter(Boolean);
    const pageStatusControls = [pageStatusEl, pageStatusBottomEl].filter(Boolean);
    const pageRangeControls = [pageRangeEl, pageRangeBottomEl].filter(Boolean);
    const pagePrevControls = [pagePrevBtn, pagePrevBottomBtn].filter(Boolean);
    const pageNextControls = [pageNextBtn, pageNextBottomBtn].filter(Boolean);

    const normalize = (value) => String(value ?? '').trim().toLowerCase();
    const toDisplay = (value) => String(value ?? '').trim();
    const normalizeKeyToken = (value) => normalize(value).replace(/[\s-]+/g, '_');
    const isSourceDerivedMissingMapping = (row) => {
        const source = normalizeKeyToken(toDisplay(row.Source_Sheet));
        const errorType = normalizeKeyToken(toDisplay(row.Error_Type));
        return source === 'missing_mappings' || errorType === 'missing_mapping';
    };
    const encodeRowId = (id) => encodeURIComponent(String(id || ''));
    const decodeRowId = (encoded) => {
        try {
            return decodeURIComponent(String(encoded || ''));
        } catch (_) {
            return String(encoded || '');
        }
    };
    const optionHtml = (value, label, selected) => `<option value="${escapeHtml(value)}"${selected ? ' selected' : ''}>${escapeHtml(label)}</option>`;
    const decisionOptionsHtml = (value) => DECISION_OPTIONS
        .map(opt => optionHtml(opt, opt || '(blank)', opt === value))
        .join('');
    const reasonOptionsHtml = (value) => REASON_CODE_OPTIONS
        .map(opt => optionHtml(opt, opt || '(blank)', opt === value))
        .join('');
    const candidateLabel = (c) => {
        const key = toDisplay(c.key || '');
        const name = toDisplay(c.name || '');
        const loc = [toDisplay(c.city || ''), toDisplay(c.state || ''), toDisplay(c.country || '')]
            .filter(Boolean)
            .join(', ');
        const scoreVal = typeof c.score === 'number' ? c.score : parseFloat(c.score);
        const score = Number.isFinite(scoreVal) ? ` | Score: ${scoreVal.toFixed(2)}` : '';
        const base = `${key}: ${name}${loc ? ` - ${loc}` : ''}`;
        return `${base}${score}`;
    };
    const normalizeKey = (value) => String(value ?? '').trim().toLowerCase();
    const buildWsuKeyLookup = () => {
        const map = new Map();
        const wsuKeyCol = keyConfig.wsu;
        const wsuNameCol = getNameColumn('wsu_org');
        const wsuStateCol = getColumnByToken('wsu_org', 'state');
        const wsuCountryCol = getColumnByToken('wsu_org', 'country');
        const wsuCityCol = getColumnByToken('wsu_org', 'city');
        (loadedData.wsu_org || []).forEach(row => {
            const keyVal = normalizeKey(wsuKeyCol ? row[wsuKeyCol] : '');
            if (!keyVal) return;
            if (map.has(keyVal)) return;
            map.set(keyVal, {
                key: toDisplay(wsuKeyCol ? row[wsuKeyCol] : ''),
                name: toDisplay(wsuNameCol ? row[wsuNameCol] : ''),
                city: toDisplay(wsuCityCol ? row[wsuCityCol] : ''),
                state: toDisplay(wsuStateCol ? row[wsuStateCol] : ''),
                country: toDisplay(wsuCountryCol ? row[wsuCountryCol] : '')
            });
        });
        return map;
    };
    const formatPreview = (entry, badgeText = '') => {
        if (!entry) return '<span class="text-xs text-gray-500">(no override selected)</span>';
        const parts = [entry.name, entry.city, entry.state, entry.country].filter(Boolean).join(' | ');
        return `
            <div class="text-xs">${escapeHtml(entry.key || '(blank)')}</div>
            <div class="text-xs text-gray-700">${escapeHtml(parts || '(no location)')}</div>
            ${badgeText ? `<div class="text-[11px] text-orange-700 mt-1">${escapeHtml(badgeText)}</div>` : ''}
        `;
    };
    const getRoleColumn = (source, roleName) => {
        const roles = columnRoles[source] || {};
        const hit = Object.keys(roles).find(col => roles[col] === roleName);
        return hit || '';
    };
    const getNameColumn = (source) => {
        if (source === 'outcomes') {
            if (lastNameCompareConfig?.outcomes) return lastNameCompareConfig.outcomes;
            const schoolRole = getRoleColumn('outcomes', 'School');
            if (schoolRole) return schoolRole;
            const cols = selectedColumns.outcomes || [];
            return cols.find(col => /name|descr|school|org/i.test(String(col || ''))) || '';
        }
        if (lastNameCompareConfig?.wsu) return lastNameCompareConfig.wsu;
        const schoolRole = getRoleColumn('wsu_org', 'School');
        if (schoolRole) return schoolRole;
        const cols = selectedColumns.wsu_org || [];
        return cols.find(col => /name|descr|school|org/i.test(String(col || ''))) || '';
    };
    const getColumnByToken = (source, token) => {
        const roleHit = getRoleColumn(source, token.charAt(0).toUpperCase() + token.slice(1));
        if (roleHit) return roleHit;
        const cols = source === 'outcomes' ? (selectedColumns.outcomes || []) : (selectedColumns.wsu_org || []);
        return cols.find(col => String(col || '').toLowerCase().includes(token)) || '';
    };
    const getPrefixedValue = (row, sourcePrefix, colName) => {
        if (!colName) return '';
        return toDisplay(row[`${sourcePrefix}_${colName}`]);
    };
    const attachRowContext = (row) => {
        const outcomesNameCol = getNameColumn('outcomes');
        const outcomesStateCol = getColumnByToken('outcomes', 'state');
        const outcomesCountryCol = getColumnByToken('outcomes', 'country');
        const wsuNameCol = getNameColumn('wsu_org');
        const wsuStateCol = getColumnByToken('wsu_org', 'state');
        const wsuCountryCol = getColumnByToken('wsu_org', 'country');
        row._ctx = {
            outcomesName: getPrefixedValue(row, 'outcomes', outcomesNameCol),
            outcomesState: getPrefixedValue(row, 'outcomes', outcomesStateCol),
            outcomesCountry: getPrefixedValue(row, 'outcomes', outcomesCountryCol),
            wsuName: getPrefixedValue(row, 'wsu', wsuNameCol),
            wsuState: getPrefixedValue(row, 'wsu', wsuStateCol),
            wsuCountry: getPrefixedValue(row, 'wsu', wsuCountryCol)
        };
        if (!Array.isArray(row._candidates)) row._candidates = [];
        if (typeof row.Decision !== 'string') row.Decision = toDisplay(row.Decision);
        if (typeof row.Reason_Code !== 'string') row.Reason_Code = toDisplay(row.Reason_Code);
        if (typeof row.Manual_Suggested_Key !== 'string') row.Manual_Suggested_Key = toDisplay(row.Manual_Suggested_Key);
        if (typeof row.Selected_Candidate_ID !== 'string') row.Selected_Candidate_ID = toDisplay(row.Selected_Candidate_ID);
        row.Reviewed = isRowReviewed(row);
    };
    const cloneRows = (rows) => cloneActionQueueRows(rows);
    const getWorkingRows = () => preEditedActionQueueRows || actionQueueRowsCache || [];
    const parsePageSize = (value) => {
        const raw = parseInt(toDisplay(value), 10);
        return Number.isFinite(raw) && raw > 0 ? raw : DEFAULT_PAGE_SIZE;
    };
    const syncPageSizeControls = (value) => {
        pageSizeControls.forEach(control => {
            if (!control) return;
            if (toDisplay(control.value) !== toDisplay(value)) {
                control.value = toDisplay(value);
            }
        });
    };
    const setPageSize = (value) => {
        currentPageSize = parsePageSize(value);
        syncPageSizeControls(String(currentPageSize));
        return currentPageSize;
    };
    const getPageSize = () => {
        return currentPageSize;
    };
    setPageSize(pageSizeControls.find(control => control && toDisplay(control.value))?.value || DEFAULT_PAGE_SIZE);
    const getTotalPages = () => {
        const count = filteredRowsCache.length;
        const pageSize = getPageSize();
        return Math.max(1, Math.ceil(count / pageSize));
    };
    const clampCurrentPage = () => {
        const totalPages = getTotalPages();
        if (currentPage < 1) currentPage = 1;
        if (currentPage > totalPages) currentPage = totalPages;
        return totalPages;
    };
    const getCurrentPageRows = () => {
        const pageSize = getPageSize();
        const startIndex = (currentPage - 1) * pageSize;
        return filteredRowsCache.slice(startIndex, startIndex + pageSize);
    };
    const sortCandidateIds = (ids) => [...ids].sort((a, b) => {
        const matchA = /^c(\d+)$/i.exec(a);
        const matchB = /^c(\d+)$/i.exec(b);
        if (matchA && matchB) return Number(matchA[1]) - Number(matchB[1]);
        return a.localeCompare(b, undefined, { sensitivity: 'base' });
    });
    const refreshBulkCandidateApplyOptions = () => {
        if (!applyCandidateIdSelect) return;
        const prev = toDisplay(applyCandidateIdSelect.value);
        const ids = new Set();
        filteredRowsCache.forEach(row => {
            (row._candidates || []).forEach(candidate => {
                const candidateId = toDisplay(candidate?.candidateId || '');
                if (candidateId) ids.add(candidateId);
            });
        });
        const optionsHtml = [
            optionHtml('', '(no change)', prev === ''),
            ...sortCandidateIds(ids).map(candidateId => optionHtml(candidateId, candidateId, candidateId === prev))
        ].join('');
        applyCandidateIdSelect.innerHTML = optionsHtml;
        if (prev && !ids.has(prev)) {
            applyCandidateIdSelect.value = '';
        }
    };
    const findRowById = (rid) => getWorkingRows().find(row => String(row.Review_Row_ID || '') === rid) || null;
    const getFilteredRows = () => {
        const errorType = toDisplay(filterSelect?.value || '');
        const decisionFilter = toDisplay(filterDecisionSelect?.value || '');
        const outcomesContains = normalize(filterOutcomesNameInput?.value || '');
        const wsuContains = normalize(filterWsuNameInput?.value || '');
        const reviewScope = normalizeBulkReviewScope(filterReviewScope?.value || 'all');
        const matches = getWorkingRows().filter(row => {
            const ctx = row._ctx || {};
            if (errorType && toDisplay(row.Error_Type) !== errorType) return false;
            if (decisionFilter === '(blank)' && toDisplay(row.Decision)) return false;
            if (decisionFilter && decisionFilter !== '(blank)' && toDisplay(row.Decision) !== decisionFilter) return false;
            if (reviewScope === 'translation_only' && isSourceDerivedMissingMapping(row)) return false;
            if (reviewScope === 'missing_only' && !isSourceDerivedMissingMapping(row)) return false;
            if (outcomesContains && !normalize(ctx.outcomesName).includes(outcomesContains)) return false;
            if (wsuContains && !normalize(ctx.wsuName).includes(wsuContains)) return false;
            return true;
        });
        matches.sort((a, b) => {
            const aName = normalize(a._ctx?.outcomesName);
            const bName = normalize(b._ctx?.outcomesName);
            if (aName !== bName) return aName.localeCompare(bName);
            return normalize(a.translate_input).localeCompare(normalize(b.translate_input));
        });
        return matches;
    };
    const refreshCounters = () => {
        const allRows = getWorkingRows();
        const validIds = new Set(allRows.map(row => String(row.Review_Row_ID || '')));
        selectedRowIds = new Set([...selectedRowIds].filter(id => validIds.has(id)));
        const reviewedCount = allRows.reduce((count, row) => (isRowReviewed(row) ? count + 1 : count), 0);
        const pageSize = getPageSize();
        const totalFiltered = filteredRowsCache.length;
        const totalPages = clampCurrentPage();
        const startIndex = totalFiltered ? ((currentPage - 1) * pageSize) + 1 : 0;
        const endIndex = totalFiltered ? Math.min(totalFiltered, (currentPage * pageSize)) : 0;
        if (rowCountEl) rowCountEl.textContent = String(allRows.length);
        if (filterCountEl) filterCountEl.textContent = String(totalFiltered);
        pageRangeControls.forEach(control => {
            control.textContent = `${startIndex}-${endIndex}`;
        });
        pageStatusControls.forEach(control => {
            control.textContent = `Page ${currentPage} of ${totalPages}`;
        });
        pagePrevControls.forEach(control => {
            control.disabled = currentPage <= 1;
        });
        pageNextControls.forEach(control => {
            control.disabled = currentPage >= totalPages;
        });
        if (selectedCountEl) selectedCountEl.textContent = String(selectedRowIds.size);
        if (reviewedCountEl) reviewedCountEl.textContent = String(reviewedCount);
    };
    const renderBulkEditTable = () => {
        filteredRowsCache = getFilteredRows();
        refreshBulkCandidateApplyOptions();
        const pageSize = getPageSize();
        const totalPages = Math.max(1, Math.ceil(filteredRowsCache.length / pageSize));
        if (currentPage > totalPages) currentPage = totalPages;
        if (currentPage < 1) currentPage = 1;
        const startIndex = (currentPage - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const visibleRows = filteredRowsCache.slice(startIndex, endIndex);
        const wsuKeyLookup = buildWsuKeyLookup();
        tbody.innerHTML = visibleRows.map(row => {
            const rid = String(row.Review_Row_ID || '');
            const ridAttr = encodeRowId(rid);
            const ctx = row._ctx || {};
            const outcomesCtx = [ctx.outcomesName, ctx.outcomesState, ctx.outcomesCountry].filter(Boolean).join(' | ');
            const wsuCtx = [ctx.wsuName, ctx.wsuState, ctx.wsuCountry].filter(Boolean).join(' | ');
            const candidateOptions = (row._candidates || []);
            const selectedCandidate = candidateOptions.find(
                c => toDisplay(c.candidateId || '') === toDisplay(row.Selected_Candidate_ID)
            );
            const manualKey = toDisplay(row.Manual_Suggested_Key);
            const currentOutputKey = toDisplay(row.translate_output);
            const currentWsuHtml = `
                <div class="text-xs">${escapeHtml(currentOutputKey || '(blank)')}</div>
                <div class="text-xs text-gray-700">${escapeHtml(wsuCtx || '(blank)')}</div>
            `;
            let effectivePreview = null;
            let effectiveBadge = '';
            if (manualKey) {
                const fromLookup = wsuKeyLookup.get(normalizeKey(manualKey));
                effectivePreview = fromLookup || {
                    key: manualKey,
                    name: '',
                    city: '',
                    state: '',
                    country: ''
                };
                effectiveBadge = 'Manual override active';
            } else if (selectedCandidate) {
                effectivePreview = {
                    key: toDisplay(selectedCandidate.key || ''),
                    name: toDisplay(selectedCandidate.name || ''),
                    city: toDisplay(selectedCandidate.city || ''),
                    state: toDisplay(selectedCandidate.state || ''),
                    country: toDisplay(selectedCandidate.country || '')
                };
                effectiveBadge = `Selected ${toDisplay(selectedCandidate.candidateId || '')}`;
            }
            const candidateOptionsHtml = [
                optionHtml('', candidateOptions.length ? '(none)' : '(no location-valid suggestions)', toDisplay(row.Selected_Candidate_ID) === ''),
                ...candidateOptions.map(c => optionHtml(
                    toDisplay(c.candidateId || ''),
                    candidateLabel(c),
                    toDisplay(c.candidateId || '') === toDisplay(row.Selected_Candidate_ID)
                ))
            ].join('');
            const manualActive = toDisplay(row.Manual_Suggested_Key) ? '<div class="text-[11px] text-orange-700 mt-1">Manual override active</div>' : '';
            return `
                <tr class="border-b align-top">
                    <td class="py-1 px-2">
                        <input type="checkbox" class="bulk-row-select" data-rid="${ridAttr}" ${selectedRowIds.has(rid) ? 'checked' : ''}>
                    </td>
                    <td class="py-1 px-2">
                        <input type="checkbox" class="bulk-row-reviewed" data-rid="${ridAttr}" ${isRowReviewed(row) ? 'checked' : ''}>
                    </td>
                    <td class="py-1 px-2">${escapeHtml(toDisplay(row.Error_Type))}</td>
                    <td class="py-1 px-2">${escapeHtml(toDisplay(row.Error_Subtype))}</td>
                    <td class="py-1 px-2 break-words">${escapeHtml(outcomesCtx || '(blank)')}</td>
                    <td class="py-1 px-2 break-words">${currentWsuHtml}</td>
                    <td class="py-1 px-2 break-words">${formatPreview(effectivePreview, effectiveBadge)}</td>
                    <td class="py-1 px-2 font-mono text-xs">${escapeHtml(toDisplay(row.translate_input))}</td>
                    <td class="py-1 px-2 font-mono text-xs">${escapeHtml(toDisplay(row.translate_output))}</td>
                    <td class="py-1 px-2">
                        <select class="bulk-row-decision border border-gray-300 rounded px-1 py-1 text-xs w-full" data-rid="${ridAttr}">
                            ${decisionOptionsHtml(toDisplay(row.Decision))}
                        </select>
                    </td>
                    <td class="py-1 px-2">
                        <select class="bulk-row-reason border border-gray-300 rounded px-1 py-1 text-xs w-full" data-rid="${ridAttr}">
                            ${reasonOptionsHtml(toDisplay(row.Reason_Code))}
                        </select>
                    </td>
                    <td class="py-1 px-2">
                        <select class="bulk-row-candidate border border-gray-300 rounded px-1 py-1 text-xs w-full" data-rid="${ridAttr}">
                            ${candidateOptionsHtml}
                        </select>
                        ${manualActive}
                    </td>
                    <td class="py-1 px-2">
                        <input type="text" class="bulk-row-manual border border-gray-300 rounded px-1 py-1 text-xs w-full" data-rid="${ridAttr}" value="${escapeHtml(toDisplay(row.Manual_Suggested_Key))}">
                    </td>
                </tr>
            `;
        }).join('');
        if (filteredRowsCache.length > pageSize) {
            const shownStart = filteredRowsCache.length ? (startIndex + 1) : 0;
            const shownEnd = Math.min(filteredRowsCache.length, endIndex);
            tbody.innerHTML += `<tr><td colspan="13" class="py-1 px-2 text-gray-500 text-xs">Showing rows ${shownStart}-${shownEnd} of ${filteredRowsCache.length} filtered rows.</td></tr>`;
        }
        refreshCounters();
        refreshErrorPresentation();
    };
    const getApplyTargets = () => {
        const applyScope = toDisplay(applyScopeSelect?.value || 'filtered');
        const scopeRows = applyScope === 'page' ? getCurrentPageRows() : filteredRowsCache;
        const selectedOnly = Boolean(applySelectedOnly?.checked);
        if (!selectedOnly) return scopeRows;
        return scopeRows.filter(row => selectedRowIds.has(String(row.Review_Row_ID || '')));
    };
    const hydrateQueueForPanel = () => {
        getWorkingRows().forEach(attachRowContext);
        updateUnresolvedErrorsToggleState();
        currentPage = 1;
        const types = [...new Set(getWorkingRows().map(r => toDisplay(r.Error_Type)).filter(Boolean))].sort();
        if (filterSelect) {
            filterSelect.innerHTML = '<option value="">All</option>' + types.map(t => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('');
        }
        if (uploadedSessionRows && uploadedSessionRows.length && !uploadedSessionApplied) {
            uploadedSessionApplied = applySessionDataToActionQueue(uploadedSessionRows, { sourceLabel: 'Upload' });
        }
    };
    const loadActionQueueRows = async () => {
        if (actionQueueRowsCache) {
            hydrateQueueForPanel();
            renderBulkEditTable();
            return;
        }
        toggleBtn.disabled = true;
        toggleBtn.textContent = 'Loading...';
        if (loadProgressWrap && loadProgressText && loadProgressPercent && loadProgressBar) {
            loadProgressWrap.classList.remove('hidden');
            loadProgressText.textContent = 'Loading review rows...';
            loadProgressPercent.textContent = '0%';
            loadProgressBar.style.width = '0%';
        }
        try {
            if (actionQueuePrefetchPromise) {
                if (loadProgressWrap && loadProgressText && loadProgressPercent && loadProgressBar) {
                    loadProgressWrap.classList.remove('hidden');
                    loadProgressText.textContent = 'Finalizing background queue prefetch...';
                    loadProgressPercent.textContent = '90%';
                    loadProgressBar.style.width = '90%';
                }
                await actionQueuePrefetchPromise;
            }
            if (!actionQueueRowsCache) {
                const result = await runExportWorkerTask(
                    'get_action_queue',
                    buildActionQueuePayload(),
                    (stage, processed, total) => {
                        if (!loadProgressWrap || !loadProgressText || !loadProgressPercent || !loadProgressBar) return;
                        const percent = total ? Math.round((processed / total) * 100) : 0;
                        loadProgressText.textContent = stage || 'Loading review rows...';
                        loadProgressPercent.textContent = `${percent}%`;
                        loadProgressBar.style.width = `${percent}%`;
                    }
                );
                actionQueueRowsCache = cloneRows(result?.actionQueueRows || []);
                preEditedActionQueueRows = cloneRows(actionQueueRowsCache);
            }
            hydrateQueueForPanel();
            renderBulkEditTable();
        } catch (err) {
            alert(`Error loading action queue: ${err.message}`);
        } finally {
            toggleBtn.disabled = false;
            toggleBtn.textContent = 'Bulk edit before export';
            if (loadProgressWrap) {
                setTimeout(() => loadProgressWrap.classList.add('hidden'), 500);
            }
        }
    };

    const openBulkEditorPanel = async (markOneTimeOpen = false) => {
        panel.classList.remove('hidden');
        if (markOneTimeOpen && !bulkEditorOpenedOnce) {
            bulkEditorOpenedOnce = true;
            syncToggleButtonVisibility();
        }
        await loadActionQueueRows();
    };

    toggleBtn.addEventListener('click', async function() {
        await openBulkEditorPanel(true);
    });

    filterSelect?.addEventListener('change', () => {
        currentPage = 1;
        renderBulkEditTable();
    });
    filterDecisionSelect?.addEventListener('change', () => {
        currentPage = 1;
        renderBulkEditTable();
    });
    filterOutcomesNameInput?.addEventListener('input', () => {
        currentPage = 1;
        renderBulkEditTable();
    });
    filterWsuNameInput?.addEventListener('input', () => {
        currentPage = 1;
        renderBulkEditTable();
    });
    filterReviewScope?.addEventListener('change', () => {
        currentPage = 1;
        renderBulkEditTable();
    });
    pageSizeControls.forEach(control => {
        control.addEventListener('change', () => {
            setPageSize(control.value || currentPageSize);
            currentPage = 1;
            renderBulkEditTable();
        });
    });
    pagePrevControls.forEach(control => {
        control.addEventListener('click', () => {
            if (currentPage <= 1) return;
            currentPage -= 1;
            renderBulkEditTable();
        });
    });
    pageNextControls.forEach(control => {
        control.addEventListener('click', () => {
            const totalPages = getTotalPages();
            if (currentPage >= totalPages) return;
            currentPage += 1;
            renderBulkEditTable();
        });
    });
    quickChipButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            const chipValue = toDisplay(btn.getAttribute('data-outcomes-name'));
            if (filterOutcomesNameInput) filterOutcomesNameInput.value = chipValue;
            currentPage = 1;
            renderBulkEditTable();
        });
    });
    quickClearBtn?.addEventListener('click', function() {
        if (filterOutcomesNameInput) filterOutcomesNameInput.value = '';
        if (filterWsuNameInput) filterWsuNameInput.value = '';
        if (filterDecisionSelect) filterDecisionSelect.value = '';
        if (filterSelect) filterSelect.value = '';
        if (filterReviewScope) filterReviewScope.value = 'all';
        currentPage = 1;
        renderBulkEditTable();
    });

    tbody.addEventListener('change', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        const rid = decodeRowId(target.getAttribute('data-rid'));
        if (!rid) return;
        const row = findRowById(rid);
        if (!row) return;
        if (target.classList.contains('bulk-row-select')) {
            if (target.checked) selectedRowIds.add(rid);
            else selectedRowIds.delete(rid);
            refreshCounters();
            return;
        }
        if (target.classList.contains('bulk-row-reviewed')) {
            row.Reviewed = Boolean(target.checked);
            preEditedActionQueueRows = getWorkingRows();
            refreshCounters();
            refreshErrorPresentation();
            return;
        }
        if (target.classList.contains('bulk-row-decision')) {
            row.Decision = toDisplay(target.value);
        } else if (target.classList.contains('bulk-row-reason')) {
            row.Reason_Code = toDisplay(target.value);
        } else if (target.classList.contains('bulk-row-candidate')) {
            row.Selected_Candidate_ID = toDisplay(target.value);
            if (row.Selected_Candidate_ID) row.Manual_Suggested_Key = '';
        }
        preEditedActionQueueRows = getWorkingRows();
        renderBulkEditTable();
    });
    tbody.addEventListener('input', (event) => {
        const target = event.target;
        if (!(target instanceof HTMLElement)) return;
        if (!target.classList.contains('bulk-row-manual')) return;
        const rid = decodeRowId(target.getAttribute('data-rid'));
        if (!rid) return;
        const row = findRowById(rid);
        if (!row) return;
        row.Manual_Suggested_Key = toDisplay(target.value);
        if (row.Manual_Suggested_Key) row.Selected_Candidate_ID = '';
        preEditedActionQueueRows = getWorkingRows();
    });

    applyBtn?.addEventListener('click', function() {
        const decision = toDisplay(applyDecisionSelect?.value);
        const reasonCode = toDisplay(applyReasonCodeSelect?.value);
        const candidateId = toDisplay(applyCandidateIdSelect?.value);
        const manualKey = (applyManualInput?.value || '').trim();
        if (!decision && !reasonCode && !candidateId && !manualKey) {
            alert('Choose at least one bulk field to apply.');
            return;
        }
        const targets = getApplyTargets();
        if (!targets.length) {
            alert('No rows match current filters/selection.');
            return;
        }
        const applyScope = toDisplay(applyScopeSelect?.value || 'filtered');
        const selectedOnly = Boolean(applySelectedOnly?.checked);
        const scopeLabel = selectedOnly
            ? (applyScope === 'page' ? 'selected rows on current page' : 'selected rows in all filtered rows')
            : (applyScope === 'page' ? 'all rows on current page' : 'all filtered rows');
        if (!confirm(`Apply changes to ${targets.length} row${targets.length === 1 ? '' : 's'} (${scopeLabel})?`)) {
            return;
        }
        targets.forEach(row => {
            if (decision) row.Decision = decision;
            if (reasonCode) row.Reason_Code = reasonCode;
            if (candidateId) {
                row.Selected_Candidate_ID = candidateId;
                row.Manual_Suggested_Key = '';
            }
            if (manualKey) {
                row.Manual_Suggested_Key = manualKey;
                row.Selected_Candidate_ID = '';
            }
        });
        preEditedActionQueueRows = getWorkingRows();
        renderBulkEditTable();
    });
    clearSelectedBtn?.addEventListener('click', function() {
        const rows = getWorkingRows().filter(row => selectedRowIds.has(String(row.Review_Row_ID || '')));
        if (!rows.length) {
            alert('No selected rows to clear.');
            return;
        }
        if (!confirm(`Clear Decision, Reason, Candidate, and Manual Key for ${rows.length} selected rows?`)) return;
        rows.forEach(row => {
            row.Decision = '';
            row.Reason_Code = '';
            row.Selected_Candidate_ID = '';
            row.Manual_Suggested_Key = '';
        });
        preEditedActionQueueRows = getWorkingRows();
        renderBulkEditTable();
    });
    selectFilteredBtn?.addEventListener('click', function() {
        filteredRowsCache.forEach(row => selectedRowIds.add(String(row.Review_Row_ID || '')));
        refreshCounters();
        renderBulkEditTable();
    });
    deselectFilteredBtn?.addEventListener('click', function() {
        filteredRowsCache.forEach(row => selectedRowIds.delete(String(row.Review_Row_ID || '')));
        refreshCounters();
        renderBulkEditTable();
    });
    markReviewedBtn?.addEventListener('click', function() {
        if (!filteredRowsCache.length) {
            alert('No filtered rows to mark reviewed.');
            return;
        }
        filteredRowsCache.forEach(row => {
            row.Reviewed = true;
        });
        preEditedActionQueueRows = getWorkingRows();
        renderBulkEditTable();
    });
    clearReviewedBtn?.addEventListener('click', function() {
        if (!filteredRowsCache.length) {
            alert('No filtered rows to clear reviewed state.');
            return;
        }
        filteredRowsCache.forEach(row => {
            row.Reviewed = false;
        });
        preEditedActionQueueRows = getWorkingRows();
        renderBulkEditTable();
    });
    saveSessionBtn?.addEventListener('click', function() {
        const rows = getWorkingRows();
        if (!rows.length) {
            alert('Nothing to save yet.');
            return;
        }
        const payload = {
            version: 1,
            savedAt: new Date().toISOString(),
            rowCount: rows.length,
            rows: rows.map(row => ({
                Review_Row_ID: row.Review_Row_ID || '',
                Decision: row.Decision || '',
                Reason_Code: row.Reason_Code || '',
                Selected_Candidate_ID: row.Selected_Candidate_ID || '',
                Manual_Suggested_Key: row.Manual_Suggested_Key || '',
                Reviewed: isRowReviewed(row)
            }))
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `validate_review_session_${Date.now()}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    });
    loadSessionInput?.addEventListener('change', async function(event) {
        const input = event.target;
        const file = input?.files?.[0];
        if (!file) return;
        try {
            await loadActionQueueRows();
            const text = await file.text();
            const parsed = JSON.parse(text);
            const rows = parseSessionRowsPayload(parsed);
            applySessionDataToActionQueue(rows);
            renderBulkEditTable();
        } catch (err) {
            alert(`Error loading session: ${err.message}`);
        } finally {
            input.value = '';
        }
    });

    const refreshFromExternalLoad = () => {
        if (!panel.classList.contains('hidden') && getWorkingRows().length) {
            getWorkingRows().forEach(attachRowContext);
            renderBulkEditTable();
        }
    };
    document.addEventListener('session-upload-applied', refreshFromExternalLoad);
    document.addEventListener('prior-upload-applied', refreshFromExternalLoad);
    document.addEventListener('validation-results-ready', async () => {
        syncToggleButtonVisibility();
        if (!bulkEditorOpenedOnce) return;
        await openBulkEditorPanel(false);
    });
}

function setupResetButton() {
    const resetBtn = document.getElementById('reset-btn');
    resetBtn.addEventListener('click', function() {
        if (confirm('Are you sure you want to start over? This will clear all uploaded files and results.')) {
            // Keep this explicit so a future soft-reset path can share the same state reset.
            bulkEditorOpenedOnce = false;
            location.reload();
        }
    });
}
