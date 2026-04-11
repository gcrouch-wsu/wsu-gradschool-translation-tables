'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const workerPath = path.join(__dirname, 'export-worker.js');
const workerCode = fs.readFileSync(workerPath, 'utf8');

let failures = 0;

async function runCheck(name, fn) {
    try {
        await fn();
        console.log(`[PASS] ${name}`);
    } catch (error) {
        failures += 1;
        console.error(`[FAIL] ${name}: ${error.message}`);
    }
}

function columnLetterToIndex(letters) {
    return String(letters || '')
        .toUpperCase()
        .split('')
        .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
}

class FakeCell {
    constructor() {
        this.value = '';
        this.font = undefined;
        this.fill = undefined;
        this.border = undefined;
        this.numFmt = undefined;
        this.protection = undefined;
    }
}

class FakeRow {
    constructor(rowNumber) {
        this.rowNumber = rowNumber;
        this._cells = new Map();
    }

    getCell(index) {
        const col = Number(index);
        if (!this._cells.has(col)) {
            this._cells.set(col, new FakeCell());
        }
        return this._cells.get(col);
    }

    eachCell(callback) {
        const cols = Array.from(this._cells.keys()).sort((a, b) => a - b);
        cols.forEach(colNumber => {
            callback(this._cells.get(colNumber), colNumber);
        });
    }
}

class FakeWorksheet {
    constructor(name) {
        this.name = name;
        this._rows = new Map();
        this._rowCount = 0;
        this._columns = [];
        this.state = 'visible';
        this.views = [];
        this.autoFilter = undefined;
        this.dataValidations = {
            items: [],
            add: (ref, config) => {
                this.dataValidations.items.push({ ref, config });
            }
        };
        this.conditionalFormatting = [];
    }

    get rowCount() {
        return this._rowCount;
    }

    get columns() {
        return this._columns;
    }

    set columns(columns) {
        this._columns = Array.isArray(columns) ? columns : [];
    }

    _ensureRow(rowNumber) {
        const rowNum = Number(rowNumber);
        if (!this._rows.has(rowNum)) {
            this._rows.set(rowNum, new FakeRow(rowNum));
        }
        if (rowNum > this._rowCount) {
            this._rowCount = rowNum;
        }
        return this._rows.get(rowNum);
    }

    addRow(values) {
        const rowNumber = this._rowCount + 1;
        const row = this._ensureRow(rowNumber);
        if (Array.isArray(values)) {
            values.forEach((value, idx) => {
                row.getCell(idx + 1).value = value;
            });
        }
        return row;
    }

    getRow(rowNumber) {
        return this._ensureRow(rowNumber);
    }

    getColumn(index) {
        const colNumber = Number(index);
        if (!this._columns[colNumber - 1]) {
            this._columns[colNumber - 1] = {};
        }
        return this._columns[colNumber - 1];
    }

    getCell(refOrRow, colNumber) {
        if (typeof refOrRow === 'string') {
            const match = /^([A-Za-z]+)(\d+)$/.exec(refOrRow);
            if (!match) {
                throw new Error(`Unsupported cell ref: ${refOrRow}`);
            }
            const col = columnLetterToIndex(match[1]);
            const row = Number(match[2]);
            return this._ensureRow(row).getCell(col);
        }
        return this._ensureRow(refOrRow).getCell(colNumber);
    }

    addConditionalFormatting(config) {
        this.conditionalFormatting.push(config);
    }

    async protect() {
        return undefined;
    }
}

function createHarness() {
    let lastWorkbook = null;
    const progressMessages = [];

    class FakeWorkbook {
        constructor() {
            this._worksheets = [];
            lastWorkbook = this;
            this.xlsx = {
                writeBuffer: async () => new ArrayBuffer(16)
            };
        }

        addWorksheet(name) {
            const sheet = new FakeWorksheet(name);
            this._worksheets.push(sheet);
            return sheet;
        }

        getWorksheet(name) {
            return this._worksheets.find(sheet => sheet.name === name);
        }
    }

    const context = {
        console,
        Buffer,
        ArrayBuffer,
        setTimeout,
        clearTimeout,
        ExcelJS: { Workbook: FakeWorkbook },
        self: {
            postMessage: (message) => {
                progressMessages.push(message);
            }
        }
    };

    context.importScripts = (...scripts) => {
        scripts.forEach(scriptRef => {
            if (!scriptRef) return;
            if (scriptRef.includes('exceljs')) {
                return;
            }
            const scriptPath = path.isAbsolute(scriptRef)
                ? scriptRef
                : path.join(__dirname, scriptRef);
            const code = fs.readFileSync(scriptPath, 'utf8');
            vm.runInContext(code, vmContext, { filename: scriptPath });
        });
    };

    const vmContext = vm.createContext(context);
    vm.runInContext(workerCode, vmContext, { filename: workerPath });

    if (typeof vmContext.buildValidationExport !== 'function') {
        throw new Error('buildValidationExport was not loaded from export-worker.js');
    }
    if (typeof vmContext.buildGenerationExport !== 'function') {
        throw new Error('buildGenerationExport was not loaded from export-worker.js');
    }

    return {
        buildValidationExport: (payload) => vmContext.buildValidationExport(payload),
        buildGenerationExport: (payload) => vmContext.buildGenerationExport(payload),
        getLastWorkbook: () => lastWorkbook,
        getProgressMessages: () => progressMessages
    };
}

function getRowValues(sheet, rowNumber, upToColumn) {
    const row = sheet.getRow(rowNumber);
    const values = [];
    for (let col = 1; col <= upToColumn; col += 1) {
        values.push(row.getCell(col).value);
    }
    return values;
}

function findHeaderIndex(sheet, headerText, scanColumns = 80) {
    const headers = getRowValues(sheet, 1, scanColumns).map(v => String(v || ''));
    return headers.indexOf(headerText) + 1;
}

function assertExportResult(result) {
    assert.ok(result, 'Expected result object');
    assert.equal(typeof result.filename, 'string');
    assert.ok(result.filename.length > 0, 'Expected non-empty filename');
    assert.ok(result.buffer instanceof ArrayBuffer, 'Expected ArrayBuffer buffer');
    assert.ok(result.buffer.byteLength > 0, 'Expected non-empty buffer');
}

async function run() {
    await runCheck('buildValidationExport handles empty payload object', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({});
        assertExportResult(result);
        assert.ok(harness.getProgressMessages().length > 0, 'Expected progress messages');
    });

    await runCheck('buildValidationExport handles explicit empty/minimal payload', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [],
            selectedCols: {},
            context: {}
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const finalSheet = workbook.getWorksheet('Final_Translation_Table');
        const filterCell = finalSheet?.getRow(2)?.getCell(1)?.value;
        assert.ok(filterCell?.formula, 'Empty payload should still produce FILTER formula');
        assert.ok(!/:\$[A-Z]+\$1\b/.test(filterCell.formula), 'FILTER include range must not end at row 1 (stagingLastRow >= 2 guard)');
    });

    await runCheck('buildValidationExport handles missing selectedCols arrays', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [],
            selectedCols: {},
            context: {
                loadedData: {
                    outcomes: [{ some_key: 'A1', school: 'Alpha' }],
                    translate: [],
                    wsu_org: [{ some_key: 'B1', school: 'Beta' }]
                },
                keyConfig: {
                    outcomes: 'some_key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'some_key'
                }
            }
        });
        assertExportResult(result);
    });

    await runCheck('buildValidationExport sanitizes object-valued cells', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Valid',
                    translate_input: 'x',
                    translate_output: 'y',
                    outcomes_school: { foo: 1 },
                    wsu_school: 'Y School'
                }
            ],
            selectedCols: {
                outcomes: ['school'],
                wsu_org: ['school']
            },
            context: {
                loadedData: { outcomes: [], translate: [], wsu_org: [] },
                keyConfig: {
                    outcomes: 'outcomes_key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'wsu_key'
                },
                keyLabels: {
                    outcomes: 'outcomes_key',
                    wsu: 'wsu_key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                }
            }
        });

        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const validSheet = workbook.getWorksheet('Valid_Mappings');
        assert.ok(validSheet, 'Expected Valid_Mappings worksheet');
        const headerValues = getRowValues(validSheet, 1, 12).map(v => String(v || ''));
        const outcomesSchoolIndex = headerValues.indexOf('Outcomes Name');
        assert.ok(outcomesSchoolIndex >= 0, 'Expected Outcomes Name column header');
        const dataCell = validSheet.getRow(2).getCell(outcomesSchoolIndex + 1).value;
        assert.equal(typeof dataCell, 'string');
        assert.equal(dataCell, '[object Object]');
    });

    await runCheck('buildValidationExport handles normal small payload', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Valid',
                    translate_input: 'IN-1',
                    translate_output: 'OUT-1',
                    outcomes_school: 'Alpha University',
                    wsu_school: 'Alpha University'
                }
            ],
            selectedCols: {
                outcomes: ['school', 'key'],
                wsu_org: ['school', 'key']
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-1', school: 'Alpha University' }],
                    translate: [{ translate_input: 'IN-1', translate_output: 'OUT-1' }],
                    wsu_org: [{ key: 'OUT-1', school: 'Alpha University' }]
                },
                keyConfig: {
                    outcomes: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'key'
                },
                keyLabels: {
                    outcomes: 'key',
                    wsu: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                },
                columnRoles: {
                    outcomes: { school: 'School' },
                    wsu_org: { school: 'School' }
                }
            },
            options: {
                fileName: 'Validation_Test.xlsx'
            }
        });
        assertExportResult(result);
        assert.equal(result.filename, 'Validation_Test.xlsx');
        const workbook = harness.getLastWorkbook();
        assert.equal(
            workbook.calcProperties && workbook.calcProperties.fullCalcOnLoad,
            true,
            'Validation workbook should force full recalculation on open'
        );
        assert.ok(workbook.getWorksheet('Review_Workbench'), 'Expected Review_Workbench worksheet');
        assert.ok(workbook.getWorksheet('QA_Checks_Validate'), 'Expected QA_Checks_Validate worksheet');
        const finalSheet = workbook.getWorksheet('Final_Translation_Table');
        assert.ok(finalSheet, 'Expected Final_Translation_Table worksheet');
        const stagingSheet = workbook.getWorksheet('Final_Staging');
        assert.ok(stagingSheet, 'Expected Final_Staging worksheet');
        const translateInputCol = findHeaderIndex(finalSheet, 'Translate Input');
        assert.ok(translateInputCol > 0, 'Final table should include Translate Input');
        const stagingTranslateInputCol = findHeaderIndex(stagingSheet, 'Translate Input');
        assert.ok(stagingTranslateInputCol > 0, 'Final_Staging should include Translate Input');
        const stagingInputCell = stagingSheet.getRow(2).getCell(stagingTranslateInputCol).value;
        assert.equal(stagingInputCell, 'IN-1', 'Auto-approved rows should be written as plain values in Final_Staging');
        const filterCell = finalSheet.getRow(2).getCell(1).value;
        assert.ok(filterCell && filterCell.formula, 'Final table A2 should have FILTER formula');
        assert.ok(!filterCell.formula.startsWith('='), 'FILTER formula must not have leading = (OOXML compliance)');
        assert.ok(filterCell.formula.includes('_xlfn._xlws.FILTER'), 'FILTER must use future-function namespace _xlfn._xlws.FILTER');
        assert.strictEqual(filterCell.shareType, 'array', 'FILTER must have shareType array');
        assert.ok(filterCell.ref && /^A2:[A-Z]+\d+$/.test(filterCell.ref), 'FILTER ref must be bounded range A2:ColN');
        assert.ok(filterCell.formula.includes('Final_Staging'), 'FILTER formula should reference Final_Staging');
        assert.ok(/\$[A-Z]+\$\d+:\$[A-Z]+\$\d+/.test(filterCell.formula), 'FILTER must use absolute ref $Col$2:$Col$N for include range');
    });

    await runCheck('buildValidationExport keeps review-to-final approval flow and output-side duplicate suggestions', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Duplicate_Target',
                    Duplicate_Group: 'G-1',
                    translate_input: 'IN-ALPHA',
                    translate_output: 'OUT-LEGACY',
                    outcomes_school: 'Alpha Campus',
                    wsu_school: 'Legacy Org'
                }
            ],
            selectedCols: {
                outcomes: ['school'],
                wsu_org: ['school']
            },
            options: {
                includeSuggestions: true,
                nameCompareConfig: {
                    enabled: true,
                    outcomes: 'school',
                    wsu: 'school',
                    threshold: 0.8
                }
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-ALPHA', school: 'Alpha Campus' }],
                    translate: [{ translate_input: 'IN-ALPHA', translate_output: 'OUT-LEGACY' }],
                    wsu_org: [
                        { key: 'OUT-LEGACY', school: 'Legacy Org' },
                        { key: 'OUT-BETTER', school: 'Alpha Campus' }
                    ]
                },
                keyConfig: {
                    outcomes: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'key'
                },
                keyLabels: {
                    outcomes: 'key',
                    wsu: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                },
                columnRoles: {
                    outcomes: { school: 'School' },
                    wsu_org: { school: 'School' }
                }
            }
        });
        assertExportResult(result);

        const workbook = harness.getLastWorkbook();
        const actionQueue = workbook.getWorksheet('Action_Queue');
        const approvedMappings = workbook.getWorksheet('Approved_Mappings');
        const errorsSheet = workbook.getWorksheet('Errors_in_Translate');
        const ambiguousOutputSheet = workbook.getWorksheet('Output_Not_Found_Ambiguous');
        const noReplacementOutputSheet = workbook.getWorksheet('Output_Not_Found_No_Replacement');
        const oneToMany = workbook.getWorksheet('One_to_Many');
        const missingMappings = workbook.getWorksheet('Missing_Mappings');
        const highConfidence = workbook.getWorksheet('High_Confidence_Matches');
        const validMappings = workbook.getWorksheet('Valid_Mappings');
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const finalSheet = workbook.getWorksheet('Final_Translation_Table');
        assert.ok(actionQueue, 'Expected Action_Queue worksheet');
        assert.ok(approvedMappings, 'Expected Approved_Mappings worksheet');
        assert.ok(errorsSheet, 'Expected Errors_in_Translate worksheet');
        assert.ok(oneToMany, 'Expected One_to_Many worksheet');
        assert.ok(missingMappings, 'Expected Missing_Mappings worksheet');
        assert.ok(highConfidence, 'Expected High_Confidence_Matches worksheet');
        assert.ok(validMappings, 'Expected Valid_Mappings worksheet');
        assert.ok(reviewSheet, 'Expected Review_Workbench worksheet');
        assert.ok(finalSheet, 'Expected Final_Translation_Table worksheet');

        assert.equal(actionQueue.state, 'hidden', 'Action_Queue should be hidden in exported workbook');
        assert.equal(approvedMappings.state, 'hidden', 'Approved_Mappings should be hidden in exported workbook');
        assert.equal(errorsSheet.state, 'hidden', 'Errors_in_Translate should be hidden in exported workbook');
        if (ambiguousOutputSheet) {
            assert.equal(ambiguousOutputSheet.state, 'hidden', 'Output_Not_Found_Ambiguous should be hidden in exported workbook');
        }
        if (noReplacementOutputSheet) {
            assert.equal(noReplacementOutputSheet.state, 'hidden', 'Output_Not_Found_No_Replacement should be hidden in exported workbook');
        }
        assert.equal(oneToMany.state, 'hidden', 'One_to_Many should be hidden in exported workbook');
        assert.equal(missingMappings.state, 'hidden', 'Missing_Mappings should be hidden in exported workbook');
        assert.equal(highConfidence.state, 'hidden', 'High_Confidence_Matches should be hidden in exported workbook');
        assert.equal(validMappings.state, 'hidden', 'Valid_Mappings should be hidden in exported workbook');

        const reviewSheetIndex = workbook._worksheets.findIndex(sheet => sheet && sheet.name === 'Review_Workbench');
        assert.ok(reviewSheetIndex >= 0, 'Review_Workbench should exist in workbook order');
        assert.equal(workbook.views?.[0]?.activeTab, reviewSheetIndex, 'Review_Workbench should be the active tab on open');
        assert.equal(workbook.views?.[0]?.firstSheet, reviewSheetIndex, 'Review_Workbench should be the first visible tab on open');

        const reviewView = Array.isArray(reviewSheet.views) && reviewSheet.views[0] ? reviewSheet.views[0] : {};
        assert.equal(reviewView.ySplit, 1, 'Review_Workbench should freeze header row');
        assert.ok(!reviewView.xSplit, 'Review_Workbench should not freeze wide left pane columns');

        const suggestedKeyCol = findHeaderIndex(oneToMany, 'Suggested Key');
        assert.ok(suggestedKeyCol > 0, 'One_to_Many should include Suggested Key');
        const suggestedKeyValue = oneToMany.getRow(2).getCell(suggestedKeyCol).value;
        assert.equal(suggestedKeyValue, 'OUT-BETTER', 'Duplicate rows should suggest myWSU output-side key');

        const reviewFinalInputCol = findHeaderIndex(reviewSheet, 'Final Translate Input');
        const reviewPublishEligibleCol = findHeaderIndex(reviewSheet, 'Publish Eligible (1=yes)');
        assert.ok(reviewFinalInputCol > 0, 'Review sheet should include Final Translate Input');
        assert.ok(reviewPublishEligibleCol > 0, 'Review sheet should include Publish Eligible');
        const finalInputFormula = reviewSheet.getRow(2).getCell(reviewFinalInputCol).value?.formula || '';
        const publishFormula = reviewSheet.getRow(2).getCell(reviewPublishEligibleCol).value?.formula || '';
        assert.ok(finalInputFormula.includes('"Keep As-Is"') && finalInputFormula.includes('"Use Suggestion"'), 'Final input formula should use renamed decision values');
        assert.ok(publishFormula.includes('"Keep As-Is"') && publishFormula.includes('"Allow One-to-Many"'), 'Publish eligibility should use renamed decision values');

        const finalFilterCell = finalSheet.getRow(2).getCell(1).value;
        assert.ok(finalFilterCell?.formula, 'Final table A2 should have FILTER formula');
        assert.ok(!finalFilterCell.formula.startsWith('='), 'FILTER formula must not have leading = (OOXML compliance)');
        assert.ok(finalFilterCell.formula.includes('_xlfn._xlws.FILTER'), 'FILTER must use future-function namespace _xlfn._xlws.FILTER');
        assert.strictEqual(finalFilterCell.shareType, 'array', 'FILTER must have shareType array');
        assert.ok(finalFilterCell.ref && /^A2:[A-Z]+\d+$/.test(finalFilterCell.ref), 'FILTER ref must be bounded range A2:ColN');
        assert.ok(finalFilterCell.formula.includes('Final_Staging'), 'Final table should pull from Final_Staging via FILTER');
        assert.ok(/\$[A-Z]+\$\d+:\$[A-Z]+\$\d+/.test(finalFilterCell.formula), 'FILTER must use absolute ref $Col$2:$Col$N for include range');
        const stagingSheet = workbook.getWorksheet('Final_Staging');
        assert.ok(stagingSheet, 'Final_Staging should exist and feed Final_Translation_Table');
        const outcomesNameCol = findHeaderIndex(finalSheet, 'Outcomes Name');
        const myWsuNameCol = findHeaderIndex(finalSheet, 'myWSU Name');
        const translateInputCol = findHeaderIndex(finalSheet, 'Translate Input');
        const translateOutputCol = findHeaderIndex(finalSheet, 'Translate Output');
        assert.ok(outcomesNameCol > 0, 'Final table should include Outcomes Name context');
        assert.ok(myWsuNameCol > 0, 'Final table should include myWSU Name context');
        assert.ok(translateInputCol > 0 && translateOutputCol > 0, 'Final table should include Translate Input and Translate Output');
        assert.equal(
            findHeaderIndex(finalSheet, '_Approved Pick'),
            0,
            'Final table should not include AGGREGATE helper columns'
        );
        assert.ok(finalSheet.autoFilter, 'Final table should have autoFilter enabled');
        assert.equal(finalSheet.autoFilter.from, 'A1', 'Final table autoFilter should start at A1');
        assert.ok(finalSheet.autoFilter.to, 'Final table autoFilter should include all output columns');

        const qaSheet = workbook.getWorksheet('QA_Checks_Validate');
        assert.ok(qaSheet, 'Expected QA_Checks_Validate worksheet');
        const publishGateRow = [...Array(20)].map((_, i) => qaSheet.getRow(i + 1).getCell(1).value).findIndex(v => String(v || '').includes('Publish gate'));
        assert.ok(publishGateRow >= 0, 'QA sheet should have Publish gate row');
        const publishGateDetail = qaSheet.getRow(publishGateRow + 1).getCell(4).value;
        assert.ok(
            String(publishGateDetail || '').includes('Diagnostic tabs are hidden'),
            'QA publish gate detail should mention hidden diagnostic tabs'
        );
    });

    await runCheck('Fix A: Output_Not_Found_No_Replacement does not get weak fallback suggestions', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Output_Not_Found',
                    _rawErrorType: 'Output_Not_Found',
                    Error_Subtype: 'Output_Not_Found_No_Replacement',
                    _rawErrorSubtype: 'Output_Not_Found_No_Replacement',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-STALE',
                    outcomes_school: 'Alpha University',
                    wsu_school: ''
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: {
                includeSuggestions: true,
                nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.5 }
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'Alpha University' }],
                    translate: [{ translate_input: 'OUT-1', translate_output: 'WSU-STALE' }],
                    wsu_org: [
                        { key: 'WSU-OTHER', school: 'Alpha University' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const suggestedKeyCol = findHeaderIndex(reviewSheet, 'Suggested Key');
        assert.ok(suggestedKeyCol > 0, 'Review sheet should have Suggested Key column');
        const suggestedKeyCell = reviewSheet.getRow(2).getCell(suggestedKeyCol).value;
        const isFormula = suggestedKeyCell && typeof suggestedKeyCell === 'object' && suggestedKeyCell.formula;
        assert.ok(!isFormula, 'NO_REPLACEMENT row should not have formula in Suggested_Key (would allow computed suggestion)');
        const suggestedKeyVal = String(suggestedKeyCell?.result ?? suggestedKeyCell ?? '');
        assert.equal(suggestedKeyVal, '', 'NO_REPLACEMENT row should not get fallback Suggested_Key from export');
    });

    await runCheck('Fix C: Manual_Suggested_Key column exists and effective key supports manual override', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Duplicate_Target',
                    _rawErrorType: 'Duplicate_Target',
                    translate_input: 'IN-1',
                    translate_output: 'OUT-LEGACY',
                    outcomes_school: 'Campus A',
                    wsu_school: 'Legacy Org',
                    _candidates: []
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-1', school: 'Campus A' }],
                    translate: [{ translate_input: 'IN-1', translate_output: 'OUT-LEGACY' }],
                    wsu_org: [
                        { key: 'OUT-LEGACY', school: 'Legacy Org' },
                        { key: 'OUT-PARENT', school: 'Parent Org' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const manualKeyCol = findHeaderIndex(reviewSheet, 'Manual Key (override when no candidates)');
        assert.ok(manualKeyCol > 0, 'Review sheet should have Manual_Suggested_Key column');
        const finalOutputCol = findHeaderIndex(reviewSheet, 'Final Translate Output');
        assert.ok(finalOutputCol > 0, 'Review sheet should have Final_Output column');
        const finalOutputCell = reviewSheet.getRow(2).getCell(finalOutputCol).value;
        const formula = (finalOutputCell && finalOutputCell.formula) ? finalOutputCell.formula : '';
        assert.ok(formula.includes('TRIM'), 'Final_Output formula should use TRIM for Manual_Suggested_Key effective key path');
    });

    await runCheck('Fix B: B14 exemption formula excludes Duplicate_Target+Keep As-Is from duplicate output check', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-PARENT', outcomes_school: 'Campus A', wsu_school: 'Parent' },
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-B', translate_output: 'OUT-PARENT', outcomes_school: 'Campus B', wsu_school: 'Parent' }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }, { key: 'IN-B', school: 'Campus B' }],
                    translate: [
                        { translate_input: 'IN-A', translate_output: 'OUT-PARENT' },
                        { translate_input: 'IN-B', translate_output: 'OUT-PARENT' }
                    ],
                    wsu_org: [{ key: 'OUT-PARENT', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const qaSheet = workbook.getWorksheet('QA_Checks_Validate');
        const dupOutputRow = [...Array(15)].map((_, i) => qaSheet.getRow(i + 1).getCell(1).value).findIndex(v => String(v || '').includes('Duplicate final output'));
        assert.ok(dupOutputRow >= 0, 'QA sheet should have Duplicate final output check');
        const dupOutputFormula = qaSheet.getRow(dupOutputRow + 1).getCell(2).value;
        const formulaStr = (dupOutputFormula && dupOutputFormula.formula) ? dupOutputFormula.formula : String(dupOutputFormula || '');
        assert.ok(formulaStr.includes('One_to_Many') && formulaStr.includes('Duplicate_Target') && formulaStr.includes('Keep As-Is'),
            'Duplicate output formula should exempt One_to_Many+Duplicate_Target+Keep As-Is');
    });

    await runCheck('buildValidationExport dedupes Input_Not_Found vs Missing_Mapping overlap', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Input_Not_Found',
                    _rawErrorType: 'Input_Not_Found',
                    translate_input: 'BAD-KEY',
                    translate_output: 'WSU-1',
                    Suggested_Key: 'OUT-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: '',
                    wsu_school: 'UH Maui'
                }
            ],
            selectedCols: { outcomes: ['school', 'state'], wsu_org: ['school', 'state'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui', state: 'HI' }],
                    translate: [{ translate_input: 'BAD-KEY', translate_output: 'WSU-1' }],
                    wsu_org: [{ key: 'WSU-1', school: 'UH Maui', state: 'HI' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        const rowCount = reviewSheet.rowCount || 0;
        let inputNotFoundCount = 0;
        let missingMappingCount = 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Input key not found in Outcomes') inputNotFoundCount += 1;
            if (val === 'Missing_Mapping') missingMappingCount += 1;
        }
        assert.equal(missingMappingCount, 1, 'Missing_Mapping row should be kept');
        assert.equal(inputNotFoundCount, 0, 'Input_Not_Found row should be deduped (dropped)');
    });

    await runCheck('buildValidationExport keeps Input_Not_Found when no Missing_Mapping overlap', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Input_Not_Found',
                    _rawErrorType: 'Input_Not_Found',
                    translate_input: 'ORPHAN-KEY',
                    translate_output: 'WSU-ORPHAN',
                    Suggested_Key: 'OUT-ORPHAN',
                    Suggested_School: 'Hanoi University',
                    Suggestion_Score: 0.85,
                    outcomes_school: '',
                    wsu_school: 'Hanoi University'
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui' }],
                    translate: [{ translate_input: 'ORPHAN-KEY', translate_output: 'WSU-ORPHAN' }],
                    wsu_org: [
                        { key: 'WSU-1', school: 'UH Maui' },
                        { key: 'WSU-ORPHAN', school: 'Hanoi University' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        let inputNotFoundCount = 0;
        for (let r = 2; r <= (reviewSheet.rowCount || 0); r += 1) {
            if (String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '') === 'Input key not found in Outcomes') {
                inputNotFoundCount += 1;
            }
        }
        assert.equal(inputNotFoundCount, 1, 'Input_Not_Found with no overlap should remain');
    });

    await runCheck('buildValidationExport dedupes Output_Not_Found vs Missing_Mapping overlap', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Output_Not_Found',
                    _rawErrorType: 'Output_Not_Found',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-STALE',
                    Suggested_Key: 'WSU-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: 'UH Maui',
                    wsu_school: ''
                }
            ],
            selectedCols: { outcomes: ['school', 'state'], wsu_org: ['school', 'state'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui', state: 'HI' }],
                    translate: [{ translate_input: 'OUT-1', translate_output: 'WSU-STALE' }],
                    wsu_org: [{ key: 'WSU-1', school: 'UH Maui', state: 'HI' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        const rowCount = reviewSheet.rowCount || 0;
        let outputNotFoundCount = 0;
        let missingMappingCount = 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Output key not found in myWSU') outputNotFoundCount += 1;
            if (val === 'Missing_Mapping') missingMappingCount += 1;
        }
        assert.equal(missingMappingCount, 1, 'Missing_Mapping row should be kept');
        assert.equal(outputNotFoundCount, 0, 'Output_Not_Found row should be deduped (dropped)');
    });

    await runCheck('buildValidationExport keeps error row when error score exceeds Missing_Mapping score', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Input_Not_Found',
                    _rawErrorType: 'Input_Not_Found',
                    translate_input: 'BAD-KEY',
                    translate_output: 'WSU-1',
                    Suggested_Key: 'OUT-1',
                    Suggested_School: 'Alpha Beta Gamma',
                    Suggestion_Score: 0.95,
                    outcomes_school: '',
                    wsu_school: 'Alpha Beta Gamma'
                }
            ],
            selectedCols: { outcomes: ['school', 'state'], wsu_org: ['school', 'state'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.75 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'Alpha Beta Gamma', state: 'HI' }],
                    translate: [{ translate_input: 'BAD-KEY', translate_output: 'WSU-1' }],
                    wsu_org: [{ key: 'WSU-1', school: 'Alpha Beta', state: 'HI' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        const rowCount = reviewSheet.rowCount || 0;
        let inputNotFoundCount = 0;
        let missingMappingCount = 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Input key not found in Outcomes') inputNotFoundCount += 1;
            if (val === 'Missing_Mapping') missingMappingCount += 1;
        }
        assert.equal(inputNotFoundCount, 1, 'Input_Not_Found should be kept when error score > Missing_Mapping score');
        assert.equal(missingMappingCount, 0, 'Missing_Mapping should be dropped when error has stronger suggestion');
    });

    await runCheck('buildValidationExport dedupes Output_Not_Found vs Duplicate_Target when same final pair', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Output_Not_Found',
                    _rawErrorType: 'Output_Not_Found',
                    Error_Subtype: 'Output_Not_Found_Likely_Stale_Key',
                    _rawErrorSubtype: 'Output_Not_Found_Likely_Stale_Key',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-STALE',
                    Suggested_Key: 'WSU-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: 'UH Maui',
                    wsu_school: ''
                },
                {
                    Error_Type: 'Duplicate_Target',
                    _rawErrorType: 'Duplicate_Target',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-OLD',
                    Suggested_Key: 'WSU-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.85,
                    Duplicate_Group: 'G-1',
                    outcomes_school: 'UH Maui',
                    wsu_school: 'Legacy Org'
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui' }],
                    translate: [
                        { translate_input: 'OUT-1', translate_output: 'WSU-STALE' },
                        { translate_input: 'OUT-1', translate_output: 'WSU-OLD' }
                    ],
                    wsu_org: [
                        { key: 'WSU-1', school: 'UH Maui' },
                        { key: 'WSU-OLD', school: 'Legacy Org' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        const rowCount = reviewSheet.rowCount || 0;
        let outputNotFoundCount = 0;
        let duplicateTargetCount = 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Output key not found in myWSU') outputNotFoundCount += 1;
            if (val === 'Duplicate_Target') duplicateTargetCount += 1;
        }
        assert.equal(outputNotFoundCount, 1, 'Output_Not_Found row should be kept');
        assert.equal(duplicateTargetCount, 0, 'Duplicate_Target row should be deduped when same final pair as Output_Not_Found');
    });

    await runCheck('buildValidationExport keeps both Input_Not_Found and Duplicate_Source (regression: no incorrect dedupe)', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Input_Not_Found',
                    _rawErrorType: 'Input_Not_Found',
                    translate_input: 'BAD-KEY',
                    translate_output: 'WSU-1',
                    Suggested_Key: 'OUT-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: '',
                    wsu_school: 'UH Maui'
                },
                {
                    Error_Type: 'Duplicate_Source',
                    _rawErrorType: 'Duplicate_Source',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-1',
                    Duplicate_Group: 'G-1',
                    outcomes_school: 'UH Maui',
                    wsu_school: 'UH Maui'
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui' }],
                    translate: [
                        { translate_input: 'BAD-KEY', translate_output: 'WSU-1' },
                        { translate_input: 'OUT-1', translate_output: 'WSU-1' }
                    ],
                    wsu_org: [{ key: 'WSU-1', school: 'UH Maui' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        const suggestedKeyCol = findHeaderIndex(reviewSheet, 'Suggested Key');
        const decisionCol = findHeaderIndex(reviewSheet, 'Decision');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        assert.ok(suggestedKeyCol > 0, 'Review sheet should have Suggested Key column');
        assert.ok(decisionCol > 0, 'Review sheet should have Decision column');
        const rowCount = reviewSheet.rowCount || 0;
        let inputNotFoundCount = 0;
        let duplicateSourceCount = 0;
        let inputNotFoundDecision = '';
        let duplicateSourceSuggestedKey = '';
        let duplicateSourceDecision = '';
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Input key not found in Outcomes') {
                inputNotFoundCount += 1;
                inputNotFoundDecision = String(reviewSheet.getRow(r).getCell(decisionCol).value || '');
            }
            if (val === 'Duplicate_Source') {
                duplicateSourceCount += 1;
                duplicateSourceSuggestedKey = String(reviewSheet.getRow(r).getCell(suggestedKeyCol).value || '');
                duplicateSourceDecision = String(reviewSheet.getRow(r).getCell(decisionCol).value || '');
            }
        }
        assert.equal(inputNotFoundCount, 1, 'Input_Not_Found row should remain');
        assert.equal(duplicateSourceCount, 1, 'Duplicate_Source row should remain (regression: do not dedupe)');
        assert.equal(
            inputNotFoundDecision,
            'Use Suggestion',
            `Input_Not_Found Decision (got: ${JSON.stringify(inputNotFoundDecision)}); canonical pair (OUT-1, WSU-1)`
        );
        // Duplicate_Source with suggestion=same-as-current: we no longer show that as Suggested_Key (avoids no-op confusion)
        assert.equal(
            duplicateSourceSuggestedKey,
            '',
            `Duplicate_Source Suggested_Key should be blank when suggestion equals current (got: ${JSON.stringify(duplicateSourceSuggestedKey)})`
        );
        // One-to-Many rows default to Allow One-to-Many when no actionable suggestion (not Use Suggestion)
        assert.equal(
            duplicateSourceDecision,
            'Allow One-to-Many',
            `Duplicate_Source Decision should be Allow One-to-Many when no actionable suggestion (got: ${JSON.stringify(duplicateSourceDecision)})`
        );
    });

    await runCheck('buildValidationExport does not dedupe Error vs Duplicate when Error does not default to Use Suggestion', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Output_Not_Found',
                    _rawErrorType: 'Output_Not_Found',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-STALE',
                    Suggested_Key: 'WSU-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: 'UH Maui',
                    wsu_school: ''
                },
                {
                    Error_Type: 'Duplicate_Target',
                    _rawErrorType: 'Duplicate_Target',
                    translate_input: 'OUT-1',
                    translate_output: 'WSU-OLD',
                    Suggested_Key: 'WSU-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.85,
                    Duplicate_Group: 'G-1',
                    outcomes_school: 'UH Maui',
                    wsu_school: 'Legacy Org'
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui' }],
                    translate: [
                        { translate_input: 'OUT-1', translate_output: 'WSU-STALE' },
                        { translate_input: 'OUT-1', translate_output: 'WSU-OLD' }
                    ],
                    wsu_org: [
                        { key: 'WSU-1', school: 'UH Maui' },
                        { key: 'WSU-OLD', school: 'Legacy Org' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        const rowCount = reviewSheet.rowCount || 0;
        let outputNotFoundCount = 0;
        let duplicateTargetCount = 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const val = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (val === 'Output key not found in myWSU') outputNotFoundCount += 1;
            if (val === 'Duplicate_Target') duplicateTargetCount += 1;
        }
        assert.equal(outputNotFoundCount, 1, 'Output_Not_Found row should be kept when no dedupe');
        assert.equal(duplicateTargetCount, 1, 'Duplicate_Target row should remain when Error does not default to Use Suggestion');
    });

    await runCheck('buildValidationExport uses Output Key_Update_Side for Duplicate_Target', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Duplicate_Target',
                    _rawErrorType: 'Duplicate_Target',
                    translate_input: 'IN-ALPHA',
                    translate_output: 'OUT-LEGACY',
                    Suggested_Key: 'OUT-BETTER',
                    Suggested_School: 'Alpha Campus',
                    Suggestion_Score: 0.9,
                    Duplicate_Group: 'G-1',
                    outcomes_school: 'Alpha Campus',
                    wsu_school: 'Legacy Org'
                }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: { includeSuggestions: true, nameCompareConfig: { enabled: true, outcomes: 'school', wsu: 'school', threshold: 0.8 } },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-ALPHA', school: 'Alpha Campus' }],
                    translate: [{ translate_input: 'IN-ALPHA', translate_output: 'OUT-LEGACY' }],
                    wsu_org: [
                        { key: 'OUT-LEGACY', school: 'Legacy Org' },
                        { key: 'OUT-BETTER', school: 'Alpha Campus' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const keyUpdateSideCol = findHeaderIndex(reviewSheet, 'Update Side');
        assert.ok(keyUpdateSideCol > 0, 'Review sheet should have Update Side column');
        const keyUpdateSideValue = String(reviewSheet.getRow(2).getCell(keyUpdateSideCol).value || '');
        assert.equal(keyUpdateSideValue, 'Output', 'Duplicate_Target should have Key_Update_Side=Output');
    });

    await runCheck('P2.1: Duplicate_Target defaults to Keep As-Is when no actionable suggestion', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-PARENT', outcomes_school: 'Campus A', wsu_school: 'Parent', Suggested_Key: '', Suggestion_Score: '' },
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-B', translate_output: 'OUT-PARENT', outcomes_school: 'Campus B', wsu_school: 'Parent', Suggested_Key: '', Suggestion_Score: '' }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }, { key: 'IN-B', school: 'Campus B' }],
                    translate: [
                        { translate_input: 'IN-A', translate_output: 'OUT-PARENT' },
                        { translate_input: 'IN-B', translate_output: 'OUT-PARENT' }
                    ],
                    wsu_org: [{ key: 'OUT-PARENT', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const decisionCol = findHeaderIndex(reviewSheet, 'Decision');
        assert.ok(decisionCol > 0, 'Review sheet should have Decision column');
        const rowCount = reviewSheet.rowCount || 0;
        for (let r = 2; r <= rowCount; r += 1) {
            const decision = String(reviewSheet.getRow(r).getCell(decisionCol).value || '');
            assert.equal(decision, 'Keep As-Is', `Duplicate_Target row ${r} should default to Keep As-Is when no suggestion (got: ${decision})`);
        }
    });

    await runCheck('P2.2 re-import: prior decisions applied when Review_Row_ID matches', async () => {
        const harness = createHarness();
        const basePayload = {
            validated: [
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-PARENT', outcomes_key: 'IN-A', wsu_key: 'OUT-PARENT', outcomes_school: 'Campus A', wsu_school: 'Parent' }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }],
                    translate: [{ translate_input: 'IN-A', translate_output: 'OUT-PARENT' }],
                    wsu_org: [{ key: 'OUT-PARENT', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        };
        const run1 = await harness.buildValidationExport(basePayload);
        assertExportResult(run1);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const reviewRowIdCol = findHeaderIndex(reviewSheet, 'Review Row ID');
        assert.ok(reviewRowIdCol > 0, 'Review sheet should have Review Row ID column');
        const reviewRowId = String(reviewSheet.getRow(2).getCell(reviewRowIdCol).value || '');
        assert.ok(reviewRowId.length > 0, 'Review_Row_ID should be non-empty');
        const priorDecisions = {
            [reviewRowId]: { Decision: 'Ignore', Reason_Code: 'reimport-test' }
        };
        const run2 = await harness.buildValidationExport({ ...basePayload, priorDecisions });
        assertExportResult(run2);
        assert.ok(run2.reimportSummary, 'Result should include reimportSummary when priorDecisions used');
        assert.equal(run2.reimportSummary.applied, 1, 'One decision should be applied');
        assert.equal(run2.reimportSummary.conflicts, 0, 'No conflicts expected');
        const reviewSheet2 = harness.getLastWorkbook().getWorksheet('Review_Workbench');
        const decisionCol = findHeaderIndex(reviewSheet2, 'Decision');
        const decisionValue = String(reviewSheet2.getRow(2).getCell(decisionCol).value || '');
        assert.equal(decisionValue, 'Ignore', `Decision should be Ignore from prior (got: ${decisionValue})`);
    });

    await runCheck('P2.2 re-import: no priorDecisions yields no reimportSummary', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [{ Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-X', translate_output: 'OUT-Y', outcomes_key: 'IN-X', wsu_key: 'OUT-Y' }],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-X', school: 'X' }],
                    translate: [{ translate_input: 'IN-X', translate_output: 'OUT-Y' }],
                    wsu_org: [{ key: 'OUT-Y', school: 'Y' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        assert.equal(result.reimportSummary, undefined, 'reimportSummary should be absent when no priorDecisions');
    });

    await runCheck('P2.2 re-import: orphaned count when prior has extra IDs', async () => {
        const harness = createHarness();
        const payload = {
            validated: [{ Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-PARENT', outcomes_key: 'IN-A', wsu_key: 'OUT-PARENT' }],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }],
                    translate: [{ translate_input: 'IN-A', translate_output: 'OUT-PARENT' }],
                    wsu_org: [{ key: 'OUT-PARENT', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        };
        const run1 = await harness.buildValidationExport(payload);
        assertExportResult(run1);
        const reviewSheet = harness.getLastWorkbook().getWorksheet('Review_Workbench');
        const reviewRowIdCol = findHeaderIndex(reviewSheet, 'Review Row ID');
        const reviewRowId = String(reviewSheet.getRow(2).getCell(reviewRowIdCol).value || '');
        const priorDecisions = {
            [reviewRowId]: { Decision: 'Keep As-Is' },
            'NonExistent|ID|X|Y||||': { Decision: 'Ignore' }
        };
        const run2 = await harness.buildValidationExport({ ...payload, priorDecisions });
        assertExportResult(run2);
        assert.equal(run2.reimportSummary.applied, 1, 'One decision applied');
        assert.equal(run2.reimportSummary.orphaned, 1, 'One orphaned (non-existent ID in prior)');
    });

    await runCheck('P1.2: Reason_Code column and risky-without-reason QA row exist', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [{ Error_Type: 'Allow One-to-Many', Source_Sheet: 'One_to_Many', translate_input: 'IN-X', translate_output: 'OUT-Y', outcomes_key: 'IN-X', wsu_key: 'OUT-Y' }],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-X', school: 'X' }],
                    translate: [{ translate_input: 'IN-X', translate_output: 'OUT-Y' }],
                    wsu_org: [{ key: 'OUT-Y', school: 'Y' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        const reasonCodeCol = findHeaderIndex(reviewSheet, 'Reason Code');
        assert.ok(reasonCodeCol > 0, 'Review_Workbench should have Reason Code column');
        const qaSheet = workbook.getWorksheet('QA_Checks_Validate');
        const riskyRow = [...Array(20)].map((_, i) => qaSheet.getRow(i + 1).getCell(1).value).findIndex(v => String(v || '').includes('Risky decisions without reason code'));
        assert.ok(riskyRow >= 0, 'QA sheet should have Risky decisions without reason code row');
    });

    await runCheck('P0.3: get_action_queue returns actionQueueRows; preEditedActionQueueRows merge into export', async () => {
        const harness = createHarness();
        const basePayload = {
            validated: [
                { Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-P', outcomes_key: 'IN-A', wsu_key: 'OUT-P', outcomes_school: 'Campus A', wsu_school: 'Parent' }
            ],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }],
                    translate: [{ translate_input: 'IN-A', translate_output: 'OUT-P' }],
                    wsu_org: [{ key: 'OUT-P', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        };
        const queueResult = await harness.buildValidationExport({ ...basePayload, returnActionQueueOnly: true });
        assert.ok(Array.isArray(queueResult.actionQueueRows), 'returnActionQueueOnly should return actionQueueRows');
        assert.ok(queueResult.actionQueueRows.length >= 1, 'Should have at least one row');
        const row = queueResult.actionQueueRows[0];
        const rid = row.Review_Row_ID || '';
        assert.ok(rid.length > 0, 'Row should have Review_Row_ID');
        const preEdited = queueResult.actionQueueRows.map(r => ({ ...r }));
        preEdited[0].Decision = 'Ignore';
        preEdited[0].Selected_Candidate_ID = 'C3';
        preEdited[0].Manual_Suggested_Key = 'CUSTOM-KEY';
        const fullResult = await harness.buildValidationExport({ ...basePayload, preEditedActionQueueRows: preEdited });
        assertExportResult(fullResult);
        const reviewSheet = harness.getLastWorkbook().getWorksheet('Review_Workbench');
        const decisionCol = findHeaderIndex(reviewSheet, 'Decision');
        const selectedCandidateCol = findHeaderIndex(reviewSheet, 'Selected Candidate ID');
        const manualCol = findHeaderIndex(reviewSheet, 'Manual Key (override when no candidates)');
        assert.equal(String(reviewSheet.getRow(2).getCell(decisionCol).value || ''), 'Ignore', 'Decision should be Ignore from preEdited');
        assert.equal(String(reviewSheet.getRow(2).getCell(selectedCandidateCol).value || ''), 'C3', 'Selected_Candidate_ID should merge from preEdited');
        assert.equal(String(reviewSheet.getRow(2).getCell(manualCol).value || ''), 'CUSTOM-KEY', 'Manual_Suggested_Key should be CUSTOM-KEY from preEdited');
    });

    await runCheck('P1.3: campus-family rules prefill Manual_Suggested_Key when pattern matches', async () => {
        const harness = createHarness();
        const payload = {
            validated: [{
                Error_Type: 'Duplicate_Target',
                Source_Sheet: 'One_to_Many',
                translate_input: 'IN-TAMU-CT',
                translate_output: 'OUT-OLD',
                outcomes_key: 'IN-TAMU-CT',
                wsu_key: 'OUT-OLD',
                outcomes_school: 'Texas A&M University - Central Texas'
            }],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            options: {
                nameCompareConfig: { outcomes: 'school', wsu: 'school', enabled: true },
                campusFamilyRules: {
                    version: 1,
                    patterns: [{ pattern: 'Texas A&M*', parentKey: 'TAMU-MAIN', enabled: true }]
                }
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-TAMU-CT', school: 'Texas A&M University - Central Texas' }],
                    translate: [{ translate_input: 'IN-TAMU-CT', translate_output: 'OUT-OLD' }],
                    wsu_org: [
                        { key: 'OUT-OLD', school: 'Old campus' },
                        { key: 'TAMU-MAIN', school: 'Texas A&M University' }
                    ]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        };
        const result = await harness.buildValidationExport(payload);
        assertExportResult(result);
        const reviewSheet = harness.getLastWorkbook().getWorksheet('Review_Workbench');
        const manualKeyCol = findHeaderIndex(reviewSheet, 'Manual Key (override when no candidates)');
        assert.ok(manualKeyCol > 0, 'Review sheet should have Manual Key column');
        const manualKeyValue = String(reviewSheet.getRow(2).getCell(manualKeyCol).value || '');
        assert.equal(manualKeyValue, 'TAMU-MAIN', `Manual_Suggested_Key should be prefilled by campus-family (got: ${manualKeyValue})`);
    });

    await runCheck('P2.2 re-import: conflict when Use Suggestion key invalid', async () => {
        const harness = createHarness();
        const payload = {
            validated: [{ Error_Type: 'Duplicate_Target', Source_Sheet: 'One_to_Many', translate_input: 'IN-A', translate_output: 'OUT-PARENT', outcomes_key: 'IN-A', wsu_key: 'OUT-PARENT' }],
            selectedCols: { outcomes: ['school'], wsu_org: ['school'] },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-A', school: 'Campus A' }],
                    translate: [{ translate_input: 'IN-A', translate_output: 'OUT-PARENT' }],
                    wsu_org: [{ key: 'OUT-PARENT', school: 'Parent' }]
                },
                keyConfig: { outcomes: 'key', translateInput: 'translate_input', translateOutput: 'translate_output', wsu: 'key' },
                keyLabels: { outcomes: 'key', wsu: 'key', translateInput: 'translate_input', translateOutput: 'translate_output' },
                columnRoles: { outcomes: { school: 'School' }, wsu_org: { school: 'School' } }
            }
        };
        const run1 = await harness.buildValidationExport(payload);
        assertExportResult(run1);
        const reviewSheet = harness.getLastWorkbook().getWorksheet('Review_Workbench');
        const reviewRowIdCol = findHeaderIndex(reviewSheet, 'Review Row ID');
        const reviewRowId = String(reviewSheet.getRow(2).getCell(reviewRowIdCol).value || '');
        const priorDecisions = {
            [reviewRowId]: { Decision: 'Use Suggestion', Manual_Suggested_Key: 'INVALID-KEY-NOT-IN-WSU' }
        };
        const run2 = await harness.buildValidationExport({ ...payload, priorDecisions });
        assertExportResult(run2);
        assert.equal(run2.reimportSummary.applied, 0, 'No decision applied (invalid key)');
        assert.equal(run2.reimportSummary.conflicts, 1, 'One conflict (invalid key)');
        const decisionCol = findHeaderIndex(harness.getLastWorkbook().getWorksheet('Review_Workbench'), 'Decision');
        const decisionValue = String(harness.getLastWorkbook().getWorksheet('Review_Workbench').getRow(2).getCell(decisionCol).value || '');
        assert.equal(decisionValue, 'Keep As-Is', 'Decision should remain default (Keep As-Is) when conflict');
    });

    await runCheck('buildGenerationExport includes create review guidance columns and instructions', async () => {
        const harness = createHarness();
        const result = await harness.buildGenerationExport({
            cleanRows: [
                {
                    outcomes_row_index: 0,
                    outcomes_record_id: 'OUT-1',
                    outcomes_display_name: 'Alpha University',
                    proposed_wsu_key: 'WSU-1',
                    proposed_wsu_name: 'Alpha University',
                    match_similarity: 97,
                    confidence_tier: 'high',
                    outcomes_school: 'Alpha University',
                    wsu_school: 'Alpha University'
                }
            ],
            errorRows: [
                {
                    outcomes_row_index: 1,
                    outcomes_record_id: 'OUT-2',
                    outcomes_display_name: 'Beta College',
                    missing_in: 'Ambiguous Match',
                    proposed_wsu_key: 'WSU-2A',
                    proposed_wsu_name: 'Beta College Main',
                    alternate_candidates: [
                        { key: 'WSU-2A', name: 'Beta College Main', similarity: 94 },
                        { key: 'WSU-2B', name: 'Beta College South', similarity: 92 }
                    ],
                    outcomes_school: 'Beta College',
                    wsu_school: ''
                },
                {
                    outcomes_row_index: 2,
                    outcomes_record_id: 'OUT-3',
                    outcomes_display_name: 'Gamma Institute',
                    missing_in: 'myWSU',
                    alternate_candidates: [],
                    outcomes_school: 'Gamma Institute',
                    wsu_school: ''
                }
            ],
            selectedCols: {
                outcomes: ['school'],
                wsu_org: ['school']
            },
            generationConfig: {
                threshold: 0.8
            }
        });
        assertExportResult(result);

        const workbook = harness.getLastWorkbook();
        assert.equal(
            workbook.calcProperties && workbook.calcProperties.fullCalcOnLoad,
            true,
            'Generation workbook should force full recalculation on open'
        );
        const ambiguousSheet = workbook.getWorksheet('Ambiguous_Candidates');
        const missingSheet = workbook.getWorksheet('Missing_In_myWSU');
        const reviewSheet = workbook.getWorksheet('Review_Decisions');
        const instructionsSheet = workbook.getWorksheet('Review_Instructions_Create');
        assert.ok(ambiguousSheet, 'Expected Ambiguous_Candidates worksheet');
        assert.ok(missingSheet, 'Expected Missing_In_myWSU worksheet');
        assert.ok(reviewSheet, 'Expected Review_Decisions worksheet');
        assert.ok(instructionsSheet, 'Expected Review_Instructions_Create worksheet');

        const ambiguousHeaders = getRowValues(ambiguousSheet, 1, 40).map(v => String(v || ''));
        assert.ok(ambiguousHeaders.includes('Resolution Type'), 'Ambiguous sheet should include Resolution Type');
        assert.ok(ambiguousHeaders.includes('Review Path'), 'Ambiguous sheet should include Review Path');
        assert.ok(ambiguousHeaders.includes('Candidate Count'), 'Ambiguous sheet should include Candidate Count');
        assert.ok(ambiguousHeaders.includes('Top Suggested myWSU Key'), 'Ambiguous sheet should include top suggested key');

        const reviewHeaders = getRowValues(reviewSheet, 1, 40).map(v => String(v || ''));
        assert.ok(reviewHeaders.includes('Resolution Type'), 'Review sheet should include Resolution Type');
        assert.ok(reviewHeaders.includes('Review Path'), 'Review sheet should include Review Path');
        assert.ok(reviewHeaders.includes('Candidate Count'), 'Review sheet should include Candidate Count');

        const candidateCountIndex = ambiguousHeaders.indexOf('Candidate Count');
        assert.ok(candidateCountIndex >= 0, 'Candidate Count column must exist');
        const candidateCountValue = ambiguousSheet.getRow(2).getCell(candidateCountIndex + 1).value;
        assert.equal(candidateCountValue, 2, 'Ambiguous row should expose candidate count');

        const allCFRules = reviewSheet.conditionalFormatting.flatMap(cf => cf.rules || []);
        const expressionRules = allCFRules.filter(r => r.type === 'expression');
        assert.ok(expressionRules.length >= 3, 'Review sheet should include expression conditional formatting rules');
        expressionRules.forEach((rule, idx) => {
            assert.ok(Array.isArray(rule.formulae), `Create expression rule ${idx} should use formulae array`);
            assert.equal(rule.formula, undefined, `Create expression rule ${idx} should not use singular formula`);
        });
    });

    await runCheck('expression-type CF rules use formulae array (not formula string)', async () => {
        const harness = createHarness();
        await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Name_Mismatch',
                    translate_input: 'IN-1',
                    translate_output: 'OUT-1',
                    outcomes_school: 'Alpha University',
                    wsu_school: 'Alpha University',
                    Source_Sheet: 'One_to_Many',
                    Is_Stale_Key: 0,
                    duplicateGroup: ''
                }
            ],
            selectedCols: {
                outcomes: ['school'],
                wsu_org: ['school']
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-1', school: 'Alpha University' }],
                    translate: [{ translate_input: 'IN-1', translate_output: 'OUT-1' }],
                    wsu_org: [{ key: 'OUT-1', school: 'Alpha University' }]
                },
                keyConfig: {
                    outcomes: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'key'
                },
                keyLabels: {
                    outcomes: 'key',
                    wsu: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                },
                columnRoles: {
                    outcomes: { school: 'School' },
                    wsu_org: { school: 'School' }
                }
            }
        });
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench worksheet');
        const allCFRules = reviewSheet.conditionalFormatting.flatMap(cf => cf.rules || []);
        const expressionRules = allCFRules.filter(r => r.type === 'expression');
        assert.ok(expressionRules.length > 0, 'Expected at least one expression-type CF rule');
        expressionRules.forEach((rule, idx) => {
            assert.ok(
                Array.isArray(rule.formulae),
                `Expression rule ${idx}: formulae must be an array, got ${typeof rule.formulae}`
            );
            assert.ok(
                rule.formulae.length > 0,
                `Expression rule ${idx}: formulae array must not be empty`
            );
            assert.equal(
                rule.formula,
                undefined,
                `Expression rule ${idx}: must not have singular 'formula' property (ExcelJS requires 'formulae' array)`
            );
        });
    });

    await runCheck('duplicate-target suggestions keep non-current candidates for many-to-one selection', async () => {
        const harness = createHarness();
        const queueResult = await harness.buildValidationExport({
            validated: [
                {
                    Error_Type: 'Duplicate_Target',
                    translate_input: 'IN-ALPHA',
                    translate_output: 'OUT-LEGACY',
                    outcomes_school: 'Alpha Campus',
                    wsu_school: 'Legacy Org'
                }
            ],
            selectedCols: {
                outcomes: ['school'],
                wsu_org: ['school']
            },
            options: {
                includeSuggestions: true,
                nameCompareConfig: {
                    enabled: true,
                    outcomes: 'school',
                    wsu: 'school',
                    threshold: 0.8
                }
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'IN-ALPHA', school: 'Alpha Campus' }],
                    translate: [{ translate_input: 'IN-ALPHA', translate_output: 'OUT-LEGACY' }],
                    wsu_org: [
                        { key: 'OUT-LEGACY', school: 'Alpha Campus' },
                        { key: 'OUT-BETTER', school: 'Alpha Campus' },
                        { key: 'OUT-BEST', school: 'Alpha Campus' },
                        { key: 'OUT-ALT-1', school: 'Alpha Campus' },
                        { key: 'OUT-ALT-2', school: 'Alpha Campus' },
                        { key: 'OUT-ALT-3', school: 'Alpha Campus' },
                        { key: 'OUT-ALT-4', school: 'Alpha Campus' },
                        { key: 'OUT-ALT-5', school: 'Alpha Campus' }
                    ]
                },
                keyConfig: {
                    outcomes: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'key'
                },
                keyLabels: {
                    outcomes: 'key',
                    wsu: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                },
                columnRoles: {
                    outcomes: { school: 'School' },
                    wsu_org: { school: 'School' }
                }
            },
            returnActionQueueOnly: true
        });
        assert.ok(Array.isArray(queueResult.actionQueueRows), 'Expected actionQueueRows array');
        const row = queueResult.actionQueueRows.find(r => String(r.Error_Type || '') === 'Duplicate_Target');
        assert.ok(row, 'Expected Duplicate_Target row in action queue');
        const candidateKeys = (row._candidates || []).map(c => String(c.key || ''));
        assert.ok(candidateKeys.includes('OUT-BETTER'), 'Expected alternate non-current candidate in dropdown list');
        assert.ok(!candidateKeys.includes('OUT-LEGACY'), 'Current output key should be removed from candidate dropdown list');
        assert.ok(candidateKeys.length > 0, 'Duplicate-target candidate list should include at least one non-current option');
        assert.ok(candidateKeys.length <= 5, 'Duplicate-target candidate list should be capped at 5 location-valid matches');
    });

    await runCheck('missing-mapping suggestions require strong non-location name evidence', async () => {
        const harness = createHarness();
        const result = await harness.buildValidationExport({
            validated: [],
            selectedCols: {
                outcomes: ['school', 'city', 'state', 'country'],
                wsu_org: ['school', 'city', 'state', 'country']
            },
            options: {
                includeSuggestions: true,
                nameCompareConfig: {
                    enabled: true,
                    outcomes: 'school',
                    wsu: 'school',
                    threshold: 0.8
                }
            },
            context: {
                loadedData: {
                    outcomes: [
                        { mdb_code: 'O1', school: 'ACADEMY IN ARCHITECTURE', city: '', state: 'OT', country: 'IN' },
                        { mdb_code: 'O2', school: 'ADLER UNIVERSITY - VANCOUVER', city: 'Vancouver', state: 'BC', country: 'CA' },
                        { mdb_code: 'O3', school: 'AGA KHAN UNIVERSITY', city: '', state: 'OT', country: 'PK' },
                        { mdb_code: 'O4', school: 'ALPHA CAMPUS', city: 'Seattle', state: 'WA', country: 'US' }
                    ],
                    translate: [],
                    wsu_org: [
                        { 'Org ID': '011612666', school: 'Jawaharlal Nehru Arch Fin Art', city: 'Hyderabad', state: 'AP', country: 'IND' },
                        { 'Org ID': '011814422', school: 'LaSalle College Vancouver', city: 'Vancouver', state: 'BC', country: 'CAN' },
                        { 'Org ID': '011456413', school: 'Al-Khair University', city: '', state: '', country: 'PAK' },
                        { 'Org ID': '011900004', school: 'Alpha Campus', city: 'Seattle', state: 'WA', country: 'USA' }
                    ]
                },
                keyConfig: {
                    outcomes: 'mdb_code',
                    translateInput: 'input',
                    translateOutput: 'output',
                    wsu: 'Org ID'
                },
                keyLabels: {
                    outcomes: 'mdb_code',
                    wsu: 'Org ID',
                    translateInput: 'input',
                    translateOutput: 'output'
                },
                columnRoles: {
                    outcomes: { school: 'School', city: 'City', state: 'State', country: 'Country' },
                    wsu_org: { school: 'School', city: 'City', state: 'State', country: 'Country' }
                }
            }
        });
        assertExportResult(result);
        const workbook = harness.getLastWorkbook();
        const reviewSheet = workbook.getWorksheet('Review_Workbench');
        assert.ok(reviewSheet, 'Expected Review_Workbench worksheet');
        const errorTypeCol = findHeaderIndex(reviewSheet, 'Error Type');
        const currentInputCol = findHeaderIndex(reviewSheet, 'Current Translate Input');
        const currentOutputCol = findHeaderIndex(reviewSheet, 'Current Translate Output');
        assert.ok(errorTypeCol > 0, 'Review sheet should have Error Type column');
        assert.ok(currentInputCol > 0, 'Review sheet should have Current Translate Input column');
        assert.ok(currentOutputCol > 0, 'Review sheet should have Current Translate Output column');

        const missingPairs = [];
        for (let r = 2; r <= (reviewSheet.rowCount || 0); r += 1) {
            const errorType = String(reviewSheet.getRow(r).getCell(errorTypeCol).value || '');
            if (errorType !== 'Missing_Mapping') continue;
            const inputKey = String(reviewSheet.getRow(r).getCell(currentInputCol).value || '');
            const outputKey = String(reviewSheet.getRow(r).getCell(currentOutputCol).value || '');
            missingPairs.push([inputKey, outputKey]);
        }

        assert.equal(missingPairs.length, 1, 'Only strong missing-mapping pair should remain');
        assert.deepEqual(
            missingPairs[0],
            ['O4', '011900004'],
            'Weak location-driven name pairs should be excluded from Missing_Mapping suggestions'
        );
    });

    await runCheck('translation-only export scope excludes Missing_Mappings from report and review queue', async () => {
        const basePayload = {
            validated: [
                {
                    Error_Type: 'Input_Not_Found',
                    _rawErrorType: 'Input_Not_Found',
                    translate_input: 'BAD-KEY',
                    translate_output: 'WSU-1',
                    Suggested_Key: 'OUT-1',
                    Suggested_School: 'UH Maui',
                    Suggestion_Score: 0.9,
                    outcomes_school: '',
                    wsu_school: 'UH Maui'
                }
            ],
            selectedCols: {
                outcomes: ['school', 'state'],
                wsu_org: ['school', 'state']
            },
            options: {
                includeSuggestions: true,
                nameCompareConfig: {
                    enabled: true,
                    outcomes: 'school',
                    wsu: 'school',
                    threshold: 0.8
                }
            },
            context: {
                loadedData: {
                    outcomes: [{ key: 'OUT-1', school: 'UH Maui', state: 'HI' }],
                    translate: [{ translate_input: 'BAD-KEY', translate_output: 'WSU-1' }],
                    wsu_org: [{ key: 'WSU-1', school: 'UH Maui', state: 'HI' }]
                },
                keyConfig: {
                    outcomes: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output',
                    wsu: 'key'
                },
                keyLabels: {
                    outcomes: 'key',
                    wsu: 'key',
                    translateInput: 'translate_input',
                    translateOutput: 'translate_output'
                },
                columnRoles: {
                    outcomes: { school: 'School', state: 'State' },
                    wsu_org: { school: 'School', state: 'State' }
                }
            }
        };

        const harnessDefault = createHarness();
        const defaultResult = await harnessDefault.buildValidationExport(basePayload);
        assertExportResult(defaultResult);
        const workbookDefault = harnessDefault.getLastWorkbook();
        const defaultMissingSheet = workbookDefault.getWorksheet('Missing_Mappings');
        assert.ok(defaultMissingSheet, 'Expected Missing_Mappings sheet in default export');
        const defaultReview = workbookDefault.getWorksheet('Review_Workbench');
        const defaultErrorTypeCol = findHeaderIndex(defaultReview, 'Error Type');
        let defaultMissingRows = 0;
        for (let r = 2; r <= (defaultReview.rowCount || 0); r += 1) {
            const value = String(defaultReview.getRow(r).getCell(defaultErrorTypeCol).value || '');
            if (value === 'Missing_Mapping') defaultMissingRows += 1;
        }
        assert.ok(defaultMissingRows > 0, 'Expected Missing_Mapping rows in default review queue');

        const harnessScoped = createHarness();
        const scopedResult = await harnessScoped.buildValidationExport({
            ...basePayload,
            options: {
                ...basePayload.options,
                reviewScope: 'translation_only'
            }
        });
        assertExportResult(scopedResult);
        const workbookScoped = harnessScoped.getLastWorkbook();
        assert.equal(workbookScoped.getWorksheet('Missing_Mappings'), undefined, 'Missing_Mappings sheet should be excluded when translation-only export scope is active');
        const scopedReview = workbookScoped.getWorksheet('Review_Workbench');
        const scopedErrorTypeCol = findHeaderIndex(scopedReview, 'Error Type');
        let scopedMissingRows = 0;
        for (let r = 2; r <= (scopedReview.rowCount || 0); r += 1) {
            const value = String(scopedReview.getRow(r).getCell(scopedErrorTypeCol).value || '');
            if (value === 'Missing_Mapping') scopedMissingRows += 1;
        }
        assert.equal(scopedMissingRows, 0, 'Missing_Mapping rows should be excluded when translation-only export scope is active');

        const harnessMissingOnly = createHarness();
        const missingOnlyResult = await harnessMissingOnly.buildValidationExport({
            ...basePayload,
            options: {
                ...basePayload.options,
                reviewScope: 'missing_only'
            }
        });
        assertExportResult(missingOnlyResult);
        const workbookMissingOnly = harnessMissingOnly.getLastWorkbook();
        const missingOnlySheet = workbookMissingOnly.getWorksheet('Missing_Mappings');
        assert.ok(missingOnlySheet, 'Missing_Mappings sheet should be present when missing-only export scope is active');
        assert.equal(workbookMissingOnly.getWorksheet('Errors_in_Translate'), undefined, 'Errors_in_Translate sheet should be excluded in missing-only export scope');
        assert.equal(workbookMissingOnly.getWorksheet('One_to_Many'), undefined, 'One_to_Many sheet should be excluded in missing-only export scope');
        assert.equal(workbookMissingOnly.getWorksheet('High_Confidence_Matches'), undefined, 'High_Confidence_Matches sheet should be excluded in missing-only export scope');
        assert.equal(workbookMissingOnly.getWorksheet('Valid_Mappings'), undefined, 'Valid_Mappings sheet should be excluded in missing-only export scope');
        const missingOnlyReview = workbookMissingOnly.getWorksheet('Review_Workbench');
        const missingOnlyErrorTypeCol = findHeaderIndex(missingOnlyReview, 'Error Type');
        let missingOnlyRows = 0;
        let nonMissingOnlyRows = 0;
        for (let r = 2; r <= (missingOnlyReview.rowCount || 0); r += 1) {
            const value = String(missingOnlyReview.getRow(r).getCell(missingOnlyErrorTypeCol).value || '');
            if (value === 'Missing_Mapping') missingOnlyRows += 1;
            else if (value) nonMissingOnlyRows += 1;
        }
        assert.ok(missingOnlyRows > 0, 'Missing-only export scope should include Missing_Mapping rows in review queue');
        assert.equal(nonMissingOnlyRows, 0, 'Missing-only export scope should exclude non-missing rows from review queue');
        const approvedSheet = workbookMissingOnly.getWorksheet('Approved_Mappings');
        const approvalSourceCol = findHeaderIndex(approvedSheet, 'Approval Source');
        let autoApprovedRows = 0;
        for (let r = 2; r <= (approvedSheet.rowCount || 0); r += 1) {
            const cellValue = approvedSheet.getRow(r).getCell(approvalSourceCol).value;
            const value = (cellValue && typeof cellValue === 'object' && Object.prototype.hasOwnProperty.call(cellValue, 'formula'))
                ? ''
                : String(cellValue || '');
            if (value === 'Valid_Mappings' || value === 'High_Confidence_Matches') autoApprovedRows += 1;
        }
        assert.equal(autoApprovedRows, 0, 'Missing-only export scope should exclude auto-approved valid/high-confidence rows from final export staging');
    });

    if (failures > 0) {
        console.error(`\n${failures} export test(s) failed.`);
        process.exit(1);
    }

    console.log('\nAll export validation tests passed.');
}

run().catch(error => {
    console.error(`[FAIL] export-test fatal error: ${error.message}`);
    process.exit(1);
});
