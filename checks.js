'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const validationPath = path.join(__dirname, 'validation.js');
const validationCode = fs.readFileSync(validationPath, 'utf8');

const context = {
    console
};
vm.createContext(context);
vm.runInContext(validationCode, context, { filename: validationPath });

let failures = 0;

function runCheck(name, fn) {
    try {
        fn();
        console.log(`[PASS] ${name}`);
    } catch (error) {
        failures += 1;
        console.error(`[FAIL] ${name}: ${error.message}`);
    }
}

runCheck('countriesMatch: BD maps to Bangladesh', () => {
    assert.equal(context.countriesMatch('BD', 'Bangladesh'), true);
});

runCheck('countriesMatch: BG does not map to Bangladesh', () => {
    assert.equal(context.countriesMatch('BG', 'Bangladesh'), false);
});

runCheck('countriesMatch: NG maps to Nigeria', () => {
    assert.equal(context.countriesMatch('NG', 'Nigeria'), true);
});

runCheck('countriesMatch: NI does not map to Nigeria', () => {
    assert.equal(context.countriesMatch('NI', 'Nigeria'), false);
});

runCheck('mergeData throws on duplicate Outcomes keys', () => {
    const keyConfig = {
        outcomes: 'outcomes_key',
        translateInput: 'translate_input',
        translateOutput: 'translate_output',
        wsu: 'wsu_key'
    };
    const outcomes = [
        { outcomes_key: '1001', school: 'Alpha One' },
        { outcomes_key: '1001', school: 'Alpha Duplicate' }
    ];
    const translate = [{ translate_input: '1001', translate_output: '2001' }];
    const wsu = [{ wsu_key: '2001', school: 'Beta One' }];

    assert.throws(
        () => context.mergeData(outcomes, translate, wsu, keyConfig),
        /Outcomes source has duplicate key values/
    );
});

runCheck('mergeData throws on duplicate myWSU keys', () => {
    const keyConfig = {
        outcomes: 'outcomes_key',
        translateInput: 'translate_input',
        translateOutput: 'translate_output',
        wsu: 'wsu_key'
    };
    const outcomes = [{ outcomes_key: '1001', school: 'Alpha One' }];
    const translate = [{ translate_input: '1001', translate_output: '2001' }];
    const wsu = [
        { wsu_key: '2001', school: 'Beta One' },
        { wsu_key: '2001', school: 'Beta Duplicate' }
    ];

    assert.throws(
        () => context.mergeData(outcomes, translate, wsu, keyConfig),
        /myWSU source has duplicate key values/
    );
});

runCheck('mergeData succeeds with unique keys', () => {
    const keyConfig = {
        outcomes: 'outcomes_key',
        translateInput: 'translate_input',
        translateOutput: 'translate_output',
        wsu: 'wsu_key'
    };
    const outcomes = [{ outcomes_key: '1001', school: 'Alpha One' }];
    const translate = [{ translate_input: '1001', translate_output: '2001' }];
    const wsu = [{ wsu_key: '2001', school: 'Beta One' }];

    const merged = context.mergeData(outcomes, translate, wsu, keyConfig);
    assert.equal(merged.length, 1);
    assert.equal(merged[0].outcomes_school, 'Alpha One');
    assert.equal(merged[0].wsu_school, 'Beta One');
});

runCheck('normalizeNameForCompare supports UCLA-style abbreviations', () => {
    assert.equal(
        context.normalizeNameForCompare('Univ of Cal, Los Angeles'),
        'california los angeles'
    );
});

runCheck('calculateNameSimilarity scores UCLA alias pair as high confidence', () => {
    const score = context.calculateNameSimilarity(
        'UNIVERSITY OF CALIFORNIA - LOS ANGELES',
        'Univ of Cal, Los Angeles'
    );
    assert.ok(score >= 0.8, `Expected score >= 0.8, got ${score}`);
});

runCheck('jaroWinkler: cal/california >= 0.80', () => {
    assert.ok(context.jaroWinkler('cal', 'california') >= 0.8);
});

runCheck('jaroWinkler: schl/school >= 0.80', () => {
    assert.ok(context.jaroWinkler('schl', 'school') >= 0.8);
});

runCheck('jaroWinkler: mgmnt/management < 0.80 (known gap)', () => {
    assert.ok(context.jaroWinkler('mgmnt', 'management') < 0.8);
});

runCheck('buildTokenIDF: faulkner > state in sample corpus', () => {
    const names = [
        'Faulkner State Community College',
        'Walters State Community College',
        'Florida State University',
        'Ohio State University',
        'Harvard University'
    ];
    const idf = context.buildTokenIDF(names);
    assert.ok(idf.faulkner > idf.state, `Expected faulkner (${idf.faulkner}) > state (${idf.state})`);
});

runCheck('Faulkner/Walters false positive is prevented with IDF weighting', () => {
    const names = [
        'FAULKNER STATE COMMUNITY COLLEGE',
        'WALTERS STATE COMMUNITY COLLEGE',
        'ROANE STATE COMMUNITY COLLEGE',
        'NASHVILLE STATE COMMUNITY COLLEGE'
    ];
    const idf = context.buildTokenIDF(names);
    const score = context.calculateNameSimilarity(
        'FAULKNER STATE COMMUNITY COLLEGE',
        'WALTERS STATE COMMUNITY COLLEGE',
        idf
    );
    assert.ok(score < 0.8, `Expected score < 0.8, got ${score}`);
});

runCheck('fieldEvidence: high when country+state agree', () => {
    const ev = context.fieldEvidence('Test U', 'Test University', 'WA', 'WA', '', '', 'US', 'USA');
    assert.ok(ev > 0.7, `Expected > 0.7, got ${ev}`);
});

runCheck('fieldEvidence: lower when country disagrees', () => {
    const ev = context.fieldEvidence('Test U', 'Test University', 'WA', 'OR', '', '', 'US', 'Nigeria');
    assert.ok(ev < 0.5, `Expected < 0.5, got ${ev}`);
});

const helpersPath = path.join(__dirname, 'validation-export-helpers.js');
const helpersCode = fs.readFileSync(helpersPath, 'utf8');
const helpersContext = { module: { exports: {} }, require: () => {} };
vm.createContext(helpersContext);
vm.runInContext(helpersCode, helpersContext, { filename: helpersPath });
const helpers = helpersContext.module.exports;
const exportWorkerPath = path.join(__dirname, 'export-worker.js');
const exportWorkerCode = fs.readFileSync(exportWorkerPath, 'utf8');

runCheck('validation-export-helpers: OUTPUT_NOT_FOUND_SUBTYPE has three values', () => {
    assert.ok(helpers.OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY);
    assert.ok(helpers.OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT);
    assert.ok(helpers.OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT);
});

runCheck('validation-export-helpers: getPriority returns correct order', () => {
    assert.ok(helpers.getPriority('Missing_Input') < helpers.getPriority('Name_Mismatch'));
    assert.ok(helpers.getPriority('Output_Not_Found', helpers.OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT) <
        helpers.getPriority('Output_Not_Found', helpers.OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY));
});

runCheck('validation-export-helpers: getRecommendedAction returns non-empty', () => {
    assert.ok(helpers.getRecommendedAction('Missing_Input').length > 0);
    assert.ok(helpers.getRecommendedAction('Output_Not_Found', helpers.OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY).length > 0);
});

runCheck('generateSummaryStats returns Output_Not_Found subtype keys', () => {
    const validated = [
        { Error_Type: 'Output_Not_Found', Error_Subtype: 'Output_Not_Found_Likely_Stale_Key' },
        { Error_Type: 'Output_Not_Found', Error_Subtype: 'Output_Not_Found_Ambiguous_Replacement' },
        { Error_Type: 'Output_Not_Found', Error_Subtype: 'Output_Not_Found_No_Replacement' }
    ];
    const outcomes = [];
    const translate = [];
    const wsu = [];
    const stats = context.generateSummaryStats(validated, outcomes, translate, wsu);
    assert.ok('output_not_found_likely_stale_key' in stats.errors);
    assert.ok('output_not_found_ambiguous_replacement' in stats.errors);
    assert.ok('output_not_found_no_replacement' in stats.errors);
});

runCheck('Action_Queue context columns include Missing_In and Similarity', () => {
    assert.ok(helpers.ACTION_QUEUE_CONTEXT_COLUMNS.includes('Missing_In'));
    assert.ok(helpers.ACTION_QUEUE_CONTEXT_COLUMNS.includes('Similarity'));
});

runCheck('filterOutputNotFoundBySubtype uses raw subtype', () => {
    const rows = [
        { _rawErrorType: 'Output_Not_Found', _rawErrorSubtype: helpers.OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT },
        { _rawErrorType: 'Output_Not_Found', _rawErrorSubtype: helpers.OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT },
        { _rawErrorType: 'Input_Not_Found', _rawErrorSubtype: '' }
    ];
    const ambiguous = helpers.filterOutputNotFoundBySubtype(rows, helpers.OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT);
    const noRepl = helpers.filterOutputNotFoundBySubtype(rows, helpers.OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT);
    assert.equal(ambiguous.length, 1);
    assert.equal(noRepl.length, 1);
    assert.equal(ambiguous[0]._rawErrorSubtype, helpers.OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT);
    assert.equal(noRepl[0]._rawErrorSubtype, helpers.OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT);
});

runCheck('getQAValidateRowsForEmptyQueue returns valid structure', () => {
    const rows = helpers.getQAValidateRowsForEmptyQueue();
    assert.equal(rows.length, 16, 'Header + 15 check rows matching non-empty QA layout');
    assert.equal(rows[0][0], 'Check');
    assert.equal(rows[1][1], 0);
    assert.equal(rows[1][2], 'PASS');
    assert.equal(rows[1][3], 'Blank or Ignore');
    assert.equal(rows[2][2], 'PASS', 'Approved review rows status should be PASS when empty');
    assert.equal(rows[6][0], 'Use Suggestion with invalid Update Side');
    assert.equal(rows[7][0], 'Use Suggestion with invalid manual key');
    assert.equal(rows[8][0], 'Use Suggestion no-op (key equals current value)');
    assert.equal(rows[9][0], 'Risky decisions without reason code');
    assert.equal(rows[10][0], 'Stale-key rows lacking decision');
    assert.equal(rows[14][0], 'Duplicate (input, output) pairs in Final_Translation_Table');
    assert.equal(rows[15][0], 'Publish gate');
    assert.equal(rows[15][1], 'PASS', 'Empty-queue publish gate result in column B to match non-empty layout');
    assert.equal(rows[15][2], '', 'Empty-queue publish gate Status column C empty to match non-empty');
});

runCheck('export-worker: Input_Not_Found uses reverse name suggestion from myWSU', () => {
    assert.ok(
        exportWorkerCode.includes('getBestOutcomesNameSuggestion'),
        'Expected reverse name suggestion helper to exist'
    );
    assert.ok(
        exportWorkerCode.includes("row[`wsu_${nameCompareConfig.wsu}`] || row.wsu_Descr || ''"),
        'Expected Input_Not_Found suggestion to anchor on myWSU name columns'
    );
});

runCheck('export-worker: suggestion blocking/indexing helpers exist', () => {
    assert.ok(exportWorkerCode.includes('buildTokenIDFLocal'));
    assert.ok(exportWorkerCode.includes('buildTokenIndex'));
    assert.ok(exportWorkerCode.includes('getBlockedCandidateIndices'));
    assert.ok(exportWorkerCode.includes('suggestionBlockStats'));
});

runCheck('export-worker: Validate decision dropdown includes Allow One-to-Many', () => {
    assert.ok(
        exportWorkerCode.includes('"Keep As-Is,Use Suggestion,Allow One-to-Many,Ignore"'),
        'Expected expanded decision dropdown values'
    );
    assert.ok(
        !exportWorkerCode.includes('Needs Research'),
        'Needs Research should no longer appear in validate/create decision models'
    );
    assert.ok(
        !exportWorkerCode.includes('No Change'),
        'No Change should no longer appear in validate decision model'
    );
});

runCheck('export-worker: Validate review/approved/final/update sheets exist', () => {
    assert.ok(exportWorkerCode.includes("sheetName: 'Review_Workbench'"));
    assert.ok(exportWorkerCode.includes("sheetName: 'Approved_Mappings'"));
    assert.ok(exportWorkerCode.includes("addWorksheet('Final_Staging')"));
    assert.ok(exportWorkerCode.includes("addWorksheet('Final_Translation_Table')"));
    assert.ok(exportWorkerCode.includes("addWorksheet('Translation_Key_Updates')"));
});

runCheck('export-worker: Validate publish gate checks exist', () => {
    assert.ok(exportWorkerCode.includes("'Publish gate'"));
    assert.ok(exportWorkerCode.includes('B2=0'));
    assert.ok(exportWorkerCode.includes('B4=0'));
    assert.ok(exportWorkerCode.includes('B5=0'));
    assert.ok(exportWorkerCode.includes('B6=0'));
    assert.ok(exportWorkerCode.includes('B7=0'));
    assert.ok(exportWorkerCode.includes('B8=0'));
    assert.ok(exportWorkerCode.includes('B9=0'), 'Publish gate should include no-op check B9');
    assert.ok(exportWorkerCode.includes('B10=0'), 'Publish gate should include risky-without-reason B10');
    assert.ok(exportWorkerCode.includes('B13=0'), 'Publish gate should include duplicate input B13');
    assert.ok(exportWorkerCode.includes('B14=0'), 'Publish gate should include duplicate output B14');
    assert.ok(exportWorkerCode.includes('B15=0'), 'Publish gate should include duplicate pairs B15');
    assert.ok(exportWorkerCode.includes('"PASS","HOLD"'));
    assert.ok(exportWorkerCode.includes('Diagnostic tabs are hidden'));
});

runCheck('export-worker: Validate export uses capped review formula rows', () => {
    assert.ok(exportWorkerCode.includes('MAX_VALIDATE_DYNAMIC_REVIEW_FORMULA_ROWS'));
    assert.ok(exportWorkerCode.includes("'Approved rows beyond formula capacity'"));
});

runCheck('export-worker: Review workbook exposes explicit current/final key columns', () => {
    assert.ok(exportWorkerCode.includes("'Current_Input'"));
    assert.ok(exportWorkerCode.includes("'Current_Output'"));
    assert.ok(exportWorkerCode.includes("'Final_Input'"));
    assert.ok(exportWorkerCode.includes("'Final_Output'"));
});

runCheck('export-worker: Human review safeguards exist', () => {
    assert.ok(exportWorkerCode.includes('Decision_Warning'));
    assert.ok(exportWorkerCode.includes('Use Suggestion without effective key'));
    assert.ok(exportWorkerCode.includes('Use Suggestion needs'));
    assert.ok(exportWorkerCode.includes('valid Update Side'));
    assert.ok(exportWorkerCode.includes('Use Suggestion with invalid Update Side'));
    assert.ok(exportWorkerCode.includes('Use Suggestion with invalid manual key'));
    assert.ok(exportWorkerCode.includes('Approved but blank final'));
});

runCheck('export-worker: Review_Workbench remains intentionally unprotected', () => {
    const match = exportWorkerCode.match(
        /async function buildValidationExport[\s\S]*?(?=self\.onmessage\s*=)/
    );
    assert.ok(match, 'Could not locate buildValidationExport block');
    const validateBlock = match[0];
    assert.ok(
        validateBlock.includes('Workbook left unprotected so sort/filter work without restriction.'),
        'Validate export should document unprotected Review_Workbench behavior'
    );
    assert.ok(
        !/reviewSheet\.protect\s*\(/.test(validateBlock),
        'Review_Workbench should remain unprotected in validate export'
    );
});

runCheck('export-worker: Validate internal staging tabs are hidden', () => {
    assert.ok(exportWorkerCode.includes("aqSheet.state = 'hidden'"));
    assert.ok(exportWorkerCode.includes("approvedSheet.state = 'hidden'"));
    assert.ok(exportWorkerCode.includes("stagingSheet.state = 'hidden'"));
});

runCheck('export-worker: Validate diagnostic tabs are hidden and Review_Workbench is active', () => {
    assert.ok(exportWorkerCode.includes("const hideValidateSheet = (sheetName) =>"));
    assert.ok(exportWorkerCode.includes("'Errors_in_Translate'"));
    assert.ok(exportWorkerCode.includes("'Output_Not_Found_Ambiguous'"));
    assert.ok(exportWorkerCode.includes("'Output_Not_Found_No_Replacement'"));
    assert.ok(exportWorkerCode.includes("'One_to_Many'"));
    assert.ok(exportWorkerCode.includes("'Missing_Mappings'"));
    assert.ok(exportWorkerCode.includes("'High_Confidence_Matches'"));
    assert.ok(exportWorkerCode.includes("'Valid_Mappings'"));
    assert.ok(exportWorkerCode.includes('workbook.views = [{ activeTab: reviewSheetIndex, firstSheet: reviewSheetIndex }]'));
});

runCheck('export-worker: Create review workflow is explicit in Excel', () => {
    assert.ok(exportWorkerCode.includes("addSheetFromObjects('Ambiguous_Candidates'"));
    assert.ok(exportWorkerCode.includes("addSheetFromObjects('Missing_In_myWSU'"));
    assert.ok(exportWorkerCode.includes("addSheetFromObjects('Review_Decisions'"));
    assert.ok(exportWorkerCode.includes("header: 'Resolution Type'"));
    assert.ok(exportWorkerCode.includes("header: 'Review Path'"));
    assert.ok(exportWorkerCode.includes("header: 'Candidate Count'"));
    assert.ok(exportWorkerCode.includes('Review_Instructions_Create'));
});

runCheck('export-worker: Create review sheet highlights unresolved manual rows', () => {
    assert.ok(exportWorkerCode.includes('$${colSourceStatus}2="Ambiguous Match"'));
    assert.ok(exportWorkerCode.includes('$${colSourceStatus}2="Missing in myWSU"'));
    assert.ok(exportWorkerCode.includes('AND($${colDecision}2="",OR($${colSourceStatus}2="Ambiguous Match",$${colSourceStatus}2="Missing in myWSU"))'));
});

runCheck('export-worker: Validate workbook omits Review_Instructions tab', () => {
    assert.ok(!exportWorkerCode.includes("addWorksheet('Review_Instructions'"));
});

runCheck('export-worker: Review_Workbench has freeze and conditional formatting', () => {
    assert.ok(exportWorkerCode.includes('addConditionalFormatting'));
    assert.ok(exportWorkerCode.includes('hiddenReviewColumns'));
});

runCheck('export-worker: Final_Translation_Table includes reviewer context and direct reviewer pull-through', () => {
    assert.ok(
        exportWorkerCode.includes('finalOutcomesCols') || exportWorkerCode.includes("header: 'Outcomes Name'"),
        'Final table should include outcomes context columns'
    );
    assert.ok(
        exportWorkerCode.includes('finalWsuCols') || exportWorkerCode.includes("header: 'myWSU Name'"),
        'Final table should include myWSU context columns'
    );
    assert.ok(exportWorkerCode.includes("header: 'Current Translate Input'"));
    assert.ok(
        exportWorkerCode.includes('const buildFinalAutoRow = (row) =>'),
        'Final table should write auto-approved rows as direct values'
    );
    assert.ok(
        exportWorkerCode.includes('const reviewFinalValueFormula = (sourceKey, rowNum) =>'),
        'Final table should pull review decisions directly from Review_Workbench'
    );
    assert.ok(
        !exportWorkerCode.includes("header: '_Approved Pick'"),
        'Final table should not require AGGREGATE helper pick columns'
    );
    assert.ok(
        exportWorkerCode.includes('finalSheet.autoFilter = {'),
        'Final table should enable an autoFilter range'
    );
});

runCheck('export-worker: workbook enforces recalc on open', () => {
    assert.ok(exportWorkerCode.includes('fullCalcOnLoad: true'));
});

runCheck('export-worker: expression CF rules use formulae array not formula string', () => {
    // ExcelJS renderExpression does model.formulae[0] without guard.
    // Using singular `formula:` on expression rules causes TypeError.
    const expressionBlocks = exportWorkerCode.split("type: 'expression'");
    // First segment is before the first match, skip it
    for (let i = 1; i < expressionBlocks.length; i += 1) {
        const block = expressionBlocks[i].slice(0, 200);
        assert.ok(
            block.includes('formulae:'),
            `expression-type CF rule #${i} must use 'formulae:' (array), not 'formula:' (string)`
        );
        assert.ok(
            !block.match(/\bformula\s*:/),
            `expression-type CF rule #${i} must not use singular 'formula:' property`
        );
    }
});

if (failures > 0) {
    console.error(`\n${failures} check(s) failed.`);
    process.exit(1);
}

console.log('\nAll validate-translation-table checks passed.');
