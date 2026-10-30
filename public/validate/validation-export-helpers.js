/**
 * Shared helpers for validation export (Action_Queue priority, recommended actions).
 * Used by export-worker.js and testable from checks.js.
 */
'use strict';

(function (global) {
    const OUTPUT_NOT_FOUND_SUBTYPE = (typeof global !== 'undefined' && global.OUTPUT_NOT_FOUND_SUBTYPE)
        ? global.OUTPUT_NOT_FOUND_SUBTYPE
        : {
            LIKELY_STALE_KEY: 'Output_Not_Found_Likely_Stale_Key',
            AMBIGUOUS_REPLACEMENT: 'Output_Not_Found_Ambiguous_Replacement',
            NO_REPLACEMENT: 'Output_Not_Found_No_Replacement'
        };

    /** Priority order: lower number = higher priority (tackle first). */
    const PRIORITY_ORDER = {
        Missing_Input: 1,
        Missing_Output: 2,
        Input_Not_Found: 3,
        [OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT]: 4,
        Duplicate_Target: 5,
        Duplicate_Source: 6,
        [OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY]: 7,
        [OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT]: 8,
        Name_Mismatch: 9,
        Ambiguous_Match: 10,
        Missing_Mapping: 11
    };

    const DEFAULT_PRIORITY = 99;

    function getPriority(errorType, errorSubtype) {
        if (errorType === 'Output_Not_Found' && errorSubtype) {
            const sub = PRIORITY_ORDER[errorSubtype];
            if (sub !== undefined) return sub;
        }
        return PRIORITY_ORDER[errorType] ?? DEFAULT_PRIORITY;
    }

    const RECOMMENDED_ACTIONS = {
        Missing_Input: 'Fix blank input in Translate',
        Missing_Output: 'Fix blank output in Translate',
        Input_Not_Found: 'Correct input key or remove row',
        Output_Not_Found: 'Correct output key or remove row',
        [OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT]: 'Verify output key; remove or correct manually',
        [OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY]: 'Update output to suggested key',
        [OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT]: 'Choose correct replacement from candidates',
        Duplicate_Target: 'Resolve many-to-one conflict',
        Duplicate_Source: 'Resolve duplicate source mapping',
        Name_Mismatch: 'Verify name match or correct mapping',
        Ambiguous_Match: 'Choose correct mapping from alternatives',
        Missing_Mapping: 'Add row to Translate table'
    };

    function getRecommendedAction(errorType, errorSubtype) {
        if (errorType === 'Output_Not_Found' && errorSubtype) {
            const act = RECOMMENDED_ACTIONS[errorSubtype];
            if (act) return act;
        }
        return RECOMMENDED_ACTIONS[errorType] ?? 'Review and resolve';
    }

    /** Columns required for Missing_Mappings context in Action_Queue */
    const ACTION_QUEUE_CONTEXT_COLUMNS = ['Missing_In', 'Similarity'];

    /** QA_Checks_Validate rows when Action_Queue is empty - matches non-empty layout for consistency */
    function getQAValidateRowsForEmptyQueue() {
        return [
            ['Check', 'Count', 'Status', 'Detail'],
            ['Unresolved actions', 0, 'PASS', 'Blank or Ignore'],
            ['Approved review rows', 0, 'PASS', 'Rows approved from Review_Workbench'],
            ['Approved rows beyond formula capacity', 0, 'PASS', 'Rows above formula capacity'],
            ['Blank final keys on publish-eligible rows (sanity)', 0, 'PASS', 'Sanity check: publish-eligible rows should already enforce non-blank finals'],
            ['Use Suggestion without effective key', 0, 'PASS', 'Use Suggestion needs Manual Key or Selected Candidate ID + Suggested Key'],
            ['Use Suggestion with invalid Update Side', 0, 'PASS', 'Use Suggestion chosen but Update Side is None; fix or change decision'],
            ['Use Suggestion with invalid manual key', 0, 'PASS', 'Manual key not found in valid myWSU/Outcomes keys'],
            ['Use Suggestion no-op (key equals current value)', 0, 'PASS', 'Use Suggestion chosen but effective key equals current; fix or change decision'],
            ['Risky decisions without reason code', 0, 'PASS', 'Reason Code required for: Use Suggestion+Manual Key, Allow One-to-Many, Duplicate_Target+Keep As-Is'],
            ['Stale-key rows lacking decision', 0, 'PASS', 'Likely stale key rows without decision (advisory)'],
            ['One-to-many rows lacking decision', 0, 'PASS', 'One-to-many rows without decision (advisory)'],
            ['Duplicate final input keys (excluding Allow One-to-Many)', 0, 'PASS', 'Duplicates in Final_Translation_Table input keys excluding approved one-to-many rows'],
            ['Duplicate final output keys (excluding Allow One-to-Many)', 0, 'PASS', 'Duplicates in Final_Translation_Table output keys excluding Allow One-to-Many and Duplicate_Target+Keep As-Is'],
            ['Duplicate (input, output) pairs in Final_Translation_Table', 0, 'PASS', 'Duplicate (input,output) pairs; set one to Ignore before publish'],
            ['Publish gate', 'PASS', '', 'Final publish gate (B11/B12 advisory). Diagnostic tabs are hidden; right-click tab bar and choose Unhide if needed.']
        ];
    }

    /** Filter errorDataRows for Output_Not_Found subtype (uses raw values) */
    function filterOutputNotFoundBySubtype(rows, subtype) {
        return rows.filter(row =>
            row._rawErrorType === 'Output_Not_Found' && row._rawErrorSubtype === subtype
        );
    }

    const exp = {
        OUTPUT_NOT_FOUND_SUBTYPE,
        PRIORITY_ORDER,
        ACTION_QUEUE_CONTEXT_COLUMNS,
        getPriority,
        getRecommendedAction,
        getQAValidateRowsForEmptyQueue,
        filterOutputNotFoundBySubtype
    };
    if (typeof module !== 'undefined' && module.exports) {
        module.exports = exp;
    } else if (typeof global !== 'undefined') {
        global.ValidationExportHelpers = exp;
    }
})(typeof self !== 'undefined' ? self : typeof global !== 'undefined' ? global : this);
