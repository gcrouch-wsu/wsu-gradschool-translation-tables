/* global normalizeKeyValue, calculateNameSimilarity, similarityRatio, tokenizeName, getInformativeTokens, countriesMatch, statesMatch, hasComparableStateValues */
importScripts('validation.js');
importScripts('validation-export-helpers.js');
importScripts('https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js');

const EXPORT_TOTAL = 100;
const EXPORT_OUTPUT_NOT_FOUND_SUBTYPE = (typeof ValidationExportHelpers !== 'undefined' && ValidationExportHelpers)
    ? ValidationExportHelpers.OUTPUT_NOT_FOUND_SUBTYPE
    : {
        LIKELY_STALE_KEY: 'Output_Not_Found_Likely_Stale_Key',
        AMBIGUOUS_REPLACEMENT: 'Output_Not_Found_Ambiguous_Replacement',
        NO_REPLACEMENT: 'Output_Not_Found_No_Replacement'
    };
const getActionPriority = (typeof ValidationExportHelpers !== 'undefined' && ValidationExportHelpers)
    ? ValidationExportHelpers.getPriority
    : () => 99;
const getRecommendedAction = (typeof ValidationExportHelpers !== 'undefined' && ValidationExportHelpers)
    ? ValidationExportHelpers.getRecommendedAction
    : () => 'Review and resolve';
const MAX_VALIDATE_DYNAMIC_REVIEW_FORMULA_ROWS = 5000;

function reportProgress(stage, processed) {
    self.postMessage({
        type: 'progress',
        stage,
        processed,
        total: EXPORT_TOTAL
    });
}

function toArrayBuffer(bufferLike) {
    if (bufferLike instanceof ArrayBuffer) {
        return bufferLike;
    }
    if (bufferLike?.buffer instanceof ArrayBuffer) {
        const start = bufferLike.byteOffset || 0;
        const end = start + (bufferLike.byteLength || bufferLike.length || 0);
        return bufferLike.buffer.slice(start, end);
    }
    throw new Error('Unable to convert workbook buffer to ArrayBuffer');
}

function downloadSafeFileName(name, fallback) {
    const raw = String(name || '').trim();
    return raw || fallback;
}

function getHeaderFill(col, style) {
    if (col === undefined || col === null) return style.defaultHeaderColor || 'FF1E3A8A';
    if (col === 'Error_Type') return style.errorHeaderColor;
    if (col === 'Error_Subtype') return style.errorHeaderColor;
    if (col === 'Duplicate_Group') return style.groupHeaderColor;
    if (col === 'Mapping_Logic') return style.logicHeaderColor;
    if (['translate_input', 'translate_output'].includes(col)) return style.translateHeaderColor;
    if (col.startsWith('outcomes_')) return style.outcomesHeaderColor;
    if (col.startsWith('wsu_')) return style.wsuHeaderColor;
    if (col.startsWith('Suggested_') || col === 'Suggestion_Score') {
        return style.suggestionHeaderColor;
    }
    return style.defaultHeaderColor || 'FF1E3A8A';
}

function getBodyFill(col, style) {
    if (col === undefined || col === null) return style.defaultBodyColor || 'FFF3F4F6';
    if (col === 'Error_Type') return style.errorBodyColor;
    if (col === 'Error_Subtype') return style.errorBodyColor;
    if (col === 'Duplicate_Group') return style.groupBodyColor;
    if (col === 'Mapping_Logic') return style.logicBodyColor;
    if (['translate_input', 'translate_output'].includes(col)) return style.translateBodyColor;
    if (col.startsWith('outcomes_')) return style.outcomesBodyColor;
    if (col.startsWith('wsu_')) return style.wsuBodyColor;
    if (col.startsWith('Suggested_') || col === 'Suggestion_Score') {
        return style.suggestionBodyColor;
    }
    return style.defaultBodyColor;
}

/** Ensure cell value is ExcelJS-safe (primitive or formula object). Avoids "reading '0'" from raw objects. */
function sanitizeCellValue(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
    if (v && typeof v === 'object' && 'formula' in v) return v;
    return String(v);
}

function columnIndexToLetter(index) {
    let result = '';
    let current = index;
    while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
    }
    return result;
}

function addSheetWithRows(workbook, config) {
    const {
        sheetName,
        outputColumns,
        rows,
        style,
        headers,
        rowBorderByError,
        groupColumn,
        freezeConfig,
        columnLayoutByKey
    } = config;
    const sheet = workbook.addWorksheet(sheetName);
    sheet.addRow(headers);
    headers.forEach((header, idx) => {
        const cell = sheet.getCell(1, idx + 1);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: getHeaderFill(outputColumns[idx], style) }
        };
    });

    const sourceFillByColumn = outputColumns.map(col => ({ argb: getBodyFill(col, style) || 'FFF3F4F6' }));
    const defaultFill = { argb: 'FFF3F4F6' };
    let prevGroup = null;
    const groupBorderColor = 'FF9CA3AF';

    rows.forEach(row => {
        let isNewGroup = false;
        if (groupColumn) {
            const currentGroup = row[groupColumn] || '';
            if (prevGroup !== null && currentGroup !== prevGroup) {
                isNewGroup = true;
            }
            prevGroup = currentGroup;
        }

        const rowData = outputColumns.map(col => sanitizeCellValue(row[col]));
        const excelRow = sheet.addRow(rowData);
        excelRow.eachCell((cell, colNumber) => {
            const fill = sourceFillByColumn[colNumber - 1] ?? sourceFillByColumn[0] ?? defaultFill;
            const argb = (fill && fill.argb) ? fill.argb : 'FFF3F4F6';
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
            if (isNewGroup) {
                const existingBorder = cell.border || {};
                cell.border = {
                    ...existingBorder,
                    top: { style: 'medium', color: { argb: groupBorderColor } }
                };
            }
            const colKey = outputColumns[colNumber - 1];
            if (colKey === 'Suggestion_Score' || colKey === 'Score') {
                cell.numFmt = '0%';
            }
        });

        if (rowBorderByError && row.Error_Type) {
            const borderColor = rowBorderByError[row.Error_Type];
            if (borderColor) {
                const indicatorCell = excelRow.getCell(1);
                const existingBorder = indicatorCell.border || {};
                indicatorCell.border = {
                    ...existingBorder,
                    left: { style: 'medium', color: { argb: borderColor } }
                };
            }
        }
    });

    sheet.views = freezeConfig
        ? [{ state: 'frozen', xSplit: freezeConfig.xSplit || 0, ySplit: freezeConfig.ySplit ?? 1 }]
        : [{ state: 'frozen', ySplit: 1 }];
    sheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headers.length }
    };
    const layoutByKey = columnLayoutByKey || {};
    sheet.columns.forEach((column, idx) => {
        const colKey = outputColumns[idx];
        const layout = colKey ? layoutByKey[colKey] : null;
        if (layout && typeof layout.width === 'number') {
            column.width = layout.width;
        } else {
            const headerStr = String(headers[idx] != null ? headers[idx] : '');
            let maxLength = headerStr.length;
            rows.forEach(row => {
                const val = colKey != null ? row[colKey] : '';
                const value = String(sanitizeCellValue(val));
                if (value.length > maxLength) maxLength = value.length;
            });
            column.width = Math.min(maxLength + 2, 70);
        }
        if (layout && Object.prototype.hasOwnProperty.call(layout, 'hidden')) {
            column.hidden = Boolean(layout.hidden);
        }
    });

    const suggestionIndex = outputColumns.indexOf('Suggestion_Score');
    if (suggestionIndex >= 0 && rows.length > 0) {
        const columnLetter = columnIndexToLetter(suggestionIndex + 1);
        const ref = `${columnLetter}2:${columnLetter}${sheet.rowCount}`;
        sheet.addConditionalFormatting({
            ref,
            rules: [
                {
                    type: 'colorScale',
                    cfvo: [
                        { type: 'min' },
                        { type: 'percentile', value: 50 },
                        { type: 'max' }
                    ],
                    color: [
                        { argb: 'FFF87171' },
                        { argb: 'FFFACC15' },
                        { argb: 'FF4ADE80' }
                    ]
                }
            ]
        });
    }
}

function formatScore(score) {
    if (!Number.isFinite(score)) return '';
    return `${Math.round(Math.max(0, Math.min(1, score)) * 100)}%`;
}

function normalizeErrorType(row) {
    if (row.Error_Type === 'Input_Not_Found') return 'Input key not found in Outcomes';
    if (row.Error_Type === 'Output_Not_Found') return 'Output key not found in myWSU';
    if (row.Error_Type === 'Missing_Input') return 'Input key is blank in Translate';
    if (row.Error_Type === 'Missing_Output') return 'Output key is blank in Translate';
    if (row.Error_Type === 'Name_Mismatch') return 'Name mismatch';
    if (row.Error_Type === 'Ambiguous_Match') return 'Ambiguous name match';
    return row.Error_Type || '';
}

function normalizeErrorSubtype(subtype) {
    if (subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY) {
        return 'Likely stale key (high-confidence replacement found)';
    }
    if (subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT) {
        return 'Ambiguous replacement (multiple high-confidence candidates)';
    }
    if (subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT) {
        return 'No high-confidence replacement found';
    }
    return subtype || '';
}

function buildHeaders(outputColumns, keyLabels) {
    return outputColumns.map(col => {
        if (col === 'translate_input') return `${keyLabels.translateInput || 'Source key'} (Translate Input)`;
        if (col === 'translate_output') return `${keyLabels.translateOutput || 'Target key'} (Translate Output)`;
        if (col === 'outcomes_name') return 'School Name (Outcomes)';
        if (col === 'outcomes_school') return 'Outcomes Name';
        if (col === 'wsu_school') return 'myWSU Name';
        if (col === `outcomes_${keyLabels.outcomes}` && keyLabels.outcomes) {
            return `${keyLabels.outcomes} (Outcomes Key)`;
        }
        if (col === 'outcomes_mdb_code') return 'Outcomes Key';
        if (col === 'wsu_Descr') return 'Organization Name (myWSU)';
        if (col === `wsu_${keyLabels.wsu}` && keyLabels.wsu) {
            return `${keyLabels.wsu} (myWSU Key)`;
        }
        if (col === 'wsu_Org ID') return 'myWSU Key';
        if (col === 'Error_Type') return 'Error Type';
        if (col === 'Error_Subtype') return 'Error Subtype';
        if (col === 'Missing_In') return 'Missing In';
        if (col === 'Similarity') return 'Similarity';
        if (col === 'Mapping_Logic') return 'Mapping Logic';
        if (col === 'Suggested_Key') return 'Suggested Key';
        if (col === 'Suggested_School') return 'Suggested School';
        if (col === 'Suggested_City') return 'Suggested City';
        if (col === 'Suggested_State') return 'Suggested State';
        if (col === 'Suggested_Country') return 'Suggested Country';
        if (col === 'Suggestion_Score') return 'Suggestion Score';
        return col;
    });
}

function buildMappingLogicRow(row, normalizedErrorType, nameCompareConfig, idfTable = null) {
    const threshold = Number.isFinite(nameCompareConfig?.threshold)
        ? nameCompareConfig.threshold
        : 0.8;
    const outcomesNameKey = nameCompareConfig?.outcomes ? `outcomes_${nameCompareConfig.outcomes}` : '';
    const wsuNameKey = nameCompareConfig?.wsu ? `wsu_${nameCompareConfig.wsu}` : '';
    const outcomesName = outcomesNameKey ? (row[outcomesNameKey] || row.outcomes_name || '') : (row.outcomes_name || '');
    const wsuName = wsuNameKey ? (row[wsuNameKey] || row.wsu_Descr || '') : (row.wsu_Descr || '');
    const similarity = (outcomesName && wsuName && typeof calculateNameSimilarity === 'function')
        ? calculateNameSimilarity(outcomesName, wsuName, idfTable)
        : null;
    const similarityText = Number.isFinite(similarity) ? formatScore(similarity) : '';
    const thresholdText = formatScore(threshold);

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
        if (row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY) {
            const suggested = row.Suggested_Key
                ? ` Suggested replacement key: ${row.Suggested_Key}.`
                : '';
            return `Key lookup failed: translate output key was not found in myWSU keys. Likely stale key.${suggested}`;
        }
        if (row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT) {
            return 'Key lookup failed: translate output key was not found in myWSU keys. Multiple high-confidence replacement candidates were found.';
        }
        if (row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT) {
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
        return normalizedErrorType
            ? `Classified as "${normalizedErrorType}" by validation rules.`
            : (row.Error_Description || 'Classified by validation rules.');
    }
}

async function buildGenerationExport(payload) {
    const {
        cleanRows = [],
        errorRows = [],
        selectedCols = { outcomes: [], wsu_org: [] },
        generationConfig = {}
    } = payload;
    const workbook = new ExcelJS.Workbook();
    workbook.calcProperties = {
        ...(workbook.calcProperties || {}),
        fullCalcOnLoad: true
    };
    const outcomesCols = selectedCols.outcomes || [];
    const wsuCols = selectedCols.wsu_org || [];
    const threshold = Number.isFinite(generationConfig.threshold)
        ? Math.max(0, Math.min(1, generationConfig.threshold))
        : 0.8;

    const headerColor = {
        meta: 'FF1E3A8A',
        outcomes: 'FF166534',
        wsu: 'FFC2410C',
        decision: 'FF7C3AED',
        qa: 'FF0F766E'
    };

    const toText = (value) => (value === null || value === undefined ? '' : String(value).trim());
    const parseSimilarity = (value) => {
        const numeric = Number(value);
        return Number.isFinite(numeric) ? numeric : null;
    };
    const confidenceOrder = { high: 3, medium: 2, low: 1, '': 0 };
    const normalizeTier = (row) => {
        const raw = String(row.confidence_tier || '').trim().toLowerCase();
        if (raw === 'high' || raw === 'medium' || raw === 'low') {
            return raw;
        }
        const similarity = parseSimilarity(row.match_similarity);
        if (!Number.isFinite(similarity)) {
            return '';
        }
        // `threshold` is 0-1; `match_similarity` is exported as 0-100.
        if (similarity >= 90) return 'high';
        if (similarity >= 80) return 'medium';
        if (similarity >= threshold * 100) return 'low';
        return 'low';
    };
    const getOutcomeSortName = (row) => {
        const direct = toText(row.outcomes_display_name);
        if (direct) return direct.toLowerCase();
        for (const col of outcomesCols) {
            const value = toText(row[`outcomes_${col}`]);
            if (value) return value.toLowerCase();
        }
        return '';
    };
    const getResolutionTypeForErrorRow = (row) => {
        if (row.missing_in === 'Ambiguous Match') return 'Ambiguous candidate set';
        if (row.missing_in === 'myWSU') return 'Missing in myWSU';
        if (row.missing_in === 'Outcomes') return 'Missing in Outcomes';
        return 'Matched 1:1';
    };
    const getAmbiguityScope = (row) => {
        const hasOutcomesRecord = Boolean(toText(row.outcomes_record_id));
        const hasWsuRecord = Boolean(toText(row.wsu_record_id));
        if (hasOutcomesRecord && !hasWsuRecord) {
            return 'Outcomes row with multiple myWSU candidates';
        }
        if (!hasOutcomesRecord && hasWsuRecord) {
            return 'myWSU row with multiple Outcomes candidates';
        }
        return 'Multiple plausible matches';
    };
    const getReviewPathForErrorRow = (row) => {
        if (row.missing_in === 'Ambiguous Match') {
            return 'Review Alt 1-3 and choose best candidate, or mark No Match';
        }
        if (row.missing_in === 'myWSU') {
            return 'No myWSU candidate found. Research target key, then set Decision';
        }
        if (row.missing_in === 'Outcomes') {
            return 'Reference only: no Outcomes source row for this myWSU record';
        }
        return 'Verify and confirm 1:1 mapping';
    };
    const getReviewPathForSourceStatus = (sourceStatus) => {
        if (sourceStatus === 'Ambiguous Match') {
            return 'Choose Alternate (Alt 1-3) or mark No Match';
        }
        if (sourceStatus === 'Missing in myWSU') {
            return 'No candidate found. Research key, then set No Match until resolved';
        }
        return 'Verify proposed match and confirm Decision';
    };

    const setSheetFormats = (sheet, columns, rows, options = {}) => {
        const freezeHeader = options.freezeHeader !== false;
        const autoFilter = options.autoFilter !== false;
        const maxWidth = options.maxWidth || 70;
        if (freezeHeader) {
            sheet.views = [{ state: 'frozen', ySplit: 1 }];
        }
        if (autoFilter && columns.length > 0) {
            sheet.autoFilter = {
                from: { row: 1, column: 1 },
                to: { row: 1, column: columns.length }
            };
        }
        sheet.columns.forEach((column, idx) => {
            const header = columns[idx]?.header || '';
            let maxLength = String(header).length;
            rows.forEach(row => {
                const key = columns[idx]?.key;
                const value = row?.[key];
                const length = String(value?.formula || value || '').length;
                if (length > maxLength) {
                    maxLength = length;
                }
            });
            column.width = Math.min(maxLength + 2, maxWidth);
        });
    };

    const addSheetFromObjects = (sheetName, columns, rows) => {
        const sheet = workbook.addWorksheet(sheetName);
        sheet.addRow(columns.map(col => col.header));
        sheet.getRow(1).eachCell((cell, idx) => {
            const col = columns[idx - 1] || {};
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: headerColor[col.group || 'meta'] || headerColor.meta }
            };
        });

        rows.forEach(row => {
            const rowData = columns.map(col => sanitizeCellValue(row[col.key]));
            sheet.addRow(rowData);
        });

        setSheetFormats(sheet, columns, rows);
        return sheet;
    };

    const cleanSorted = [...cleanRows]
        .map(row => ({
            ...row,
            confidence_tier: normalizeTier(row)
        }))
        .sort((a, b) => {
            const tierDiff = (confidenceOrder[b.confidence_tier] || 0) - (confidenceOrder[a.confidence_tier] || 0);
            if (tierDiff !== 0) return tierDiff;
            const simA = parseSimilarity(a.match_similarity);
            const simB = parseSimilarity(b.match_similarity);
            const simDiff = (Number.isFinite(simB) ? simB : -1) - (Number.isFinite(simA) ? simA : -1);
            if (simDiff !== 0) return simDiff;
            return getOutcomeSortName(a).localeCompare(getOutcomeSortName(b));
        });

    const ambiguousRows = errorRows.filter(row => row.missing_in === 'Ambiguous Match');
    const missingInMyWsuRows = errorRows.filter(row => row.missing_in === 'myWSU');
    const missingInOutcomesRows = errorRows.filter(row => row.missing_in === 'Outcomes');
    const ambiguousOutcomesCount = ambiguousRows.filter(
        row => row.outcomes_row_index !== '' && row.outcomes_row_index !== null && row.outcomes_row_index !== undefined
    ).length;

    reportProgress('Building summary sheet...', 10);
    const summarySheet = workbook.addWorksheet('Summary');
    const summaryRows = [
        ['Metric', 'Count', 'Notes'],
        ['Matched 1:1 rows', cleanSorted.length, 'Rows with a single assigned match'],
        ['Ambiguous candidates', ambiguousRows.length, 'Rows with multiple plausible matches; choose from Alt candidates'],
        ['Missing in myWSU', missingInMyWsuRows.length, 'Outcomes rows with no candidate target match in myWSU'],
        ['Missing in Outcomes', missingInOutcomesRows.length, 'myWSU rows without a source row'],
        ['Review_Decisions rows', cleanSorted.length + missingInMyWsuRows.length + ambiguousOutcomesCount, 'One row per Outcomes record'],
        ['Workflow note', '', 'Use Review_Decisions for actions. Ambiguous rows need candidate selection; Missing in myWSU rows need research/manual mapping.']
    ];
    summaryRows.forEach(row => summarySheet.addRow(row));
    summarySheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor.meta } };
    });
    summarySheet.columns = [
        { width: 34 },
        { width: 16 },
        { width: 70 }
    ];

    reportProgress('Building matched sheet...', 22);
    const newTranslationColumns = [
        { key: 'outcomes_record_id', header: 'Outcomes Record ID', group: 'meta' },
        { key: 'outcomes_display_name', header: 'Outcomes Name', group: 'meta' },
        { key: 'proposed_wsu_key', header: 'Proposed myWSU Key', group: 'meta' },
        { key: 'proposed_wsu_name', header: 'Proposed myWSU Name', group: 'meta' },
        { key: 'match_similarity', header: 'Similarity %', group: 'meta' },
        { key: 'confidence_tier', header: 'Confidence Tier', group: 'meta' }
    ];
    outcomesCols.forEach(col => {
        newTranslationColumns.push({ key: `outcomes_${col}`, header: `Outcomes: ${col}`, group: 'outcomes' });
    });
    wsuCols.forEach(col => {
        newTranslationColumns.push({ key: `wsu_${col}`, header: `myWSU: ${col}`, group: 'wsu' });
    });
    addSheetFromObjects('New_Translation_Candidates', newTranslationColumns, cleanSorted);

    reportProgress('Building ambiguous sheet...', 34);
    const ambiguousRowsMapped = ambiguousRows.map(row => {
        const alternates = Array.isArray(row.alternate_candidates) ? row.alternate_candidates : [];
        const alt = (index) => alternates[index] || {};
        return {
            ...row,
            resolution_type: getResolutionTypeForErrorRow(row),
            review_path: getReviewPathForErrorRow(row),
            ambiguity_scope: getAmbiguityScope(row),
            candidate_count: alternates.length,
            proposed_wsu_key: row.proposed_wsu_key || '',
            proposed_wsu_name: row.proposed_wsu_name || '',
            Alt_1_Key: alt(0).key || '',
            Alt_1_Name: alt(0).name || '',
            Alt_1_Similarity: alt(0).similarity ?? '',
            Alt_2_Key: alt(1).key || '',
            Alt_2_Name: alt(1).name || '',
            Alt_2_Similarity: alt(1).similarity ?? '',
            Alt_3_Key: alt(2).key || '',
            Alt_3_Name: alt(2).name || '',
            Alt_3_Similarity: alt(2).similarity ?? ''
        };
    });
    const ambiguousColumns = [
        { key: 'outcomes_record_id', header: 'Outcomes Record ID', group: 'meta' },
        { key: 'outcomes_display_name', header: 'Outcomes Name', group: 'meta' },
        { key: 'missing_in', header: 'Missing In', group: 'meta' },
        { key: 'resolution_type', header: 'Resolution Type', group: 'meta' },
        { key: 'review_path', header: 'Review Path', group: 'meta' },
        { key: 'ambiguity_scope', header: 'Ambiguity Scope', group: 'meta' },
        { key: 'candidate_count', header: 'Candidate Count', group: 'meta' },
        { key: 'proposed_wsu_key', header: 'Top Suggested myWSU Key', group: 'decision' },
        { key: 'proposed_wsu_name', header: 'Top Suggested myWSU Name', group: 'decision' },
        { key: 'Alt_1_Key', header: 'Alt 1 Key', group: 'decision' },
        { key: 'Alt_1_Name', header: 'Alt 1 Name', group: 'decision' },
        { key: 'Alt_1_Similarity', header: 'Alt 1 Similarity %', group: 'decision' },
        { key: 'Alt_2_Key', header: 'Alt 2 Key', group: 'decision' },
        { key: 'Alt_2_Name', header: 'Alt 2 Name', group: 'decision' },
        { key: 'Alt_2_Similarity', header: 'Alt 2 Similarity %', group: 'decision' },
        { key: 'Alt_3_Key', header: 'Alt 3 Key', group: 'decision' },
        { key: 'Alt_3_Name', header: 'Alt 3 Name', group: 'decision' },
        { key: 'Alt_3_Similarity', header: 'Alt 3 Similarity %', group: 'decision' }
    ];
    outcomesCols.forEach(col => {
        ambiguousColumns.push({ key: `outcomes_${col}`, header: `Outcomes: ${col}`, group: 'outcomes' });
    });
    wsuCols.forEach(col => {
        ambiguousColumns.push({ key: `wsu_${col}`, header: `myWSU: ${col}`, group: 'wsu' });
    });
    addSheetFromObjects('Ambiguous_Candidates', ambiguousColumns, ambiguousRowsMapped);

    reportProgress('Building missing sheets...', 46);
    const missingInMyWsuRowsMapped = missingInMyWsuRows.map(row => ({
        ...row,
        resolution_type: getResolutionTypeForErrorRow(row),
        review_path: getReviewPathForErrorRow(row)
    }));
    const missingMyWsuColumns = [
        { key: 'outcomes_record_id', header: 'Outcomes Record ID', group: 'meta' },
        { key: 'outcomes_display_name', header: 'Outcomes Name', group: 'meta' },
        { key: 'missing_in', header: 'Missing In', group: 'meta' },
        { key: 'resolution_type', header: 'Resolution Type', group: 'meta' },
        { key: 'review_path', header: 'Review Path', group: 'meta' }
    ];
    outcomesCols.forEach(col => {
        missingMyWsuColumns.push({ key: `outcomes_${col}`, header: `Outcomes: ${col}`, group: 'outcomes' });
    });
    addSheetFromObjects('Missing_In_myWSU', missingMyWsuColumns, missingInMyWsuRowsMapped);

    const missingInOutcomesRowsMapped = missingInOutcomesRows.map(row => ({
        ...row,
        resolution_type: getResolutionTypeForErrorRow(row),
        review_path: getReviewPathForErrorRow(row)
    }));
    const missingOutcomesColumns = [
        { key: 'wsu_record_id', header: 'myWSU Record ID', group: 'meta' },
        { key: 'wsu_display_name', header: 'myWSU Name', group: 'meta' },
        { key: 'missing_in', header: 'Missing In', group: 'meta' },
        { key: 'resolution_type', header: 'Resolution Type', group: 'meta' },
        { key: 'review_path', header: 'Review Path', group: 'meta' }
    ];
    wsuCols.forEach(col => {
        missingOutcomesColumns.push({ key: `wsu_${col}`, header: `myWSU: ${col}`, group: 'wsu' });
    });
    addSheetFromObjects('Missing_In_Outcomes', missingOutcomesColumns, missingInOutcomesRowsMapped);

    reportProgress('Building review decisions...', 60);
    const reviewIndexMap = new Map();
    cleanSorted.forEach(row => {
        if (row.outcomes_row_index === '' || row.outcomes_row_index === null || row.outcomes_row_index === undefined) {
            return;
        }
        const idx = Number(row.outcomes_row_index);
        if (!Number.isFinite(idx)) return;
        reviewIndexMap.set(idx, {
            ...row,
            source_status: 'Matched',
            resolution_type: 'Matched 1:1',
            review_path: getReviewPathForSourceStatus('Matched'),
            confidence_tier: normalizeTier(row)
        });
    });
    errorRows.forEach(row => {
        if (row.outcomes_row_index === '' || row.outcomes_row_index === null || row.outcomes_row_index === undefined) {
            return;
        }
        const idx = Number(row.outcomes_row_index);
        if (!Number.isFinite(idx) || reviewIndexMap.has(idx)) {
            return;
        }
        reviewIndexMap.set(idx, {
            ...row,
            source_status: row.missing_in === 'Ambiguous Match' ? 'Ambiguous Match' : 'Missing in myWSU',
            resolution_type: getResolutionTypeForErrorRow(row),
            review_path: getReviewPathForSourceStatus(row.missing_in === 'Ambiguous Match' ? 'Ambiguous Match' : 'Missing in myWSU'),
            confidence_tier: normalizeTier(row)
        });
    });
    const reviewRows = Array.from(reviewIndexMap.entries())
        .sort((a, b) => a[0] - b[0])
        .map(([, row]) => {
            const alternates = Array.isArray(row.alternate_candidates) ? row.alternate_candidates : [];
            const alt = (index) => alternates[index] || {};
            const defaultDecision = row.source_status === 'Matched' && row.confidence_tier === 'high'
                ? 'Accept'
                : '';
            return {
                outcomes_row_index: row.outcomes_row_index,
                outcomes_record_id: row.outcomes_record_id || '',
                outcomes_display_name: row.outcomes_display_name || '',
                proposed_wsu_key: row.proposed_wsu_key || '',
                proposed_wsu_name: row.proposed_wsu_name || '',
                match_similarity: row.match_similarity ?? '',
                confidence_tier: row.confidence_tier || '',
                source_status: row.source_status || '',
                resolution_type: row.resolution_type || '',
                review_path: row.review_path || '',
                candidate_count: alternates.length,
                Alt_1_Key: alt(0).key || '',
                Alt_1_Name: alt(0).name || '',
                Alt_1_Similarity: alt(0).similarity ?? '',
                Alt_2_Key: alt(1).key || '',
                Alt_2_Name: alt(1).name || '',
                Alt_2_Similarity: alt(1).similarity ?? '',
                Alt_3_Key: alt(2).key || '',
                Alt_3_Name: alt(2).name || '',
                Alt_3_Similarity: alt(2).similarity ?? '',
                Decision: defaultDecision,
                Alternate_Choice: '',
                Final_myWSU_Key: '',
                Final_myWSU_Name: '',
                Reason_Code: '',
                Reviewer: '',
                Review_Date: '',
                Notes: ''
            };
        });

    const reviewColumns = [
        { key: 'outcomes_row_index', header: 'Outcomes Row #', group: 'meta' },
        { key: 'outcomes_record_id', header: 'Outcomes Record ID', group: 'meta' },
        { key: 'outcomes_display_name', header: 'Outcomes Name', group: 'meta' },
        { key: 'proposed_wsu_key', header: 'Proposed myWSU Key', group: 'meta' },
        { key: 'proposed_wsu_name', header: 'Proposed myWSU Name', group: 'meta' },
        { key: 'match_similarity', header: 'Similarity %', group: 'meta' },
        { key: 'confidence_tier', header: 'Confidence Tier', group: 'meta' },
        { key: 'source_status', header: 'Source Status', group: 'meta' },
        { key: 'resolution_type', header: 'Resolution Type', group: 'meta' },
        { key: 'review_path', header: 'Review Path', group: 'meta' },
        { key: 'candidate_count', header: 'Candidate Count', group: 'meta' },
        { key: 'Alt_1_Key', header: 'Alt 1 Key', group: 'decision' },
        { key: 'Alt_1_Name', header: 'Alt 1 Name', group: 'decision' },
        { key: 'Alt_1_Similarity', header: 'Alt 1 Similarity %', group: 'decision' },
        { key: 'Alt_2_Key', header: 'Alt 2 Key', group: 'decision' },
        { key: 'Alt_2_Name', header: 'Alt 2 Name', group: 'decision' },
        { key: 'Alt_2_Similarity', header: 'Alt 2 Similarity %', group: 'decision' },
        { key: 'Alt_3_Key', header: 'Alt 3 Key', group: 'decision' },
        { key: 'Alt_3_Name', header: 'Alt 3 Name', group: 'decision' },
        { key: 'Alt_3_Similarity', header: 'Alt 3 Similarity %', group: 'decision' },
        { key: 'Decision', header: 'Decision', group: 'decision' },
        { key: 'Alternate_Choice', header: 'Alternate Choice', group: 'decision' },
        { key: 'Final_myWSU_Key', header: 'Final myWSU Key', group: 'decision' },
        { key: 'Final_myWSU_Name', header: 'Final myWSU Name', group: 'decision' },
        { key: '_dup_final_key_count', header: '_Dup Final Key Count', group: 'decision' },
        { key: 'Reason_Code', header: 'Reason Code', group: 'decision' },
        { key: 'Reviewer', header: 'Reviewer', group: 'decision' },
        { key: 'Review_Date', header: 'Review Date', group: 'decision' },
        { key: 'Notes', header: 'Notes', group: 'decision' }
    ];
    const reviewSheet = addSheetFromObjects('Review_Decisions', reviewColumns, reviewRows);

    const reviewColIndexByKey = {};
    reviewColumns.forEach((col, idx) => {
        reviewColIndexByKey[col.key] = idx + 1;
    });
    const reviewColLetterByKey = {};
    Object.keys(reviewColIndexByKey).forEach(key => {
        reviewColLetterByKey[key] = columnIndexToLetter(reviewColIndexByKey[key]);
    });

    const colDecision = reviewColLetterByKey.Decision;
    const colSourceStatus = reviewColLetterByKey.source_status;
    const colAltChoice = reviewColLetterByKey.Alternate_Choice;
    const colPropKey = reviewColLetterByKey.proposed_wsu_key;
    const colPropName = reviewColLetterByKey.proposed_wsu_name;
    const colAlt1Key = reviewColLetterByKey.Alt_1_Key;
    const colAlt1Name = reviewColLetterByKey.Alt_1_Name;
    const colAlt2Key = reviewColLetterByKey.Alt_2_Key;
    const colAlt2Name = reviewColLetterByKey.Alt_2_Name;
    const colAlt3Key = reviewColLetterByKey.Alt_3_Key;
    const colAlt3Name = reviewColLetterByKey.Alt_3_Name;
    const colFinalKey = reviewColLetterByKey.Final_myWSU_Key;
    const colFinalName = reviewColLetterByKey.Final_myWSU_Name;
    const colDupFinalKeyCount = reviewColLetterByKey._dup_final_key_count;
    const reviewDecisionLastRow = Math.max(2, reviewSheet.rowCount);
    const reviewDecisionFinalKeyRange = `$${colFinalKey}$2:$${colFinalKey}$${reviewDecisionLastRow}`;
    for (let rowNum = 2; rowNum <= reviewSheet.rowCount; rowNum += 1) {
        reviewSheet.getCell(`${colFinalKey}${rowNum}`).value = {
            formula: `IF($${colDecision}${rowNum}="Accept",$${colPropKey}${rowNum},IF($${colDecision}${rowNum}="Choose Alternate",IF($${colAltChoice}${rowNum}="Alt 1",$${colAlt1Key}${rowNum},IF($${colAltChoice}${rowNum}="Alt 2",$${colAlt2Key}${rowNum},IF($${colAltChoice}${rowNum}="Alt 3",$${colAlt3Key}${rowNum},""))),""))`
        };
        reviewSheet.getCell(`${colFinalName}${rowNum}`).value = {
            formula: `IF($${colDecision}${rowNum}="Accept",$${colPropName}${rowNum},IF($${colDecision}${rowNum}="Choose Alternate",IF($${colAltChoice}${rowNum}="Alt 1",$${colAlt1Name}${rowNum},IF($${colAltChoice}${rowNum}="Alt 2",$${colAlt2Name}${rowNum},IF($${colAltChoice}${rowNum}="Alt 3",$${colAlt3Name}${rowNum},""))),""))`
        };
        reviewSheet.getCell(`${colDupFinalKeyCount}${rowNum}`).value = {
            formula: `IF($${colFinalKey}${rowNum}="","",COUNTIF(${reviewDecisionFinalKeyRange},$${colFinalKey}${rowNum}))`
        };
    }
    const dupCountColumn = reviewSheet.getColumn(reviewColIndexByKey._dup_final_key_count);
    if (dupCountColumn) {
        dupCountColumn.hidden = true;
        dupCountColumn.width = 4;
    }
    if (reviewSheet.rowCount > 1) {
        reviewSheet.dataValidations.add(
            `${colDecision}2:${colDecision}${reviewSheet.rowCount}`,
            {
                type: 'list',
                allowBlank: true,
                formulae: ['"Accept,Choose Alternate,No Match"']
            }
        );
        reviewSheet.dataValidations.add(
            `${colAltChoice}2:${colAltChoice}${reviewSheet.rowCount}`,
            {
                type: 'list',
                allowBlank: true,
                formulae: ['"Alt 1,Alt 2,Alt 3"']
            }
        );
        const reviewLastColLetter = columnIndexToLetter(reviewColumns.length);
        const reviewDataRef = `A2:${reviewLastColLetter}${reviewSheet.rowCount}`;
        reviewSheet.addConditionalFormatting({
            ref: reviewDataRef,
            rules: [
                {
                    type: 'expression',
                    formulae: [`$${colSourceStatus}2="Ambiguous Match"`],
                    style: {
                        fill: {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFEF3C7' }
                        }
                    }
                },
                {
                    type: 'expression',
                    formulae: [`$${colSourceStatus}2="Missing in myWSU"`],
                    style: {
                        fill: {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFEE2E2' }
                        }
                    }
                },
                {
                    type: 'expression',
                    formulae: [`AND($${colDecision}2="",OR($${colSourceStatus}2="Ambiguous Match",$${colSourceStatus}2="Missing in myWSU"))`],
                    style: {
                        border: {
                            left: { style: 'medium', color: { argb: 'FFF59E0B' } }
                        }
                    }
                }
            ]
        });
    }

    reportProgress('Building final translation sheet...', 74);
    const finalColumns = [
        { key: 'outcomes_record_id', header: 'Outcomes Record ID', group: 'meta' },
        { key: 'outcomes_display_name', header: 'Outcomes Name', group: 'meta' },
        { key: 'final_wsu_key', header: 'Final myWSU Key', group: 'meta' },
        { key: 'final_wsu_name', header: 'Final myWSU Name', group: 'meta' },
        { key: 'decision', header: 'Decision', group: 'decision' },
        { key: 'confidence_tier', header: 'Confidence Tier', group: 'meta' },
        { key: 'similarity', header: 'Similarity %', group: 'meta' }
    ];
    const finalSheet = workbook.addWorksheet('Final_Translation_Table');
    finalSheet.addRow(finalColumns.map(col => col.header));
    finalSheet.getRow(1).eachCell((cell, idx) => {
        const col = finalColumns[idx - 1];
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: headerColor[col.group || 'meta'] || headerColor.meta }
        };
    });
    const reviewLastRow = Math.max(2, reviewSheet.rowCount);
    const approvedMask = `( (Review_Decisions!$${colDecision}$2:$${colDecision}$${reviewLastRow}="Accept") + (Review_Decisions!$${colDecision}$2:$${colDecision}$${reviewLastRow}="Choose Alternate") )`;
    const approvedRelativeRows = `(ROW(Review_Decisions!$${colDecision}$2:$${colDecision}$${reviewLastRow})-ROW(Review_Decisions!$${colDecision}$2)+1)`;
    const approvedPick = (k) => `AGGREGATE(15,6,${approvedRelativeRows}/(${approvedMask}),${k})`;
    const indexApproved = (colLetter, k) => (
        `IFERROR(INDEX(Review_Decisions!$${colLetter}$2:$${colLetter}$${reviewLastRow},${approvedPick(k)}),"")`
    );
    const finalFormulaRows = Math.max(1, reviewSheet.rowCount - 1);
    for (let outputIndex = 1; outputIndex <= finalFormulaRows; outputIndex += 1) {
        finalSheet.addRow([
            { formula: indexApproved(reviewColLetterByKey.outcomes_record_id, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.outcomes_display_name, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.Final_myWSU_Key, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.Final_myWSU_Name, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.Decision, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.confidence_tier, outputIndex) },
            { formula: indexApproved(reviewColLetterByKey.match_similarity, outputIndex) }
        ]);
    }
    setSheetFormats(finalSheet, finalColumns, [], { maxWidth: 60 });

    reportProgress('Building QA sheet...', 86);
    const qaSheet = workbook.addWorksheet('QA_Checks');
    const decisionRange = `Review_Decisions!$${colDecision}$2:$${colDecision}$${reviewLastRow}`;
    const reviewFinalKeyRange = `Review_Decisions!$${colFinalKey}$2:$${colFinalKey}$${reviewLastRow}`;
    const reviewDupFinalKeyRange = `Review_Decisions!$${colDupFinalKeyCount}$2:$${colDupFinalKeyCount}$${reviewLastRow}`;
    const qaRows = [
        ['Check', 'Count', 'Status', 'Detail'],
        ['Unresolved count', `=COUNTIF(${decisionRange},"")+COUNTIF(${decisionRange},"No Match")`, '=IF(B2=0,"PASS","CHECK")', 'Blank or No Match decisions'],
        ['Blank final key with approved decision', `=COUNTIFS(${decisionRange},"Accept",${reviewFinalKeyRange},"")+COUNTIFS(${decisionRange},"Choose Alternate",${reviewFinalKeyRange},"")`, '=IF(B3=0,"PASS","FAIL")', 'Approved rows should produce final keys'],
        ['Duplicate final keys (approved)', `=SUMPRODUCT(((${decisionRange}="Accept")+(${decisionRange}="Choose Alternate"))*(${reviewFinalKeyRange}<>"")*(${reviewDupFinalKeyRange}>1))/2`, '=IF(B4=0,"PASS","CHECK")', 'Duplicates among approved final keys'],
        ['Approved total', `=COUNTIF(${decisionRange},"Accept")+COUNTIF(${decisionRange},"Choose Alternate")`, '', 'Rows that will publish to final translation table']
    ];
    qaRows.forEach(row => qaSheet.addRow(row));
    qaSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor.qa } };
    });
    qaSheet.columns = [
        { width: 42 },
        { width: 26 },
        { width: 14 },
        { width: 70 }
    ];

    reportProgress('Building review instructions...', 90);
    const createInstructionsSheet = workbook.addWorksheet('Review_Instructions_Create', {
        properties: { tabColor: { argb: 'FF1E3A8A' } }
    });
    const createInstructionsRows = [
        ['Section', 'Guidance'],
        ['Primary action sheet', 'Use Review_Decisions to set Decision values for each Outcomes row.'],
        ['Source Status: Matched', 'Verify the proposed key/name and keep Accept unless correction is needed.'],
        ['Source Status: Ambiguous Match', 'Use Alt 1-3 candidates and Alternate Choice to resolve low-confidence mapping.'],
        ['Source Status: Missing in myWSU', 'No target candidate found. Research key and use No Match until resolved.'],
        ['Ambiguous_Candidates tab', 'Reference tab with candidate options, ambiguity scope, and review path guidance.'],
        ['Missing_In_myWSU tab', 'Reference tab listing Outcomes rows that currently have no myWSU target row.'],
        ['Final_Translation_Table', 'Publishes only Accept and Choose Alternate decisions from Review_Decisions.'],
        ['QA_Checks', 'Review unresolved count, blank finals, and duplicate final keys before publishing.']
    ];
    createInstructionsRows.forEach(row => createInstructionsSheet.addRow(row));
    createInstructionsSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor.meta } };
    });
    createInstructionsSheet.columns = [{ width: 34 }, { width: 110 }];
    createInstructionsSheet.views = [{ state: 'frozen', ySplit: 1 }];

    reportProgress('Finalizing Excel file...', 94);
    const buffer = toArrayBuffer(await workbook.xlsx.writeBuffer());
    reportProgress('Saving file...', 100);
    return {
        buffer,
        filename: 'Generated_Translation_Table.xlsx'
    };
}

async function buildValidationExport(payload) {
    const { validated = [], selectedCols = {}, options = {}, context = {}, priorDecisions: priorDecisionsPayload = null, preEditedActionQueueRows = null, returnActionQueueOnly = false } = payload || {};
    const workbook = new ExcelJS.Workbook();
    workbook.calcProperties = {
        ...(workbook.calcProperties || {}),
        fullCalcOnLoad: true
    };
    const includeSuggestions = Boolean(options.includeSuggestions);
    const showMappingLogic = Boolean(options.showMappingLogic);
    const requestedReviewScope = String(options.reviewScope || '').trim().toLowerCase();
    const reviewScope = (
        requestedReviewScope === 'translation_only' || requestedReviewScope === 'missing_only'
    )
        ? requestedReviewScope
        : (Boolean(options.translationOnlyExport) ? 'translation_only' : 'all');
    const translationOnlyExport = reviewScope === 'translation_only';
    const missingOnlyExport = reviewScope === 'missing_only';
    const nameCompareConfig = options.nameCompareConfig || {};
    const campusFamilyRules = options.campusFamilyRules || null;
    const loadedData = context.loadedData || { outcomes: [], translate: [], wsu_org: [] };
    const columnRoles = context.columnRoles || { outcomes: {}, wsu_org: {} };
    const keyConfig = context.keyConfig || {};
    const keyLabels = context.keyLabels || {};

    reportProgress('Building export...', 5);

    const outcomesColumns = (selectedCols.outcomes || []).map(col => `outcomes_${col}`);
    const wsuColumns = (selectedCols.wsu_org || []).map(col => `wsu_${col}`);
    const mappingColumns = showMappingLogic ? ['Mapping_Logic'] : [];
    const suggestionColumns = includeSuggestions
        ? [
            'Suggested_Key',
            'Suggested_School',
            'Suggested_City',
            'Suggested_State',
            'Suggested_Country',
            'Suggestion_Score'
        ]
        : [];
    const reviewSuggestionColumns = [
        'Suggested_Key',
        'Suggested_School',
        'Suggested_City',
        'Suggested_State',
        'Suggested_Country',
        'Suggestion_Score'
    ];

    const roleOrder = ['School', 'City', 'State', 'Country', 'Other'];
    const getRoleMap = (sourceKey) => {
        const roles = columnRoles[sourceKey] || {};
        const roleMap = {};
        roleOrder.forEach(role => {
            Object.keys(roles).forEach(col => {
                if (roles[col] === role && !roleMap[role]) {
                    roleMap[role] = col;
                }
            });
        });
        return roleMap;
    };

    const roleMapOutcomes = getRoleMap('outcomes');
    const roleMapWsu = getRoleMap('wsu_org');
    const getFallbackRoleColumn = (columns, roleName) => {
        const roleLower = roleName.toLowerCase();
        const hints = roleLower === 'school'
            ? ['school', 'descr', 'name']
            : [roleLower];
        for (const hint of hints) {
            const found = columns.find(col => String(col).toLowerCase().includes(hint));
            if (found) {
                return found;
            }
        }
        return '';
    };
    const getRoleValue = (row, roleMap, fallbackColumns, roleName, prefix) => {
        const roleColumn = roleMap[roleName] || getFallbackRoleColumn(fallbackColumns, roleName);
        if (!roleColumn) return '';
        const keyCandidates = [];
        if (prefix) {
            keyCandidates.push(`${prefix}${roleColumn}`);
        }
        keyCandidates.push(roleColumn);
        if (!prefix) {
            keyCandidates.push(`outcomes_${roleColumn}`, `wsu_${roleColumn}`);
        }
        for (const key of keyCandidates) {
            const value = row?.[key];
            if (value !== undefined && value !== null && value !== '') {
                return value;
            }
        }
        for (const key of keyCandidates) {
            const value = row?.[key];
            if (value !== undefined && value !== null) {
                return value;
            }
        }
        return '';
    };
    const fillSuggestedFields = (rowData, row, roleMap, fallbackColumns, keyValue, prefix) => {
        rowData.Suggested_Key = keyValue || '';
        rowData.Suggested_School = getRoleValue(row, roleMap, fallbackColumns, 'School', prefix);
        rowData.Suggested_City = getRoleValue(row, roleMap, fallbackColumns, 'City', prefix);
        rowData.Suggested_State = getRoleValue(row, roleMap, fallbackColumns, 'State', prefix);
        rowData.Suggested_Country = getRoleValue(row, roleMap, fallbackColumns, 'Country', prefix);
    };

    const normalizeValue = (value) => (
        typeof normalizeKeyValue === 'function'
            ? normalizeKeyValue(value)
            : String(value || '').trim().toLowerCase()
    );
    const similarityScore = (valueA, valueB) => (
        typeof similarityRatio === 'function'
            ? similarityRatio(valueA, valueB)
            : (valueA && valueB && valueA === valueB ? 1 : 0)
    );
    const getNameTokens = (nameValue) => {
        const raw = String(nameValue || '').trim();
        if (!raw) return [];
        if (typeof getInformativeTokens === 'function' && typeof tokenizeName === 'function') {
            return getInformativeTokens(tokenizeName(raw));
        }
        return raw
            .toLowerCase()
            .split(/[^a-z0-9]+/)
            .filter(token => token.length > 1);
    };
    const buildTokenIDFLocal = (allNames) => {
        const names = Array.isArray(allNames) ? allNames : [];
        if (!names.length) return {};
        const df = {};
        names.forEach(name => {
            const tokens = new Set(getNameTokens(name));
            tokens.forEach(token => {
                df[token] = (df[token] || 0) + 1;
            });
        });
        const idf = {};
        const docCount = names.length;
        Object.keys(df).forEach(token => {
            idf[token] = Math.log((docCount + 1) / (df[token] + 1)) + 1;
        });
        const values = Object.values(idf).sort((a, b) => a - b);
        const medianIDF = values.length ? values[Math.floor(values.length / 2)] : 0;
        try {
            Object.defineProperty(idf, '__median', {
                value: medianIDF,
                writable: true,
                enumerable: false,
                configurable: true
            });
        } catch (error) {
            idf.__median = medianIDF;
        }
        return idf;
    };
    const buildTokenIndex = (candidates, idf, minIdf) => {
        const index = {};
        candidates.forEach((candidate, idx) => {
            const tokens = new Set(getNameTokens(candidate.name));
            tokens.forEach(token => {
                if ((idf[token] || 0) >= minIdf) {
                    if (!index[token]) {
                        index[token] = [];
                    }
                    index[token].push(idx);
                }
            });
        });
        return index;
    };
    const getBlockedCandidateIndices = (queryName, tokenIndex, idf, minCandidates = 5, maxTokens = 3) => {
        const queryTokens = Array.from(new Set(getNameTokens(queryName)));
        if (!queryTokens.length) return null;
        queryTokens.sort((a, b) => (idf[b] || 0) - (idf[a] || 0));
        const blocked = new Set();
        queryTokens.slice(0, maxTokens).forEach(token => {
            (tokenIndex[token] || []).forEach(idx => blocked.add(idx));
        });
        if (blocked.size < minCandidates) {
            return null;
        }
        return Array.from(blocked);
    };

    const MIN_KEY_SUGGESTION_SCORE = 0.6;
    const MIN_NAME_SUGGESTION_DISPLAY_SCORE = 0.4;
    const canSuggestNames = Boolean(
        includeSuggestions &&
        nameCompareConfig.enabled &&
        nameCompareConfig.outcomes &&
        nameCompareConfig.wsu
    );
    const outcomesSuggestionCityColumn = nameCompareConfig.city_outcomes || roleMapOutcomes.City || getFallbackRoleColumn(selectedCols.outcomes || [], 'city');
    const outcomesSuggestionStateColumn = nameCompareConfig.state_outcomes || roleMapOutcomes.State || getFallbackRoleColumn(selectedCols.outcomes || [], 'state');
    const outcomesSuggestionCountryColumn = nameCompareConfig.country_outcomes || roleMapOutcomes.Country || getFallbackRoleColumn(selectedCols.outcomes || [], 'country');
    const wsuSuggestionCityColumn = nameCompareConfig.city_wsu || roleMapWsu.City || getFallbackRoleColumn(selectedCols.wsu_org || [], 'city');
    const wsuSuggestionStateColumn = nameCompareConfig.state_wsu || roleMapWsu.State || getFallbackRoleColumn(selectedCols.wsu_org || [], 'state');
    const wsuSuggestionCountryColumn = nameCompareConfig.country_wsu || roleMapWsu.Country || getFallbackRoleColumn(selectedCols.wsu_org || [], 'country');

    const outcomesKeyCandidates = (loadedData.outcomes || [])
        .map(row => ({
            raw: row[keyConfig.outcomes],
            norm: normalizeValue(row[keyConfig.outcomes]),
            row
        }))
        .filter(entry => entry.norm);

    const wsuKeyCandidates = (loadedData.wsu_org || [])
        .map(row => ({
            raw: row[keyConfig.wsu],
            norm: normalizeValue(row[keyConfig.wsu]),
            row
        }))
        .filter(entry => entry.norm);

    const wsuNameCandidates = canSuggestNames
        ? (loadedData.wsu_org || [])
            .map(row => ({
                key: row[keyConfig.wsu],
                name: row[nameCompareConfig.wsu],
                normName: normalizeValue(row[nameCompareConfig.wsu]),
                city: wsuSuggestionCityColumn ? (row[wsuSuggestionCityColumn] ?? '') : '',
                state: wsuSuggestionStateColumn ? (row[wsuSuggestionStateColumn] ?? '') : '',
                country: wsuSuggestionCountryColumn ? (row[wsuSuggestionCountryColumn] ?? '') : '',
                row
            }))
            .filter(entry => entry.normName)
        : [];
    const outcomesNameCandidates = canSuggestNames
        ? (loadedData.outcomes || [])
            .map(row => ({
                key: row[keyConfig.outcomes],
                name: row[nameCompareConfig.outcomes],
                normName: normalizeValue(row[nameCompareConfig.outcomes]),
                city: outcomesSuggestionCityColumn ? (row[outcomesSuggestionCityColumn] ?? '') : '',
                state: outcomesSuggestionStateColumn ? (row[outcomesSuggestionStateColumn] ?? '') : '',
                country: outcomesSuggestionCountryColumn ? (row[outcomesSuggestionCountryColumn] ?? '') : '',
                row
            }))
            .filter(entry => entry.normName)
        : [];
    const suggestionIDFTable = canSuggestNames
        ? buildTokenIDFLocal([
            ...wsuNameCandidates.map(candidate => candidate.name),
            ...outcomesNameCandidates.map(candidate => candidate.name)
        ])
        : {};
    const suggestionMedianIDF = typeof suggestionIDFTable.__median === 'number'
        ? suggestionIDFTable.__median
        : 0;
    const wsuNameTokenIndex = canSuggestNames
        ? buildTokenIndex(wsuNameCandidates, suggestionIDFTable, suggestionMedianIDF)
        : {};
    const outcomesNameTokenIndex = canSuggestNames
        ? buildTokenIndex(outcomesNameCandidates, suggestionIDFTable, suggestionMedianIDF)
        : {};
    const suggestionBlockStats = {
        forwardQueries: 0,
        reverseQueries: 0,
        forwardFallbacks: 0,
        reverseFallbacks: 0
    };

    const formatSuggestionScore = (score) => (
        Number.isFinite(score) ? Math.max(0, Math.min(1, score)) : null
    );

    const getBestKeySuggestion = (value, candidates) => {
        const normalized = normalizeValue(value);
        if (!normalized) return null;
        let best = null;
        candidates.forEach(candidate => {
            const score = similarityScore(normalized, candidate.norm);
            if (!best || score > best.score) {
                best = { key: candidate.raw, score, row: candidate.row };
            }
        });
        if (!best || best.score < MIN_KEY_SUGGESTION_SCORE) return null;
        return best;
    };

    const MAX_CANDIDATES = 5;
    const dedupeCandidatesByKey = (candidates) => {
        const seen = new Set();
        const deduped = [];
        candidates.forEach(candidate => {
            const keyNorm = normalizeValue(candidate?.key || '');
            const dedupeKey = keyNorm || [
                normalizeValue(candidate?.name || ''),
                normalizeValue(candidate?.city || ''),
                normalizeValue(candidate?.state || ''),
                normalizeValue(candidate?.country || '')
            ].join('|');
            if (seen.has(dedupeKey)) return;
            seen.add(dedupeKey);
            deduped.push(candidate);
        });
        return deduped;
    };
    const getBestNameCandidates = (outcomesName, sourceState, sourceCountry, topN = MAX_CANDIDATES) => {
        if (!canSuggestNames || !outcomesName) return [];
        const candidates = [];
        suggestionBlockStats.forwardQueries += 1;
        const blockedIndices = getBlockedCandidateIndices(
            outcomesName,
            wsuNameTokenIndex,
            suggestionIDFTable
        );
        let scanCandidates = blockedIndices
            ? blockedIndices.map(idx => wsuNameCandidates[idx]).filter(Boolean)
            : wsuNameCandidates;
        if (!blockedIndices) {
            suggestionBlockStats.forwardFallbacks += 1;
        }
        const rejectByLocation = (c) => {
            if (typeof countriesMatch === 'function' && sourceCountry && c.country && !countriesMatch(sourceCountry, c.country)) return true;
            if (!sourceCountry && c.country) return true;
            if (typeof hasComparableStateValues === 'function' && typeof countriesMatch === 'function' && typeof statesMatch === 'function' &&
                hasComparableStateValues(sourceState, c.state) &&
                sourceCountry && c.country &&
                countriesMatch(sourceCountry, c.country) &&
                !statesMatch(sourceState, c.state)
            ) return true;
            return false;
        };
        scanCandidates.forEach(candidate => {
            if (rejectByLocation(candidate)) return;
            const score = typeof calculateNameSimilarity === 'function'
                ? calculateNameSimilarity(outcomesName, candidate.name, suggestionIDFTable)
                : similarityScore(normalizeValue(outcomesName), candidate.normName);
            if (score >= MIN_NAME_SUGGESTION_DISPLAY_SCORE) {
                candidates.push({
                    row: candidate.row,
                    key: candidate.key,
                    name: candidate.name,
                    city: candidate.city,
                    state: candidate.state,
                    country: candidate.country,
                    score
                });
            }
        });
        if (!candidates.length && blockedIndices) {
            suggestionBlockStats.forwardFallbacks += 1;
            scanCandidates = wsuNameCandidates;
            scanCandidates.forEach(candidate => {
                if (rejectByLocation(candidate)) return;
                const score = typeof calculateNameSimilarity === 'function'
                    ? calculateNameSimilarity(outcomesName, candidate.name, suggestionIDFTable)
                    : similarityScore(normalizeValue(outcomesName), candidate.normName);
                if (score >= MIN_NAME_SUGGESTION_DISPLAY_SCORE) {
                    candidates.push({
                        row: candidate.row,
                        key: candidate.key,
                        name: candidate.name,
                        city: candidate.city,
                        state: candidate.state,
                        country: candidate.country,
                        score
                    });
                }
            });
        }
        if (!candidates.length) return [];
        candidates.sort((a, b) => b.score - a.score);
        const uniqueCandidates = dedupeCandidatesByKey(candidates);
        if (!Number.isFinite(topN) || topN <= 0) return uniqueCandidates;
        return uniqueCandidates.slice(0, topN);
    };
    const getBestNameSuggestion = (outcomesName, sourceState, sourceCountry) => {
        const arr = getBestNameCandidates(outcomesName, sourceState, sourceCountry, 1);
        return arr.length ? arr[0] : null;
    };
    const getBestOutcomesNameCandidates = (wsuName, sourceState, sourceCountry, topN = MAX_CANDIDATES) => {
        if (!canSuggestNames || !wsuName) return [];
        const candidates = [];
        suggestionBlockStats.reverseQueries += 1;
        const rejectByLocationRev = (c) => {
            if (typeof countriesMatch === 'function' && sourceCountry && c.country && !countriesMatch(sourceCountry, c.country)) return true;
            if (!sourceCountry && c.country) return true;
            if (typeof hasComparableStateValues === 'function' && typeof countriesMatch === 'function' && typeof statesMatch === 'function' &&
                hasComparableStateValues(sourceState, c.state) &&
                sourceCountry && c.country &&
                countriesMatch(sourceCountry, c.country) &&
                !statesMatch(sourceState, c.state)
            ) return true;
            return false;
        };
        const blockedIndices = getBlockedCandidateIndices(
            wsuName,
            outcomesNameTokenIndex,
            suggestionIDFTable
        );
        let scanCandidates = blockedIndices
            ? blockedIndices.map(idx => outcomesNameCandidates[idx]).filter(Boolean)
            : outcomesNameCandidates;
        if (!blockedIndices) {
            suggestionBlockStats.reverseFallbacks += 1;
        }
        scanCandidates.forEach(candidate => {
            if (rejectByLocationRev(candidate)) return;
            const score = typeof calculateNameSimilarity === 'function'
                ? calculateNameSimilarity(wsuName, candidate.name, suggestionIDFTable)
                : similarityScore(normalizeValue(wsuName), candidate.normName);
            if (score >= MIN_NAME_SUGGESTION_DISPLAY_SCORE) {
                candidates.push({
                    row: candidate.row,
                    key: candidate.key,
                    name: candidate.name,
                    city: candidate.city,
                    state: candidate.state,
                    country: candidate.country,
                    score
                });
            }
        });
        if (!candidates.length && blockedIndices) {
            suggestionBlockStats.reverseFallbacks += 1;
            scanCandidates = outcomesNameCandidates;
            scanCandidates.forEach(candidate => {
                if (rejectByLocationRev(candidate)) return;
                const score = typeof calculateNameSimilarity === 'function'
                    ? calculateNameSimilarity(wsuName, candidate.name, suggestionIDFTable)
                    : similarityScore(normalizeValue(wsuName), candidate.normName);
                if (score >= MIN_NAME_SUGGESTION_DISPLAY_SCORE) {
                    candidates.push({
                        row: candidate.row,
                        key: candidate.key,
                        name: candidate.name,
                        city: candidate.city,
                        state: candidate.state,
                        country: candidate.country,
                        score
                    });
                }
            });
        }
        if (!candidates.length) return [];
        candidates.sort((a, b) => b.score - a.score);
        return candidates.slice(0, topN);
    };
    const getBestOutcomesNameSuggestion = (wsuName, sourceState, sourceCountry) => {
        const arr = getBestOutcomesNameCandidates(wsuName, sourceState, sourceCountry, 1);
        return arr.length ? arr[0] : null;
    };
    const applySuggestionFallbacks = (rowData, sourceRow, suggestion) => {
        const source = sourceRow || {};
        const candidate = suggestion || {};
        if (!rowData.Suggested_School) {
            rowData.Suggested_School = source.Suggested_School || candidate.name || '';
        }
        if (!rowData.Suggested_City) {
            rowData.Suggested_City = source.Suggested_City || candidate.city || '';
        }
        if (!rowData.Suggested_State) {
            rowData.Suggested_State = source.Suggested_State || candidate.state || '';
        }
        if (!rowData.Suggested_Country) {
            rowData.Suggested_Country = source.Suggested_Country || candidate.country || '';
        }
    };

    const applySuggestionColumns = (row, rowData, errorType) => {
        if (!includeSuggestions) return;
        rowData.Suggested_Key = row.Suggested_Key ?? '';
        rowData.Suggested_School = row.Suggested_School ?? '';
        rowData.Suggested_City = row.Suggested_City ?? '';
        rowData.Suggested_State = row.Suggested_State ?? '';
        rowData.Suggested_Country = row.Suggested_Country ?? '';
        rowData.Suggestion_Score = row.Suggestion_Score ?? '';
        rowData._candidates = rowData._candidates ?? [];

        const hasPresetSuggestion = Boolean(
            rowData.Suggested_Key ||
            rowData.Suggested_School ||
            rowData.Suggested_City ||
            rowData.Suggested_State ||
            rowData.Suggested_Country ||
            rowData.Suggestion_Score !== ''
        );
        const hasCompletePresetSuggestion = Boolean(
            rowData.Suggested_Key &&
            rowData.Suggested_School &&
            rowData.Suggestion_Score !== ''
        );

        if (errorType === 'Input_Not_Found') {
            const wsuState = wsuSuggestionStateColumn ? (row[`wsu_${wsuSuggestionStateColumn}`] ?? '') : '';
            const wsuCountry = wsuSuggestionCountryColumn ? (row[`wsu_${wsuSuggestionCountryColumn}`] ?? '') : '';
            let wsuName = (row[`wsu_${nameCompareConfig.wsu}`] || row.wsu_Descr || '').toString().trim();
            if (!wsuName && row && typeof row === 'object') {
                const nameHints = ['Descr', 'Org Name', 'name', 'Description', 'School', 'OrgName'];
                for (const hint of nameHints) {
                    const val = row[`wsu_${hint}`] ?? row[`wsu_${hint.replace(/\s+/g, ' ')}`];
                    if (val && String(val).trim()) {
                        wsuName = String(val).trim();
                        break;
                    }
                }
                if (!wsuName) {
                    const wsuKey = Object.keys(row || {}).find(k => k.startsWith('wsu_') && /descr|name|school|org/i.test(k));
                    if (wsuKey) wsuName = String(row[wsuKey] || '').trim();
                }
            }
            const candidates = canSuggestNames && wsuName
                ? getBestOutcomesNameCandidates(wsuName, wsuState, wsuCountry)
                : [];
            let suggestion = candidates.length ? candidates[0] : getBestKeySuggestion(row.translate_input, outcomesKeyCandidates);
            if (suggestion && normalizeValue(suggestion.key) === normalizeValue(row.translate_input)) {
                suggestion = null;
                rowData._candidates = [];
            } else if (candidates.length) {
                rowData._candidates = candidates.map((c, i) => ({ candidateId: `C${i + 1}`, ...c }));
            } else {
                rowData._candidates = [];
            }
            if (suggestion) {
                fillSuggestedFields(
                    rowData,
                    suggestion.row,
                    roleMapOutcomes,
                    selectedCols.outcomes,
                    suggestion.key,
                    ''
                );
                applySuggestionFallbacks(rowData, row, suggestion);
                rowData.Suggestion_Score = formatSuggestionScore(suggestion.score);
            }
        } else if (errorType === 'Output_Not_Found') {
            if (row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT) {
                return;
            }
            if (
                row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY &&
                hasCompletePresetSuggestion
            ) {
                return;
            }
            const outcomesStateOut = outcomesSuggestionStateColumn ? (row[`outcomes_${outcomesSuggestionStateColumn}`] ?? '') : '';
            const outcomesCountryOut = outcomesSuggestionCountryColumn ? (row[`outcomes_${outcomesSuggestionCountryColumn}`] ?? '') : '';
            const candidates = canSuggestNames
                ? getBestNameCandidates(row[`outcomes_${nameCompareConfig.outcomes}`] || row.outcomes_name || '', outcomesStateOut, outcomesCountryOut)
                : [];
            let suggestion = candidates.length ? candidates[0] : getBestKeySuggestion(row.translate_output, wsuKeyCandidates);
            if (suggestion && normalizeValue(suggestion.key) === normalizeValue(row.translate_output)) {
                suggestion = null;
                rowData._candidates = [];
            } else if (candidates.length) {
                rowData._candidates = candidates.map((c, i) => ({ candidateId: `C${i + 1}`, ...c }));
            } else {
                rowData._candidates = [];
            }
            if (suggestion) {
                fillSuggestedFields(
                    rowData,
                    suggestion.row,
                    roleMapWsu,
                    selectedCols.wsu_org,
                    suggestion.key,
                    ''
                );
                applySuggestionFallbacks(rowData, row, suggestion);
                rowData.Suggestion_Score = formatSuggestionScore(suggestion.score);
            } else if (
                row.Error_Subtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY &&
                hasPresetSuggestion
            ) {
                applySuggestionFallbacks(rowData, row, null);
            }
        } else if (
            errorType === 'Name_Mismatch' ||
            errorType === 'Ambiguous_Match' ||
            errorType === 'Duplicate_Target' ||
            errorType === 'Duplicate_Source' ||
            errorType === 'High_Confidence_Match'
        ) {
            const outcomesName = row[`outcomes_${nameCompareConfig.outcomes}`] || row.outcomes_name || '';
            const wsuName = row[`wsu_${nameCompareConfig.wsu}`] || row.wsu_Descr || '';
            if (errorType === 'Duplicate_Target' || errorType === 'Duplicate_Source') {
                // One-to-many review defaults to output-side (myWSU) suggestions.
                const outcomesState = outcomesSuggestionStateColumn ? (row[`outcomes_${outcomesSuggestionStateColumn}`] ?? '') : '';
                const outcomesCountry = outcomesSuggestionCountryColumn ? (row[`outcomes_${outcomesSuggestionCountryColumn}`] ?? '') : '';
                const candidates = canSuggestNames
                    ? getBestNameCandidates(outcomesName, outcomesState, outcomesCountry)
                    : [];
                const currentOutputNorm = normalizeValue(row.translate_output);
                const nonCurrentCandidates = candidates.filter(c => normalizeValue(c.key) !== currentOutputNorm);
                let outputSuggestion = nonCurrentCandidates.length
                    ? nonCurrentCandidates[0]
                    : getBestKeySuggestion(row.translate_output, wsuKeyCandidates);
                if (outputSuggestion && normalizeValue(outputSuggestion.key) === currentOutputNorm) {
                    outputSuggestion = null;
                }
                rowData._candidates = nonCurrentCandidates.map((c, i) => ({ candidateId: `C${i + 1}`, ...c }));
                if (outputSuggestion) {
                    fillSuggestedFields(
                        rowData,
                        outputSuggestion.row,
                        roleMapWsu,
                        selectedCols.wsu_org,
                        outputSuggestion.key,
                        ''
                    );
                    applySuggestionFallbacks(rowData, row, outputSuggestion);
                    rowData.Suggestion_Score = formatSuggestionScore(outputSuggestion.score);
                } else {
                    // Do not show current value as suggestion; leave Suggested_Key blank to avoid no-op confusion
                    fillSuggestedFields(rowData, row, roleMapWsu, selectedCols.wsu_org, '', 'wsu_');
                }
            } else if (errorType === 'High_Confidence_Match') {
                // Do not show current value as suggestion; leave Suggested_Key blank to avoid no-op confusion
                fillSuggestedFields(
                    rowData,
                    row,
                    roleMapWsu,
                    selectedCols.wsu_org,
                    '',
                    'wsu_'
                );
                const similarity = typeof calculateNameSimilarity === 'function'
                    ? calculateNameSimilarity(outcomesName, wsuName, suggestionIDFTable)
                    : null;
                if (typeof similarity === 'number') {
                    rowData.Suggestion_Score = formatSuggestionScore(similarity);
                }
            } else {
                const outcomesState = outcomesSuggestionStateColumn ? (row[`outcomes_${outcomesSuggestionStateColumn}`] ?? '') : '';
                const outcomesCountry = outcomesSuggestionCountryColumn ? (row[`outcomes_${outcomesSuggestionCountryColumn}`] ?? '') : '';
                const candidates = getBestNameCandidates(outcomesName, outcomesState, outcomesCountry);
                let suggestion = candidates.length ? candidates[0] : null;
                if (suggestion && normalizeValue(suggestion.key) === normalizeValue(row.translate_output)) {
                    suggestion = null;
                    rowData._candidates = [];
                } else if (candidates.length) {
                    rowData._candidates = candidates.map((c, i) => ({ candidateId: `C${i + 1}`, ...c }));
                } else {
                    rowData._candidates = [];
                }
                if (suggestion) {
                    fillSuggestedFields(
                        rowData,
                        suggestion.row,
                        roleMapWsu,
                        selectedCols.wsu_org,
                        suggestion.key,
                        ''
                    );
                    applySuggestionFallbacks(rowData, row, suggestion);
                    rowData.Suggestion_Score = formatSuggestionScore(suggestion.score);
                }
            }
        }
    };

    const threshold = Number.isFinite(nameCompareConfig.threshold)
        ? nameCompareConfig.threshold
        : 0.8;
    const ambiguityGap = Number.isFinite(nameCompareConfig.ambiguity_gap)
        ? nameCompareConfig.ambiguity_gap
        : 0.03;
    const resolvePreferredColumn = (preferred, roleMap, selected, priorityHints = []) => {
        if (preferred && selected.includes(preferred)) return preferred;
        if (roleMap && roleMap.School && selected.includes(roleMap.School)) return roleMap.School;
        for (const hint of priorityHints) {
            const match = selected.find(col => String(col || '').toLowerCase().includes(hint));
            if (match) return match;
        }
        return selected[0] || '';
    };
    const outcomesNameColumn = resolvePreferredColumn(
        nameCompareConfig.outcomes,
        roleMapOutcomes,
        selectedCols.outcomes || [],
        ['school', 'name']
    );
    const wsuNameColumn = resolvePreferredColumn(
        nameCompareConfig.wsu,
        roleMapWsu,
        selectedCols.wsu_org || [],
        ['descr', 'name', 'school']
    );
    const outcomesStateColumn = roleMapOutcomes.State || getFallbackRoleColumn(selectedCols.outcomes || [], 'state');
    const wsuStateColumn = roleMapWsu.State || getFallbackRoleColumn(selectedCols.wsu_org || [], 'state');
    const outcomesCityColumn = roleMapOutcomes.City || getFallbackRoleColumn(selectedCols.outcomes || [], 'city');
    const wsuCityColumn = roleMapWsu.City || getFallbackRoleColumn(selectedCols.wsu_org || [], 'city');
    const outcomesCountryColumn = roleMapOutcomes.Country || getFallbackRoleColumn(selectedCols.outcomes || [], 'country');
    const wsuCountryColumn = roleMapWsu.Country || getFallbackRoleColumn(selectedCols.wsu_org || [], 'country');
    const getCell = (row, col) => (col ? (row?.[col] ?? '') : '');

    const outcomesEntriesForMissing = (loadedData.outcomes || [])
        .map((row, idx) => ({
            idx,
            row,
            keyRaw: row[keyConfig.outcomes],
            keyNorm: normalizeKeyValue(row[keyConfig.outcomes]),
            name: getCell(row, outcomesNameColumn),
            state: getCell(row, outcomesStateColumn),
            city: getCell(row, outcomesCityColumn),
            country: getCell(row, outcomesCountryColumn)
        }))
        .filter(entry => entry.keyNorm && entry.name);

    const wsuEntriesForMissing = (loadedData.wsu_org || [])
        .map((row, idx) => ({
            idx,
            row,
            keyRaw: row[keyConfig.wsu],
            keyNorm: normalizeKeyValue(row[keyConfig.wsu]),
            name: getCell(row, wsuNameColumn),
            state: getCell(row, wsuStateColumn),
            city: getCell(row, wsuCityColumn),
            country: getCell(row, wsuCountryColumn)
        }))
        .filter(entry => entry.keyNorm && entry.name);

    const translateInputs = new Set(
        (loadedData.translate || [])
            .map(row => normalizeKeyValue(row[keyConfig.translateInput]))
            .filter(Boolean)
    );
    const translateOutputs = new Set(
        (loadedData.translate || [])
            .map(row => normalizeKeyValue(row[keyConfig.translateOutput]))
            .filter(Boolean)
    );

    const getLocationTokenSet = (outcomesEntry, wsuEntry) => {
        const tokens = new Set();
        [
            outcomesEntry?.city,
            outcomesEntry?.state,
            outcomesEntry?.country,
            wsuEntry?.city,
            wsuEntry?.state,
            wsuEntry?.country
        ].forEach(value => {
            getNameTokens(value).forEach(token => tokens.add(token));
        });
        return tokens;
    };

    const hasStrongMissingMappingNameEvidence = (outcomesEntry, wsuEntry, similarity) => {
        const outcomesTokens = new Set(getNameTokens(outcomesEntry?.name));
        const wsuTokens = new Set(getNameTokens(wsuEntry?.name));
        const overlapTokens = [];
        outcomesTokens.forEach(token => {
            if (wsuTokens.has(token)) {
                overlapTokens.push(token);
            }
        });
        const locationTokens = getLocationTokenSet(outcomesEntry, wsuEntry);
        const nonLocationOverlapCount = overlapTokens.filter(token => !locationTokens.has(token)).length;

        const countriesAlign = Boolean(
            outcomesEntry?.country &&
            wsuEntry?.country &&
            (
                (typeof countriesMatch === 'function' && countriesMatch(outcomesEntry.country, wsuEntry.country)) ||
                normalizeValue(outcomesEntry.country) === normalizeValue(wsuEntry.country)
            )
        );
        const hasComparableStates = typeof hasComparableStateValues === 'function'
            ? hasComparableStateValues(outcomesEntry?.state, wsuEntry?.state)
            : Boolean(
                String(outcomesEntry?.state || '').trim() &&
                String(wsuEntry?.state || '').trim()
            );
        const statesAlign = Boolean(
            outcomesEntry?.state &&
            wsuEntry?.state &&
            (
                (typeof statesMatch === 'function' && statesMatch(outcomesEntry.state, wsuEntry.state)) ||
                normalizeValue(outcomesEntry.state) === normalizeValue(wsuEntry.state)
            )
        );

        const singleTokenFloor = Math.max(0.72, threshold - 0.08);
        const strictFloor = (countriesAlign && hasComparableStates && statesAlign)
            ? Math.max(0.7, threshold - 0.1)
            : Math.max(0.82, threshold);

        if (nonLocationOverlapCount >= 2) return true;
        if (nonLocationOverlapCount === 1 && similarity >= singleTokenFloor) return true;
        if (nonLocationOverlapCount >= 1 && similarity >= strictFloor) return true;
        return similarity >= Math.max(0.9, threshold + 0.08);
    };

    const rowBorderByError = {
        'Input key not found in Outcomes': 'FFEF4444',
        'Output key not found in myWSU': 'FFEF4444'
    };

    const validatedList = Array.isArray(validated) ? validated : [];
    const errorRows = validatedList.filter(row => (
        row.Error_Type !== 'Valid' && row.Error_Type !== 'High_Confidence_Match'
    ));
    const translateErrorRows = errorRows.filter(row => !['Duplicate_Target', 'Duplicate_Source'].includes(row.Error_Type));
    const oneToManyRows = errorRows.filter(row => ['Duplicate_Target', 'Duplicate_Source'].includes(row.Error_Type));
    const highConfidenceRows = validatedList.filter(row => row.Error_Type === 'High_Confidence_Match');
    const validRows = validatedList.filter(row => row.Error_Type === 'Valid');

    const errorColumns = [
        'Error_Type',
        'Error_Subtype',
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        ...mappingColumns,
        ...suggestionColumns
    ];
    const oneToManyColumns = [
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        ...mappingColumns,
        ...suggestionColumns
    ];
    const validColumns = [
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        ...mappingColumns
    ];
    const highConfidenceColumns = [
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        ...mappingColumns,
        ...suggestionColumns
    ];

    const errorDataRows = translateErrorRows.map(row => {
        const rowData = {};
        errorColumns.forEach(col => {
            rowData[col] = row[col] !== undefined ? row[col] : '';
        });
        rowData.Error_Type = normalizeErrorType(row) || row.Error_Type;
        rowData.Error_Subtype = normalizeErrorSubtype(row.Error_Subtype);
        rowData._rawErrorType = row.Error_Type;
        rowData._rawErrorSubtype = row.Error_Subtype;
        if (showMappingLogic) {
            rowData.Mapping_Logic = buildMappingLogicRow(row, rowData.Error_Type, nameCompareConfig, suggestionIDFTable);
        }
        applySuggestionColumns(row, rowData, row.Error_Type);
        return rowData;
    });

    const oneToManyDataRows = oneToManyRows.map(row => {
        const rowData = {};
        oneToManyColumns.forEach(col => {
            rowData[col] = row[col] !== undefined ? row[col] : '';
        });
        rowData.Error_Type = row.Error_Type;
        rowData.Duplicate_Group = row.Duplicate_Group || '';
        if (showMappingLogic) {
            rowData.Mapping_Logic = buildMappingLogicRow(row, row.Error_Type, nameCompareConfig, suggestionIDFTable);
        }
        applySuggestionColumns(row, rowData, row.Error_Type);
        return rowData;
    });

    oneToManyDataRows.sort((a, b) => {
        const groupA = a.Duplicate_Group || '';
        const groupB = b.Duplicate_Group || '';
        if (groupA !== groupB) return groupA.localeCompare(groupB);
        const typeA = a.Error_Type || '';
        const typeB = b.Error_Type || '';
        return typeA.localeCompare(typeB);
    });

    const validDataRows = validRows.map(row => {
        const rowData = {};
        validColumns.forEach(col => {
            rowData[col] = row[col] !== undefined ? row[col] : '';
        });
        if (showMappingLogic) {
            rowData.Mapping_Logic = buildMappingLogicRow(row, 'Valid', nameCompareConfig, suggestionIDFTable);
        }
        return rowData;
    });

    const highConfidenceDataRows = highConfidenceRows.map(row => {
        const rowData = {};
        highConfidenceColumns.forEach(col => {
            rowData[col] = row[col] !== undefined ? row[col] : '';
        });
        if (showMappingLogic) {
            rowData.Mapping_Logic = buildMappingLogicRow(row, 'High_Confidence_Match', nameCompareConfig, suggestionIDFTable);
        }
        applySuggestionColumns(row, rowData, row.Error_Type);
        return rowData;
    });

    const baseStyle = {
        errorHeaderColor: 'FF991B1B',
        errorBodyColor: 'FFFEE2E2',
        groupHeaderColor: 'FFF59E0B',
        groupBodyColor: 'FFFEF3C7',
        logicHeaderColor: 'FF374151',
        logicBodyColor: 'FFF3F4F6',
        translateHeaderColor: 'FF1E40AF',
        translateBodyColor: 'FFDBEAFE',
        outcomesHeaderColor: 'FF166534',
        outcomesBodyColor: 'FFDCFCE7',
        wsuHeaderColor: 'FFC2410C',
        wsuBodyColor: 'FFFFEDD5',
        suggestionHeaderColor: 'FF6D28D9',
        suggestionBodyColor: 'FFEDE9FE',
        defaultHeaderColor: 'FF981e32',
        defaultBodyColor: 'FFFFFFFF'
    };

    const includeNonMissingDiagnosticSheets = !missingOnlyExport;
    if (includeNonMissingDiagnosticSheets) {
        reportProgress('Building Errors_in_Translate...', 20);
        addSheetWithRows(workbook, {
            sheetName: 'Errors_in_Translate',
            outputColumns: errorColumns,
            rows: errorDataRows,
            style: baseStyle,
            headers: buildHeaders(errorColumns, keyLabels),
            rowBorderByError
        });

        const filterBySubtype = (typeof ValidationExportHelpers !== 'undefined' && ValidationExportHelpers &&
            typeof ValidationExportHelpers.filterOutputNotFoundBySubtype === 'function')
            ? ValidationExportHelpers.filterOutputNotFoundBySubtype
            : (rows, subtype) => rows.filter(row =>
                row._rawErrorType === 'Output_Not_Found' && row._rawErrorSubtype === subtype
            );
        const outputNotFoundAmbiguousFiltered = filterBySubtype(errorDataRows, EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT);
        const outputNotFoundNoReplacementFiltered = filterBySubtype(errorDataRows, EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT);

        reportProgress('Building Output_Not_Found_Ambiguous...', 28);
        if (outputNotFoundAmbiguousFiltered.length > 0) {
            addSheetWithRows(workbook, {
                sheetName: 'Output_Not_Found_Ambiguous',
                outputColumns: errorColumns,
                rows: outputNotFoundAmbiguousFiltered,
                style: baseStyle,
                headers: buildHeaders(errorColumns, keyLabels),
                rowBorderByError
            });
        }
        reportProgress('Building Output_Not_Found_No_Replacement...', 32);
        if (outputNotFoundNoReplacementFiltered.length > 0) {
            addSheetWithRows(workbook, {
                sheetName: 'Output_Not_Found_No_Replacement',
                outputColumns: errorColumns,
                rows: outputNotFoundNoReplacementFiltered,
                style: baseStyle,
                headers: buildHeaders(errorColumns, keyLabels),
                rowBorderByError
            });
        }

        reportProgress('Building One_to_Many...', 35);
        addSheetWithRows(workbook, {
            sheetName: 'One_to_Many',
            outputColumns: oneToManyColumns,
            rows: oneToManyDataRows,
            style: baseStyle,
            headers: buildHeaders(oneToManyColumns, keyLabels),
            rowBorderByError,
            groupColumn: 'Duplicate_Group'
        });

        reportProgress('Building High_Confidence_Matches...', 50);
        addSheetWithRows(workbook, {
            sheetName: 'High_Confidence_Matches',
            outputColumns: highConfidenceColumns,
            rows: highConfidenceDataRows,
            style: baseStyle,
            headers: buildHeaders(highConfidenceColumns, keyLabels),
            rowBorderByError
        });

        reportProgress('Building Valid_Mappings...', 62);
        addSheetWithRows(workbook, {
            sheetName: 'Valid_Mappings',
            outputColumns: validColumns,
            rows: validDataRows,
            style: {
                ...baseStyle,
                errorHeaderColor: 'FF16A34A',
                errorBodyColor: 'FFDCFCE7'
            },
            headers: buildHeaders(validColumns, keyLabels),
            rowBorderByError
        });
    } else {
        reportProgress('Skipping non-missing diagnostic sheets for missing-only scope...', 62);
    }

    reportProgress('Building Missing_Mappings...', 74);
    const missingMappingColumns = [
        'Missing_In',
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        'Similarity',
        ...mappingColumns
    ];
    let missingMappingsRows = [];
    if (!translationOnlyExport) {
        const highConfidenceCandidates = [];
        outcomesEntriesForMissing.forEach(outcomesEntry => {
            let best = null;
            let secondBest = null;
            wsuEntriesForMissing.forEach(wsuEntry => {
                if (
                    outcomesEntry.country && wsuEntry.country &&
                    !countriesMatch(outcomesEntry.country, wsuEntry.country)
                ) {
                    return;
                }
                const stateComparable = Boolean(
                    outcomesEntry.state &&
                    wsuEntry.state &&
                    String(outcomesEntry.state).trim().toLowerCase() !== 'ot' &&
                    String(wsuEntry.state).trim().toLowerCase() !== 'ot'
                );
                if (
                    stateComparable &&
                    outcomesEntry.country &&
                    wsuEntry.country &&
                    countriesMatch(outcomesEntry.country, wsuEntry.country) &&
                    !statesMatch(outcomesEntry.state, wsuEntry.state)
                ) {
                    return;
                }

                const similarity = calculateNameSimilarity(outcomesEntry.name, wsuEntry.name, suggestionIDFTable);
                const highConfidence = isHighConfidenceNameMatch(
                    outcomesEntry.name,
                    wsuEntry.name,
                    outcomesEntry.state,
                    wsuEntry.state,
                    outcomesEntry.city,
                    wsuEntry.city,
                    outcomesEntry.country,
                    wsuEntry.country,
                    similarity,
                    threshold,
                    suggestionIDFTable
                );
                if (!highConfidence) {
                    return;
                }
                if (!hasStrongMissingMappingNameEvidence(outcomesEntry, wsuEntry, similarity)) {
                    return;
                }
                const candidate = {
                    outcomesEntry,
                    wsuEntry,
                    score: similarity
                };
                if (!best || candidate.score > best.score) {
                    secondBest = best;
                    best = candidate;
                } else if (!secondBest || candidate.score > secondBest.score) {
                    secondBest = candidate;
                }
            });
            if (!best) {
                return;
            }
            if (secondBest && (best.score - secondBest.score) < ambiguityGap) {
                return;
            }
            highConfidenceCandidates.push(best);
        });

        highConfidenceCandidates.sort((a, b) => b.score - a.score);
        const usedOutcomes = new Set();
        const usedWsu = new Set();
        const highConfidencePairs = [];
        highConfidenceCandidates.forEach(candidate => {
            const outcomesId = candidate.outcomesEntry.idx;
            const wsuId = candidate.wsuEntry.idx;
            if (usedOutcomes.has(outcomesId) || usedWsu.has(wsuId)) {
                return;
            }
            usedOutcomes.add(outcomesId);
            usedWsu.add(wsuId);
            highConfidencePairs.push(candidate);
        });

        missingMappingsRows = highConfidencePairs
            .map(pair => {
                const outcomesEntry = pair.outcomesEntry;
                const wsuEntry = pair.wsuEntry;
                const inputPresent = translateInputs.has(outcomesEntry.keyNorm);
                const outputPresent = translateOutputs.has(wsuEntry.keyNorm);
                if (inputPresent && outputPresent) {
                    return null;
                }

                const rowData = {};
                rowData.Missing_In = (!inputPresent && !outputPresent)
                    ? 'Input and Output missing in Translate'
                    : (!inputPresent ? 'Input missing in Translate' : 'Output missing in Translate');
                selectedCols.outcomes.forEach(col => {
                    rowData[`outcomes_${col}`] = outcomesEntry.row[col] ?? '';
                });
                rowData.translate_input = outcomesEntry.keyRaw ?? outcomesEntry.keyNorm;
                rowData.translate_output = wsuEntry.keyRaw ?? wsuEntry.keyNorm;
                selectedCols.wsu_org.forEach(col => {
                    rowData[`wsu_${col}`] = wsuEntry.row[col] ?? '';
                });
                rowData.Similarity = formatSuggestionScore(pair.score);
                if (showMappingLogic) {
                    rowData.Mapping_Logic = `High-confidence Outcomes<->myWSU name/location match (${formatScore(pair.score)}); ${rowData.Missing_In.toLowerCase()}.`;
                }
                return rowData;
            })
            .filter(Boolean)
            .sort((a, b) => {
                const missingOrder = { 'Input and Output missing in Translate': 0, 'Input missing in Translate': 1, 'Output missing in Translate': 2 };
                const orderA = missingOrder[a.Missing_In] ?? 99;
                const orderB = missingOrder[b.Missing_In] ?? 99;
                if (orderA !== orderB) return orderA - orderB;
                return String(a.translate_input || '').localeCompare(String(b.translate_input || ''));
            });

        addSheetWithRows(workbook, {
            sheetName: 'Missing_Mappings',
            outputColumns: missingMappingColumns,
            rows: missingMappingsRows,
            style: baseStyle,
            headers: buildHeaders(missingMappingColumns, keyLabels),
            rowBorderByError
        });
    }

    reportProgress('Building Action_Queue...', 78);
    const defaultDecisionThreshold = Number.isFinite(nameCompareConfig?.threshold)
        ? nameCompareConfig.threshold
        : 0.85;
    const getDefaultDecision = (row) => {
        const rawType = row._rawErrorType || row.Error_Type || '';
        const rawSub = String(row._rawErrorSubtype || row.Error_Subtype || '');
        const suggestedKey = String(row.Suggested_Key || '').trim();
        const scoreVal = row.Suggestion_Score;
        const score = Number.isFinite(scoreVal) ? scoreVal : (typeof scoreVal === 'string' ? parseFloat(scoreVal) : NaN);
        const hasSuggestion = suggestedKey !== '' || (Number.isFinite(score) && score >= defaultDecisionThreshold);
        const hasActionableSuggestion = suggestedKey !== '';
        if (rawType === 'Name_Mismatch' && Number.isFinite(score) && score >= defaultDecisionThreshold) return 'Keep As-Is';
        if (rawType === 'Output_Not_Found') {
            if (rawSub === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT) return 'Ignore';
            if (rawSub === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY && hasActionableSuggestion) return 'Use Suggestion';
        }
        if (rawType === 'Missing_Mapping') return 'Keep As-Is';
        if ((rawType === 'Duplicate_Source' || rawType === 'Duplicate_Target') && hasActionableSuggestion) return 'Use Suggestion';
        if (rawType === 'Input_Not_Found' && hasActionableSuggestion) return 'Use Suggestion';
        if (rawType === 'Duplicate_Target') return 'Keep As-Is';
        return '';
    };
    const actionQueueFromErrors = missingOnlyExport ? [] : errorDataRows.map(row => {
        const rawType = row._rawErrorType || row.Error_Type;
        const rawSub = row._rawErrorSubtype || row.Error_Subtype;
        const priority = getActionPriority(rawType, rawSub);
        const action = getRecommendedAction(rawType, rawSub);
        return {
            ...row,
            Is_Stale_Key: row._rawErrorSubtype === EXPORT_OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY ? 1 : 0,
            Priority: priority,
            Recommended_Action: action,
            Source_Sheet: 'Errors_in_Translate',
            Decision: getDefaultDecision(row) || '',
            Owner: '',
            Status: '',
            Resolution_Note: '',
            Resolved_Date: '',
            Reviewer: '',
            Review_Date: '',
            Reason_Code: '',
            Notes: ''
        };
    });
    const actionQueueFromOneToMany = missingOnlyExport ? [] : oneToManyDataRows.map(row => {
        const rawType = row.Error_Type || '';
        const priority = getActionPriority(rawType, '');
        const action = getRecommendedAction(rawType, '');
        return {
            ...row,
            _rawErrorType: rawType,
            _rawErrorSubtype: '',
            Is_Stale_Key: 0,
            Priority: priority,
            Recommended_Action: action,
            Source_Sheet: 'One_to_Many',
            Error_Type: row.Error_Type,
            Error_Subtype: '',
            Missing_In: '',
            Similarity: '',
            Decision: getDefaultDecision({ ...row, _rawErrorType: row.Error_Type }) || 'Allow One-to-Many',
            Owner: '',
            Status: '',
            Resolution_Note: '',
            Resolved_Date: '',
            Reviewer: '',
            Review_Date: '',
            Reason_Code: '',
            Notes: ''
        };
    });
    const actionQueueFromMissing = translationOnlyExport ? [] : missingMappingsRows.map(row => {
        const priority = getActionPriority('Missing_Mapping', row.Missing_In);
        const action = getRecommendedAction('Missing_Mapping', row.Missing_In);
        return {
            ...row,
            _rawErrorType: 'Missing_Mapping',
            _rawErrorSubtype: row.Missing_In || '',
            Is_Stale_Key: 0,
            Error_Type: 'Missing_Mapping',
            Error_Subtype: row.Missing_In || '',
            Priority: priority,
            Recommended_Action: action,
            Source_Sheet: 'Missing_Mappings',
            Missing_In: row.Missing_In || '',
            Similarity: row.Similarity ?? '',
            Decision: getDefaultDecision({ ...row, _rawErrorType: 'Missing_Mapping' }) || 'Keep As-Is',
            Owner: '',
            Status: '',
            Resolution_Note: '',
            Resolved_Date: '',
            Reviewer: '',
            Review_Date: '',
            Reason_Code: '',
            Notes: ''
        };
    });
    const actionQueueFromErrorsWithMissing = actionQueueFromErrors.map(row => ({
        ...row,
        Missing_In: row.Missing_In ?? '',
        Similarity: row.Similarity ?? row.Suggestion_Score ?? ''
    }));
    const normKey = (k) => String(k || '').trim();
    const isCandidateErrorRow = (row) => {
        const raw = row._rawErrorType || row.Error_Type || '';
        return raw === 'Input_Not_Found' || raw === 'Output_Not_Found';
    };
    const getErrorCanonicalPair = (row) => {
        const raw = row._rawErrorType || row.Error_Type || '';
        const sug = normKey(row.Suggested_Key);
        if (!sug) return null;
        if (raw === 'Input_Not_Found') return [sug, normKey(row.translate_output)];
        if (raw === 'Output_Not_Found') return [normKey(row.translate_input), sug];
        return null;
    };
    const getMissingCanonicalPair = (row) => {
        const inp = normKey(row.translate_input);
        const out = normKey(row.translate_output);
        if (!inp || !out) return null;
        return [inp, out];
    };
    const parseScore = (v) => {
        if (Number.isFinite(v)) return v;
        if (typeof v === 'string') return parseFloat(v) || 0;
        return 0;
    };
    const missingPairs = new Map();
    actionQueueFromMissing.forEach((row, idx) => {
        const pair = getMissingCanonicalPair(row);
        if (!pair) return;
        const key = `${pair[0]}\t${pair[1]}`;
        const existing = missingPairs.get(key);
        const score = parseScore(row.Similarity);
        if (!existing || score > parseScore(existing.row.Similarity)) {
            missingPairs.set(key, { row, idx, pair });
        } else if (score === parseScore(existing.row.Similarity)) {
            const tiA = String(row.translate_input || '');
            const tiB = String(existing.row.translate_input || '');
            if (tiA !== tiB ? tiA < tiB : String(row.translate_output || '') < String(existing.row.translate_output || '')) {
                missingPairs.set(key, { row, idx, pair });
            }
        }
    });
    const candidateErrors = actionQueueFromErrorsWithMissing.filter(isCandidateErrorRow);
    const hasCandidates = candidateErrors.length > 0;
    const hasMissing = actionQueueFromMissing.length > 0;
    const errorRowsToDrop = new Set();
    const missingRowsToDrop = new Set();
    if (hasCandidates && hasMissing) {
        candidateErrors.forEach((errRow) => {
            const pair = getErrorCanonicalPair(errRow);
            if (!pair) return;
            const [inp, out] = pair;
            const key = `${inp}\t${out}`;
            const missingEntry = missingPairs.get(key);
            if (!missingEntry) return;
            const missingRow = missingEntry.row;
            const missingScore = parseScore(missingRow.Similarity);
            const errScore = parseScore(errRow.Suggestion_Score);
            const missingHasConfidence = missingScore > 0 || String(missingRow.Similarity || '').trim() !== '';
            const errHasSuggestion = normKey(errRow.Suggested_Key) !== '';
            if (!errHasSuggestion) return;
            if (missingHasConfidence && missingScore >= errScore) {
                errorRowsToDrop.add(errRow);
                if (!missingRow.Merged_Sources) missingRow.Merged_Sources = [];
                missingRow.Merged_Sources.push(errRow._rawErrorType || errRow.Error_Type || '');
            } else if (errScore > missingScore) {
                missingRowsToDrop.add(missingRow);
            }
        });
    }
    const filteredErrors = actionQueueFromErrorsWithMissing.filter((row) => {
        if (!isCandidateErrorRow(row)) return true;
        return !errorRowsToDrop.has(row);
    });
    const filteredMissing = actionQueueFromMissing.filter((row) => !missingRowsToDrop.has(row));
    const getDuplicateCanonicalPair = (row) => {
        const raw = row._rawErrorType || row.Error_Type || '';
        const sug = normKey(row.Suggested_Key);
        if (!sug) return null;
        if (raw === 'Duplicate_Target') return [normKey(row.translate_input), sug];
        if (raw === 'Duplicate_Source') return [sug, normKey(row.translate_output)];
        return null;
    };
    const errorPairKeys = new Set();
    filteredErrors.forEach((row) => {
        if (row.Decision !== 'Use Suggestion') return;
        const pair = getErrorCanonicalPair(row);
        if (pair) errorPairKeys.add(`${pair[0]}\t${pair[1]}`);
    });
    const duplicateRowsToDrop = new Set();
    actionQueueFromOneToMany.forEach((row) => {
        if (row.Decision !== 'Use Suggestion') return;
        const pair = getDuplicateCanonicalPair(row);
        if (!pair) return;
        const key = `${pair[0]}\t${pair[1]}`;
        if (errorPairKeys.has(key)) duplicateRowsToDrop.add(row);
    });
    const filteredOneToMany = actionQueueFromOneToMany.filter((row) => !duplicateRowsToDrop.has(row));
    const actionQueueRowsUnstable = [...filteredErrors, ...filteredOneToMany, ...filteredMissing]
        .sort((a, b) => {
            const pa = a.Priority ?? 99;
            const pb = b.Priority ?? 99;
            if (pa !== pb) return pa - pb;
            const sa = a.Source_Sheet || '';
            const sb = b.Source_Sheet || '';
            if (sa !== sb) return sa.localeCompare(sb);
            const ea = a.Error_Type || '';
            const eb = b.Error_Type || '';
            if (ea !== eb) return ea.localeCompare(eb);
            const tiA = String(a.translate_input || '');
            const tiB = String(b.translate_input || '');
            if (tiA !== tiB) return tiA.localeCompare(tiB);
            const toA = String(a.translate_output || '');
            const toB = String(b.translate_output || '');
            return toA.localeCompare(toB);
        });
    const inferKeyUpdateSide = (row) => {
        const rawType = row._rawErrorType || '';
        if (rawType === 'Input_Not_Found' || rawType === 'Missing_Input') {
            return 'Input';
        }
        if (rawType === 'Output_Not_Found' || rawType === 'Missing_Output') {
            return 'Output';
        }
        if (rawType === 'Missing_Mapping') {
            const missing = String(row.Missing_In || '');
            if (missing.includes('Input and Output')) return 'Both';
            if (missing.includes('Input')) return 'Input';
            if (missing.includes('Output')) return 'Output';
        }
        if (rawType === 'Duplicate_Source') return 'Input';
        if (rawType === 'Duplicate_Target') return 'Output';
        return 'None';
    };
    const sanitizeIdPart = (value) => String(normalizeValue(value || '')).replace(/\|/g, '/');
    const buildStableReviewId = (row) => {
        const outcomesKeyValue = keyLabels.outcomes ? row[`outcomes_${keyLabels.outcomes}`] : '';
        const wsuKeyValue = keyLabels.wsu ? row[`wsu_${keyLabels.wsu}`] : '';
        return [
            sanitizeIdPart(row.Source_Sheet || ''),
            sanitizeIdPart(row._rawErrorType || row.Error_Type || ''),
            sanitizeIdPart(row.translate_input || ''),
            sanitizeIdPart(row.translate_output || ''),
            sanitizeIdPart(outcomesKeyValue || ''),
            sanitizeIdPart(wsuKeyValue || ''),
            sanitizeIdPart(row.Missing_In || ''),
            sanitizeIdPart(row.Duplicate_Group || '')
        ].join('|');
    };
    const fmtCandidateChoice = (c) => {
        if (!c || (!c.key && !c.name)) return '';
        const loc = [c.city, c.state, c.country].filter(Boolean).join(', ');
        const locPart = loc ? ` - ${loc}` : '';
        const scoreVal = typeof c.score === 'number' ? c.score : (parseFloat(c.score) || NaN);
        const scorePart = Number.isFinite(scoreVal) ? ` | Score: ${scoreVal.toFixed(2)}` : '';
        return `${c.key || ''}: ${c.name || ''}${locPart}${scorePart}`;
    };
    const actionQueueRows = actionQueueRowsUnstable.map(row => {
        const keyUpdateSide = inferKeyUpdateSide(row);
        const currentValue = keyUpdateSide === 'Input' ? (row.translate_input ?? '') : (row.translate_output ?? '');
        const candidates = row._candidates || [];
        const c1 = candidates[0];
        const c1IsCurrentValue = c1 && normalizeValue(c1.key) === normalizeValue(currentValue);
        const defaultUseSuggestion = (row.Decision === 'Use Suggestion') && candidates.length && !c1IsCurrentValue;
        const selectedCandidateId = defaultUseSuggestion ? 'C1' : '';
        return {
            ...row,
            Suggested_Key: row.Suggested_Key ?? '',
            Suggested_School: row.Suggested_School ?? '',
            Suggested_City: row.Suggested_City ?? '',
            Suggested_State: row.Suggested_State ?? '',
            Suggested_Country: row.Suggested_Country ?? '',
            Suggestion_Score: row.Suggestion_Score ?? '',
            Selected_Candidate_ID: selectedCandidateId,
            Manual_Suggested_Key: '',
            C1_Choice: fmtCandidateChoice(candidates[0]),
            C2_Choice: fmtCandidateChoice(candidates[1]),
            C3_Choice: fmtCandidateChoice(candidates[2]),
            Current_Input: row.translate_input ?? '',
            Current_Output: row.translate_output ?? '',
            Key_Update_Side: keyUpdateSide,
            Review_Row_ID: buildStableReviewId(row),
            Final_Input: '',
            Final_Output: '',
            Publish_Eligible: '',
            Approval_Source: '',
            Has_Update: ''
        };
    });
    const reviewIdCount = new Map();
    actionQueueRows.forEach(row => {
        const baseId = row.Review_Row_ID;
        const seen = (reviewIdCount.get(baseId) || 0) + 1;
        reviewIdCount.set(baseId, seen);
        if (seen > 1) {
            row.Review_Row_ID = `${baseId}#${seen}`;
        }
    });
    const matchCampusFamilyPattern = (pattern, value) => {
        if (!value || !pattern) return false;
        const str = String(value).trim();
        if (pattern.endsWith('*')) {
            const prefix = pattern.slice(0, -1).trim();
            return str.toUpperCase().startsWith(prefix.toUpperCase());
        }
        return str.toUpperCase() === String(pattern).trim().toUpperCase();
    };
    const getRowNameForCampusFamily = (row) => {
        const outcomesNameCol = nameCompareConfig.outcomes;
        if (outcomesNameCol && row[`outcomes_${outcomesNameCol}`]) {
            return String(row[`outcomes_${outcomesNameCol}`] || '').trim();
        }
        const nameHints = ['Descr', 'Org Name', 'name', 'School', 'Description', 'OrgName'];
        for (const hint of nameHints) {
            const val = row[`outcomes_${hint}`] ?? row[`wsu_${hint}`];
            if (val && String(val).trim()) return String(val).trim();
        }
        const outcomesKey = Object.keys(row || {}).find(k => k.startsWith('outcomes_') && /descr|name|school|org/i.test(k));
        if (outcomesKey) return String(row[outcomesKey] || '').trim();
        return '';
    };
    if (campusFamilyRules && Array.isArray(campusFamilyRules.patterns) && campusFamilyRules.patterns.length > 0) {
        const validOutputKeysSet = new Set((wsuKeyCandidates || []).map(c => normalizeValue(c.raw || '')));
        const rules = campusFamilyRules.patterns
            .filter(r => r && r.enabled !== false && r.pattern && r.parentKey)
            .sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
        actionQueueRows.forEach(row => {
            if (row.Manual_Suggested_Key && String(row.Manual_Suggested_Key).trim()) return;
            const keyUpdateSide = row.Key_Update_Side || '';
            if (keyUpdateSide !== 'Output' && keyUpdateSide !== 'Both') return;
            const rowName = getRowNameForCampusFamily(row);
            if (!rowName) return;
            for (const rule of rules) {
                if (!matchCampusFamilyPattern(rule.pattern, rowName)) continue;
                const parentNorm = normalizeValue(rule.parentKey);
                if (!validOutputKeysSet.has(parentNorm)) continue;
                if (rule.country && outcomesSuggestionCountryColumn && row[`outcomes_${outcomesSuggestionCountryColumn}`] && typeof countriesMatch === 'function' && !countriesMatch(row[`outcomes_${outcomesSuggestionCountryColumn}`], rule.country)) continue;
                if (rule.state && outcomesSuggestionStateColumn && String(row[`outcomes_${outcomesSuggestionStateColumn}`] || '').trim() !== String(rule.state || '').trim()) continue;
                row.Manual_Suggested_Key = rule.parentKey;
                break;
            }
        });
    }
    const reimportSummary = { applied: 0, conflicts: 0, newRows: 0, orphaned: 0 };
    const priorDecisions = priorDecisionsPayload && typeof priorDecisionsPayload === 'object' && Object.keys(priorDecisionsPayload).length > 0
        ? priorDecisionsPayload
        : null;
    if (priorDecisions) {
        const validOutputKeysSet = new Set((wsuKeyCandidates || []).map(c => normalizeValue(c.raw || '')));
        const validInputKeysSet = new Set((outcomesKeyCandidates || []).map(c => normalizeValue(c.raw || '')));
        const priorKeysVisited = new Set();
        actionQueueRows.forEach(row => {
            const reviewRowId = row.Review_Row_ID || '';
            const prior = priorDecisions[reviewRowId];
            if (!prior) {
                reimportSummary.newRows += 1;
                return;
            }
            priorKeysVisited.add(reviewRowId);
            const priorDecision = String(prior.Decision || '').trim();
            const priorManual = String(prior.Manual_Suggested_Key || '').trim();
            const priorSuggested = String(prior.Suggested_Key || '').trim();
            const effectiveKey = priorManual || priorSuggested;
            if (priorDecision === 'Use Suggestion') {
                if (!effectiveKey) {
                    reimportSummary.conflicts += 1;
                    return;
                }
                const keyUpdateSide = row.Key_Update_Side || '';
                const validKeys = (keyUpdateSide === 'Output' || keyUpdateSide === 'Both') ? validOutputKeysSet : validInputKeysSet;
                const keyNorm = normalizeValue(effectiveKey);
                if (!validKeys.has(keyNorm)) {
                    reimportSummary.conflicts += 1;
                    return;
                }
                row.Manual_Suggested_Key = priorManual || priorSuggested;
            }
            row.Decision = priorDecision || row.Decision;
            if (prior.Reason_Code !== undefined && prior.Reason_Code !== null) row.Reason_Code = prior.Reason_Code;
            reimportSummary.applied += 1;
        });
        reimportSummary.orphaned = Object.keys(priorDecisions).filter(k => !priorKeysVisited.has(k)).length;
    }
    if (preEditedActionQueueRows && Array.isArray(preEditedActionQueueRows) && preEditedActionQueueRows.length > 0) {
        const preEditedMap = new Map();
        preEditedActionQueueRows.forEach(r => {
            const id = r.Review_Row_ID || '';
            if (id) preEditedMap.set(id, r);
        });
        actionQueueRows.forEach(row => {
            const edited = preEditedMap.get(row.Review_Row_ID || '');
            if (!edited) return;
            if (edited.Decision !== undefined && edited.Decision !== null) row.Decision = edited.Decision;
            if (edited.Selected_Candidate_ID !== undefined && edited.Selected_Candidate_ID !== null) row.Selected_Candidate_ID = edited.Selected_Candidate_ID;
            if (edited.Manual_Suggested_Key !== undefined && edited.Manual_Suggested_Key !== null) row.Manual_Suggested_Key = edited.Manual_Suggested_Key;
            if (edited.Reason_Code !== undefined && edited.Reason_Code !== null) row.Reason_Code = edited.Reason_Code;
        });
    }
    if (returnActionQueueOnly) {
        return { actionQueueRows };
    }
    const candidatePoolRows = [];
    actionQueueRows.forEach(row => {
        const candidates = row._candidates || [];
        const reviewRowId = row.Review_Row_ID || '';
        candidates.forEach(c => {
            candidatePoolRows.push({
                Composite_Key: `${reviewRowId}|${c.candidateId || ''}`,
                Key: c.key ?? '',
                Name: c.name ?? '',
                City: c.city ?? '',
                State: c.state ?? '',
                Country: c.country ?? '',
                Score: typeof c.score === 'number' ? c.score : (parseFloat(c.score) || 0)
            });
        });
    });
    const candidatePoolColumns = ['Composite_Key', 'Key', 'Name', 'City', 'State', 'Country', 'Score'];
    const candidatePoolLastRow = Math.max(2, candidatePoolRows.length + 1);
    const actionQueueBaseCols = [
        'Review_Row_ID',
        'Priority',
        'Recommended_Action',
        'Error_Type',
        'Error_Subtype',
        'Source_Sheet',
        'Key_Update_Side',
        'Is_Stale_Key',
        'Missing_In',
        'Similarity',
        ...errorColumns.filter(c => !['Error_Type', 'Error_Subtype', '_rawErrorType', '_rawErrorSubtype'].includes(c)),
        ...reviewSuggestionColumns,
        'Selected_Candidate_ID',
        'Decision',
        'Owner',
        'Status',
        'Resolution_Note',
        'Resolved_Date',
        'Reviewer',
        'Review_Date',
        'Reason_Code',
        'Notes'
    ];
    const actionQueueColumns = actionQueueBaseCols.filter((v, i, arr) => arr.indexOf(v) === i);
    const actionQueueHeaders = buildHeaders(actionQueueColumns, keyLabels).map((h, i) => {
        const col = actionQueueColumns[i];
        if (col === 'Review_Row_ID') return 'Review Row ID';
        if (col === 'Recommended_Action') return 'Recommended Action';
        if (col === 'Source_Sheet') return 'Source Sheet';
        if (col === 'Key_Update_Side') return 'Update Side';
        if (col === 'Is_Stale_Key') return 'Stale Key (1=yes)';
        if (col === 'Selected_Candidate_ID') return 'Selected Candidate ID';
        if (col === 'Resolution_Note') return 'Resolution Note';
        if (col === 'Resolved_Date') return 'Resolved Date';
        if (col === 'Review_Date') return 'Review Date';
        if (col === 'Reason_Code') return 'Reason Code';
        return h || col;
    });
    addSheetWithRows(workbook, {
        sheetName: 'Action_Queue',
        outputColumns: actionQueueColumns,
        rows: actionQueueRows,
        style: baseStyle,
        headers: actionQueueHeaders,
        rowBorderByError: null
    });
    const colDecisionAq = columnIndexToLetter(actionQueueColumns.indexOf('Decision') + 1);
    const aqSheet = workbook.getWorksheet('Action_Queue');
    if (aqSheet) {
        // Internal staging sheet; keep workbook navigation focused on reviewer tabs.
        aqSheet.state = 'hidden';
    }
    if (aqSheet && actionQueueRows.length > 0) {
        aqSheet.dataValidations.add(
            `${colDecisionAq}2:${colDecisionAq}${actionQueueRows.length + 1}`,
            {
                type: 'list',
                allowBlank: true,
                formulae: ['"Keep As-Is,Use Suggestion,Allow One-to-Many,Ignore"']
            }
        );
    }

    reportProgress('Building Candidate_Pool...', 81);
    addSheetWithRows(workbook, {
        sheetName: 'Candidate_Pool',
        outputColumns: candidatePoolColumns,
        rows: candidatePoolRows,
        style: baseStyle,
        headers: ['Composite Key', 'Key', 'Name', 'City', 'State', 'Country', 'Score']
    });
    const cpSheet = workbook.getWorksheet('Candidate_Pool');
    if (cpSheet) cpSheet.state = 'hidden';
    const candidateReferenceRows = [];
    actionQueueRows.forEach(row => {
        const candidates = row._candidates || [];
        if (candidates.length === 0) return;
        candidates.forEach(c => {
            candidateReferenceRows.push({
                Review_Row_ID: row.Review_Row_ID,
                Candidate_ID: c.candidateId || '',
                Name: c.name ?? '',
                City: c.city ?? '',
                State: c.state ?? '',
                Country: c.country ?? '',
                Score: typeof c.score === 'number' ? c.score : (parseFloat(c.score) || 0)
            });
        });
    });
    addSheetWithRows(workbook, {
        sheetName: 'Candidate_Reference',
        outputColumns: ['Review_Row_ID', 'Candidate_ID', 'Name', 'City', 'State', 'Country', 'Score'],
        rows: candidateReferenceRows,
        style: baseStyle,
        headers: ['Review Row ID', 'Candidate ID', 'Name', 'City', 'State', 'Country', 'Similarity'],
        groupColumn: 'Review_Row_ID'
    });
    const validOutputKeys = [...new Set((wsuKeyCandidates || []).map(c => String(c.raw || '').trim()).filter(Boolean))];
    const validInputKeys = [...new Set((outcomesKeyCandidates || []).map(c => String(c.raw || '').trim()).filter(Boolean))];
    const validOutputKeysSheet = workbook.addWorksheet('Valid_Output_Keys', { state: 'hidden' });
    validOutputKeysSheet.addRow(['Key']);
    validOutputKeys.forEach(k => validOutputKeysSheet.addRow([k]));
    const validInputKeysSheet = workbook.addWorksheet('Valid_Input_Keys', { state: 'hidden' });
    validInputKeysSheet.addRow(['Key']);
    validInputKeys.forEach(k => validInputKeysSheet.addRow([k]));
    const validOutputKeysLastRow = Math.max(2, validOutputKeys.length + 1);
    const validInputKeysLastRow = Math.max(2, validInputKeys.length + 1);
    reportProgress('Building Review_Workbench...', 82);
    const reviewOutcomesKeyContextKey = keyLabels.outcomes ? `outcomes_${keyLabels.outcomes}` : 'outcomes_key';
    const reviewOutcomesNameCandidates = [
        nameCompareConfig.outcomes ? `outcomes_${nameCompareConfig.outcomes}` : '',
        ...outcomesColumns
    ].filter(Boolean);
    const reviewOutcomesNameContextKey = reviewOutcomesNameCandidates.find(key => key !== reviewOutcomesKeyContextKey) || reviewOutcomesNameCandidates[0] || 'outcomes_name';
    const reviewWsuKeyContextKey = keyLabels.wsu ? `wsu_${keyLabels.wsu}` : 'wsu_key';
    const reviewWsuNameCandidates = [
        nameCompareConfig.wsu ? `wsu_${nameCompareConfig.wsu}` : '',
        ...wsuColumns
    ].filter(Boolean);
    const reviewWsuNameContextKey = reviewWsuNameCandidates.find(key => key !== reviewWsuKeyContextKey) || reviewWsuNameCandidates[0] || 'wsu_name';
    const outcomesStateKey = outcomesStateColumn ? `outcomes_${outcomesStateColumn}` : null;
    const outcomesCountryKey = outcomesCountryColumn ? `outcomes_${outcomesCountryColumn}` : null;
    const wsuCityKey = wsuCityColumn ? `wsu_${wsuCityColumn}` : null;
    const wsuStateKey = wsuStateColumn ? `wsu_${wsuStateColumn}` : null;
    const wsuCountryKey = wsuCountryColumn ? `wsu_${wsuCountryColumn}` : null;
    const sourceContextKeys = [outcomesStateKey, outcomesCountryKey, wsuCityKey, wsuStateKey, wsuCountryKey].filter(Boolean);
    const reviewWorkbenchColumns = [
        'Decision',
        'Reason_Code',
        'Error_Type',
        'Error_Subtype',
        ...outcomesColumns.filter(k => k && k !== reviewOutcomesKeyContextKey),
        reviewOutcomesKeyContextKey,
        ...wsuColumns.filter(k => k && k !== reviewWsuKeyContextKey),
        reviewWsuKeyContextKey,
        'translate_input',
        'translate_output',
        'Selected_Candidate_ID',
        'Manual_Suggested_Key',
        'C1_Choice',
        'C2_Choice',
        'C3_Choice',
        'Suggested_Key',
        'Suggested_School',
        'Suggested_City',
        'Suggested_State',
        'Suggested_Country',
        'Current_Input',
        'Current_Output',
        'Final_Input',
        'Final_Output',
        'Decision_Warning',
        'Review_Row_ID',
        'Priority',
        'Source_Sheet',
        'Key_Update_Side',
        'Is_Stale_Key',
        'Missing_In',
        'Similarity',
        ...mappingColumns,
        'Recommended_Action',
        'Publish_Eligible',
        'Approval_Source',
        'Has_Update'
        // De-duplicate to guard against rare source-column name collisions.
    ].filter((v, i, arr) => arr.indexOf(v) === i);
    const reviewColumnWidths = {
        Decision: 20,
        Reason_Code: 24,
        Error_Type: 20,
        Error_Subtype: 20,
        translate_input: 24,
        translate_output: 24,
        Selected_Candidate_ID: 18,
        Manual_Suggested_Key: 22,
        C1_Choice: 52,
        C2_Choice: 52,
        C3_Choice: 52,
        Suggested_Key: 22,
        Suggested_School: 28,
        Suggested_City: 18,
        Suggested_State: 14,
        Suggested_Country: 14,
        Current_Input: 24,
        Current_Output: 24,
        Final_Input: 24,
        Final_Output: 24,
        Decision_Warning: 36
    };
    const hiddenReviewColumns = new Set([
        'Review_Row_ID',
        'Priority',
        'Source_Sheet',
        'Key_Update_Side',
        'Is_Stale_Key',
        'Missing_In',
        'Similarity',
        'Recommended_Action',
        'Current_Input',
        'Current_Output',
        'Publish_Eligible',
        'Approval_Source',
        'Has_Update'
    ]);
    const reviewColumnLayoutByKey = {};
    reviewWorkbenchColumns.forEach(col => {
        reviewColumnLayoutByKey[col] = {
            width: reviewColumnWidths[col] || 22,
            hidden: hiddenReviewColumns.has(col)
        };
    });
    const reviewWorkbenchHeaders = buildHeaders(reviewWorkbenchColumns, keyLabels).map((h, i) => {
        const col = reviewWorkbenchColumns[i];
        if (col === 'Review_Row_ID') return 'Review Row ID';
        if (col === 'Recommended_Action') return 'Recommended Action';
        if (col === 'Source_Sheet') return 'Source Sheet';
        if (col === 'Key_Update_Side') return 'Update Side';
        if (col === 'Is_Stale_Key') return 'Stale Key (1=yes)';
        if (col === 'Current_Input') return 'Current Translate Input';
        if (col === 'Current_Output') return 'Current Translate Output';
        if (col === 'Final_Input') return 'Final Translate Input';
        if (col === 'Final_Output') return 'Final Translate Output';
        if (col === 'Publish_Eligible') return 'Publish Eligible (1=yes)';
        if (col === 'Approval_Source') return 'Approval Source';
        if (col === 'Has_Update') return 'Has Update (1=yes)';
        if (col === 'Decision_Warning') return 'Decision Warning';
        if (col === 'Decision') return 'Decision';
        if (col === 'Reason_Code') return 'Reason Code';
        if (col === reviewOutcomesNameContextKey) return 'Outcomes Name';
        if (col === reviewOutcomesKeyContextKey) return 'Outcomes Key';
        if (outcomesStateKey && col === outcomesStateKey) return 'Outcomes State';
        if (outcomesCountryKey && col === outcomesCountryKey) return 'Outcomes Country';
        if (col === reviewWsuNameContextKey) return 'myWSU Name';
        if (col === reviewWsuKeyContextKey) return 'myWSU Key';
        if (wsuCityKey && col === wsuCityKey) return 'myWSU City';
        if (wsuStateKey && col === wsuStateKey) return 'myWSU State';
        if (wsuCountryKey && col === wsuCountryKey) return 'myWSU Country';
        if (col === 'Selected_Candidate_ID') return 'Selected Candidate ID';
        if (col === 'Manual_Suggested_Key') return 'Manual Key (override when no candidates)';
        if (col === 'C1_Choice') return 'C1 (Key: Name - City, State, Country | Score)';
        if (col === 'C2_Choice') return 'C2 (Key: Name - City, State, Country | Score)';
        if (col === 'C3_Choice') return 'C3 (Key: Name - City, State, Country | Score)';
        if (col === 'Suggested_Key') return 'Suggested Key';
        if (col === 'Suggested_School') return 'Suggested School';
        if (col === 'Suggested_City') return 'Suggested City';
        if (col === 'Suggested_State') return 'Suggested State';
        if (col === 'Suggested_Country') return 'Suggested Country';
        return h || col;
    });
    addSheetWithRows(workbook, {
        sheetName: 'Review_Workbench',
        outputColumns: reviewWorkbenchColumns,
        rows: actionQueueRows,
        style: baseStyle,
        headers: reviewWorkbenchHeaders,
        rowBorderByError: null,
        columnLayoutByKey: reviewColumnLayoutByKey
    });
    const reviewSheet = workbook.getWorksheet('Review_Workbench');
    const reviewColIndex = {};
    reviewWorkbenchColumns.forEach((col, idx) => {
        reviewColIndex[col] = idx + 1;
    });
    const reviewColLetter = {};
    Object.keys(reviewColIndex).forEach(key => {
        reviewColLetter[key] = columnIndexToLetter(reviewColIndex[key]);
    });
    const decisionListFormula = '"Keep As-Is,Use Suggestion,Allow One-to-Many,Ignore"';
    const reasonCodeListFormula = '"Campus consolidation,Data steward approved,Manual correction,Name match,Other"';
    if (reviewSheet && actionQueueRows.length > 0) {
        const rowEnd = actionQueueRows.length + 1;
        reviewSheet.dataValidations.add(
            `${reviewColLetter.Decision}2:${reviewColLetter.Decision}${rowEnd}`,
            {
                type: 'list',
                allowBlank: true,
                formulae: [decisionListFormula]
            }
        );
        reviewSheet.dataValidations.add(
            `${reviewColLetter.Reason_Code}2:${reviewColLetter.Reason_Code}${rowEnd}`,
            {
                type: 'list',
                allowBlank: true,
                formulae: [reasonCodeListFormula]
            }
        );
        for (let rowNum = 2; rowNum <= rowEnd; rowNum += 1) {
            const manCol = reviewColLetter.Manual_Suggested_Key;
            const sugCol = reviewColLetter.Suggested_Key;
            const effKeyInput = `IF($${manCol}$${rowNum}<>"",TRIM($${manCol}$${rowNum}),$${sugCol}$${rowNum})`;
            const effKeyOutput = effKeyInput;
            reviewSheet.getCell(`${reviewColLetter.Final_Input}${rowNum}`).value = {
                formula: `IF(OR($${reviewColLetter.Decision}${rowNum}="Keep As-Is",$${reviewColLetter.Decision}${rowNum}="Allow One-to-Many"),$${reviewColLetter.Current_Input}${rowNum},IF($${reviewColLetter.Decision}${rowNum}="Use Suggestion",IF(OR($${reviewColLetter.Key_Update_Side}${rowNum}="Input",$${reviewColLetter.Key_Update_Side}${rowNum}="Both"),${effKeyInput},$${reviewColLetter.Current_Input}${rowNum}),""))`
            };
            reviewSheet.getCell(`${reviewColLetter.Final_Output}${rowNum}`).value = {
                formula: `IF(OR($${reviewColLetter.Decision}${rowNum}="Keep As-Is",$${reviewColLetter.Decision}${rowNum}="Allow One-to-Many"),$${reviewColLetter.Current_Output}${rowNum},IF($${reviewColLetter.Decision}${rowNum}="Use Suggestion",IF(OR($${reviewColLetter.Key_Update_Side}${rowNum}="Output",$${reviewColLetter.Key_Update_Side}${rowNum}="Both"),${effKeyOutput},$${reviewColLetter.Current_Output}${rowNum}),""))`
            };
            reviewSheet.getCell(`${reviewColLetter.Publish_Eligible}${rowNum}`).value = {
                formula: `IF(AND(OR($${reviewColLetter.Decision}${rowNum}="Keep As-Is",$${reviewColLetter.Decision}${rowNum}="Use Suggestion",$${reviewColLetter.Decision}${rowNum}="Allow One-to-Many"),$${reviewColLetter.Final_Input}${rowNum}<>"",$${reviewColLetter.Final_Output}${rowNum}<>""),1,0)`
            };
            reviewSheet.getCell(`${reviewColLetter.Approval_Source}${rowNum}`).value = {
                formula: `IF($${reviewColLetter.Publish_Eligible}${rowNum}=1,"Review Decision","")`
            };
            reviewSheet.getCell(`${reviewColLetter.Has_Update}${rowNum}`).value = {
                formula: `IF(OR($${reviewColLetter.Current_Input}${rowNum}<>$${reviewColLetter.Final_Input}${rowNum},$${reviewColLetter.Current_Output}${rowNum}<>$${reviewColLetter.Final_Output}${rowNum}),1,0)`
            };
            const curInCol = reviewColLetter.Current_Input;
            const curOutCol = reviewColLetter.Current_Output;
            const sideCol = reviewColLetter.Key_Update_Side;
            const effKey = `IF(TRIM($${manCol}$${rowNum})<>"",TRIM($${manCol}$${rowNum}),$${sugCol}$${rowNum})`;
            const effBlank = `AND(TRIM($${manCol}$${rowNum})="",$${sugCol}$${rowNum}="")`;
            const noOpOutput = `AND(OR($${sideCol}$${rowNum}="Output",$${sideCol}$${rowNum}="Both"),${effKey}=TRIM($${curOutCol}$${rowNum}),${effKey}<>"")`;
            const noOpInput = `AND(OR($${sideCol}$${rowNum}="Input",$${sideCol}$${rowNum}="Both"),${effKey}=TRIM($${curInCol}$${rowNum}),${effKey}<>"")`;
            const invalidManualOutput = `AND(TRIM($${manCol}$${rowNum})<>"",OR($${sideCol}$${rowNum}="Output",$${sideCol}$${rowNum}="Both"),COUNTIF(Valid_Output_Keys!$A$2:$A$${validOutputKeysLastRow},TRIM($${manCol}$${rowNum}))=0)`;
            const invalidManualInput = `AND(TRIM($${manCol}$${rowNum})<>"",$${sideCol}$${rowNum}="Input",COUNTIF(Valid_Input_Keys!$A$2:$A$${validInputKeysLastRow},TRIM($${manCol}$${rowNum}))=0)`;
            const restOfWarning = `IF(AND($${reviewColLetter.Decision}${rowNum}="Use Suggestion",${effBlank}),"Use Suggestion needs selection",IF(AND($${reviewColLetter.Decision}${rowNum}="Use Suggestion",OR(${noOpOutput},${noOpInput})),"Use Suggestion key equals current value (no change)",IF(AND($${reviewColLetter.Decision}${rowNum}="Use Suggestion",OR(${invalidManualOutput},${invalidManualInput})),"Invalid manual key: not found in valid keys",IF(AND($${reviewColLetter.Decision}${rowNum}="Use Suggestion",$${reviewColLetter.Key_Update_Side}${rowNum}="None"),"Use Suggestion needs valid Update Side",IF(AND(OR($${reviewColLetter.Decision}${rowNum}="Keep As-Is",$${reviewColLetter.Decision}${rowNum}="Use Suggestion",$${reviewColLetter.Decision}${rowNum}="Allow One-to-Many"),OR($${reviewColLetter.Final_Input}${rowNum}="",$${reviewColLetter.Final_Output}${rowNum}="")),"Approved but blank final","")))))`;
            const dupPairCond = `AND($${reviewColLetter.Publish_Eligible}${rowNum}=1,$${reviewColLetter.Final_Input}${rowNum}<>"",$${reviewColLetter.Final_Output}${rowNum}<>"",COUNTIFS(Review_Workbench!$${reviewColLetter.Publish_Eligible}$2:$${reviewColLetter.Publish_Eligible}$${rowEnd},1,Review_Workbench!$${reviewColLetter.Final_Input}$2:$${reviewColLetter.Final_Input}$${rowEnd},$${reviewColLetter.Final_Input}$${rowNum},Review_Workbench!$${reviewColLetter.Final_Output}$2:$${reviewColLetter.Final_Output}$${rowEnd},$${reviewColLetter.Final_Output}$${rowNum})>1)`;
            reviewSheet.getCell(`${reviewColLetter.Decision_Warning}${rowNum}`).value = {
                formula: `IF(${dupPairCond},"Duplicate pair: set one to Ignore",${restOfWarning})`
            };
            const rowIndex = rowNum - 2;
            const row = actionQueueRows[rowIndex];
            const hasCandidates = row && Array.isArray(row._candidates) && row._candidates.length > 0;
            if (hasCandidates && candidatePoolLastRow > 1) {
                const ridCol = reviewColLetter.Review_Row_ID;
                const sidCol = reviewColLetter.Selected_Candidate_ID;
                const lookupVal = `$${ridCol}$${rowNum}&"|"&$${sidCol}$${rowNum}`;
                const poolARef = `Candidate_Pool!$A$2:$A$${candidatePoolLastRow}`;
                reviewSheet.getCell(`${reviewColLetter.Suggested_Key}${rowNum}`).value = {
                    formula: `IF($${sidCol}$${rowNum}="","",IFERROR(XLOOKUP(${lookupVal},${poolARef},Candidate_Pool!$B$2:$B$${candidatePoolLastRow},""),""))`
                };
                reviewSheet.getCell(`${reviewColLetter.Suggested_School}${rowNum}`).value = {
                    formula: `IF($${sidCol}$${rowNum}="","",IFERROR(XLOOKUP(${lookupVal},${poolARef},Candidate_Pool!$C$2:$C$${candidatePoolLastRow},""),""))`
                };
                reviewSheet.getCell(`${reviewColLetter.Suggested_City}${rowNum}`).value = {
                    formula: `IF($${sidCol}$${rowNum}="","",IFERROR(XLOOKUP(${lookupVal},${poolARef},Candidate_Pool!$D$2:$D$${candidatePoolLastRow},""),""))`
                };
                reviewSheet.getCell(`${reviewColLetter.Suggested_State}${rowNum}`).value = {
                    formula: `IF($${sidCol}$${rowNum}="","",IFERROR(XLOOKUP(${lookupVal},${poolARef},Candidate_Pool!$E$2:$E$${candidatePoolLastRow},""),""))`
                };
                reviewSheet.getCell(`${reviewColLetter.Suggested_Country}${rowNum}`).value = {
                    formula: `IF($${sidCol}$${rowNum}="","",IFERROR(XLOOKUP(${lookupVal},${poolARef},Candidate_Pool!$F$2:$F$${candidatePoolLastRow},""),""))`
                };
            }
        }
        // Workbook left unprotected so sort/filter work without restriction.
        const decCol = reviewColLetter.Decision;
        const decRef = `${decCol}2:${decCol}${rowEnd}`;
        const fullDataRef = `A2:${columnIndexToLetter(reviewWorkbenchColumns.length)}${rowEnd}`;
        const warningCol = reviewColLetter.Decision_Warning;
        const staleCol = reviewColLetter.Is_Stale_Key;
        const sourceCol = reviewColLetter.Source_Sheet;
        try {
            reviewSheet.addConditionalFormatting({
                ref: decRef,
                rules: [
                    { type: 'containsText', operator: 'containsText', text: 'Keep As-Is', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Use Suggestion', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Allow One-to-Many', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Ignore', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } }
                ]
            });
            const cfDec = reviewColLetter.Decision;
            const cfStale = reviewColLetter.Is_Stale_Key;
            const cfSource = reviewColLetter.Source_Sheet;
            const ruleStale = {
                type: 'expression',
                formulae: ['AND($' + cfDec + '2="",$' + cfStale + '2=1)'],
                style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF7ED' } } }
            };
            const ruleOneToMany = {
                type: 'expression',
                formulae: ['AND($' + cfDec + '2="",$' + cfSource + '2="One_to_Many")'],
                style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF9C3' } } }
            };
            const ruleBlankDec = {
                type: 'expression',
                formulae: ['$' + cfDec + '2=""'],
                style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF08A' } } }
            };
            reviewSheet.addConditionalFormatting({
                ref: fullDataRef,
                rules: [ruleStale, ruleOneToMany, ruleBlankDec]
            });
            reviewSheet.addConditionalFormatting({
                ref: `${warningCol}2:${warningCol}${rowEnd}`,
                rules: [
                    { type: 'containsText', operator: 'containsText', text: 'Duplicate pair', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Use Suggestion needs selection', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Invalid manual key', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'key equals current value', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF08A' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'valid Update Side', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } },
                    { type: 'containsText', operator: 'containsText', text: 'Approved but blank', style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } } } }
                ]
            });
        } catch (cfError) {
            if (typeof console !== 'undefined' && typeof console.warn === 'function') {
                console.warn('Review_Workbench conditional formatting skipped:', cfError?.message || String(cfError));
            }
        }
    }

    // Keep reviewer navigation focused on the primary validate workflow tabs.
    const hideValidateSheet = (sheetName) => {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet) {
            sheet.state = 'hidden';
        }
    };
    [
        'Errors_in_Translate',
        'Output_Not_Found_Ambiguous',
        'Output_Not_Found_No_Replacement',
        'One_to_Many',
        'Missing_Mappings',
        'High_Confidence_Matches',
        'Valid_Mappings',
        'Candidate_Reference'
    ].forEach(hideValidateSheet);

    reportProgress('Building Approved_Mappings...', 84);
    const buildAutoApprovedRow = (row, sourceSheet, sourceType) => {
        const outcomesKeyValue = keyLabels.outcomes ? row[`outcomes_${keyLabels.outcomes}`] : '';
        const wsuKeyValue = keyLabels.wsu ? row[`wsu_${keyLabels.wsu}`] : '';
        return {
            Review_Row_ID: [
                'auto',
                sanitizeIdPart(sourceType),
                sanitizeIdPart(row.translate_input || ''),
                sanitizeIdPart(row.translate_output || ''),
                sanitizeIdPart(outcomesKeyValue || ''),
                sanitizeIdPart(wsuKeyValue || '')
            ].join('|'),
            Source_Sheet: sourceSheet,
            Error_Type: sourceType,
            ...Object.fromEntries(outcomesColumns.map(col => [col, row[col] ?? ''])),
            translate_input: row.translate_input ?? '',
            translate_output: row.translate_output ?? '',
            ...Object.fromEntries(wsuColumns.map(col => [col, row[col] ?? ''])),
            Suggested_Key: row.Suggested_Key ?? '',
            Suggested_School: row.Suggested_School ?? '',
            Suggested_City: row.Suggested_City ?? '',
            Suggested_State: row.Suggested_State ?? '',
            Suggested_Country: row.Suggested_Country ?? '',
            Suggestion_Score: row.Suggestion_Score ?? '',
            Current_Input: row.translate_input ?? '',
            Current_Output: row.translate_output ?? '',
            Decision: 'Auto Approve',
            Owner: '',
            Status: 'Auto',
            Resolution_Note: sourceType === 'High_Confidence_Match'
                ? 'Auto-approved high-confidence mapping'
                : 'Auto-approved valid mapping',
            Resolved_Date: '',
            Reviewer: '',
            Review_Date: '',
            Reason_Code: 'AUTO',
            Notes: '',
            Final_Input: row.translate_input ?? '',
            Final_Output: row.translate_output ?? '',
            Publish_Eligible: 1,
            Approval_Source: sourceType === 'High_Confidence_Match' ? 'High_Confidence_Matches' : 'Valid_Mappings',
            Has_Update: 0
        };
    };
    const includeAutoApprovedRows = !missingOnlyExport;
    const autoApprovedRows = includeAutoApprovedRows
        ? [
            ...validDataRows.map(row => buildAutoApprovedRow(row, 'Valid_Mappings', 'Valid')),
            ...highConfidenceDataRows.map(row => buildAutoApprovedRow(row, 'High_Confidence_Matches', 'High_Confidence_Match'))
        ]
        : [];
    const approvedColumns = [
        'Review_Row_ID',
        'Approval_Source',
        'Source_Sheet',
        'Error_Type',
        ...outcomesColumns,
        'translate_input',
        'translate_output',
        ...wsuColumns,
        ...mappingColumns,
        ...reviewSuggestionColumns,
        'Current_Input',
        'Current_Output',
        'Decision',
        'Owner',
        'Status',
        'Resolution_Note',
        'Resolved_Date',
        'Reviewer',
        'Review_Date',
        'Reason_Code',
        'Notes',
        'Final_Input',
        'Final_Output',
        'Publish_Eligible',
        'Has_Update'
    ].filter((v, i, arr) => arr.indexOf(v) === i);
    const approvedRows = [...autoApprovedRows];
    const reviewRowCount = actionQueueRows.length;
    const cappedReviewFormulaRows = Math.min(reviewRowCount, MAX_VALIDATE_DYNAMIC_REVIEW_FORMULA_ROWS);
    const reviewLastRow = Math.max(2, reviewRowCount + 1);
    const reviewPublishCellRef = (rowNum) => `Review_Workbench!$${reviewColLetter.Publish_Eligible}$${rowNum}`;
    const reviewApprovedValue = (key, rowNum) => {
        const letter = reviewColLetter[key];
        if (!letter) return '""';
        return `IF(${reviewPublishCellRef(rowNum)}=1,Review_Workbench!$${letter}$${rowNum},"")`;
    };
    for (let outputIndex = 1; outputIndex <= cappedReviewFormulaRows; outputIndex += 1) {
        const reviewRowNum = outputIndex + 1;
        const formulaRow = {};
        approvedColumns.forEach(col => {
            if (col === 'Approval_Source') {
                formulaRow[col] = { formula: `IF(${reviewPublishCellRef(reviewRowNum)}=1,"Review Decision","")` };
                return;
            }
            if (reviewColLetter[col]) {
                formulaRow[col] = { formula: reviewApprovedValue(col, reviewRowNum) };
                return;
            }
            formulaRow[col] = '';
        });
        approvedRows.push(formulaRow);
    }
    const approvedHeaders = buildHeaders(approvedColumns, keyLabels).map((h, i) => {
        const col = approvedColumns[i];
        if (col === 'Review_Row_ID') return 'Review Row ID';
        if (col === 'Approval_Source') return 'Approval Source';
        if (col === 'Source_Sheet') return 'Source Sheet';
        if (col === 'Current_Input') return 'Current Translate Input';
        if (col === 'Current_Output') return 'Current Translate Output';
        if (col === 'Final_Input') return 'Final Translate Input';
        if (col === 'Final_Output') return 'Final Translate Output';
        if (col === 'Publish_Eligible') return 'Publish Eligible (1=yes)';
        if (col === 'Has_Update') return 'Has Update (1=yes)';
        if (col === 'Resolution_Note') return 'Resolution Note';
        if (col === 'Resolved_Date') return 'Resolved Date';
        if (col === 'Review_Date') return 'Review Date';
        if (col === 'Reason_Code') return 'Reason Code';
        return h || col;
    });
    addSheetWithRows(workbook, {
        sheetName: 'Approved_Mappings',
        outputColumns: approvedColumns,
        rows: approvedRows,
        style: baseStyle,
        headers: approvedHeaders,
        rowBorderByError: null
    });
    const approvedSheet = workbook.getWorksheet('Approved_Mappings');
    if (approvedSheet) {
        // Internal staging sheet; keep workbook navigation focused on reviewer tabs.
        approvedSheet.state = 'hidden';
    }

    reportProgress('Building Final_Translation_Table...', 87);
    const finalSheet = workbook.addWorksheet('Final_Translation_Table');
    const outcomesKeyContextKey = keyLabels.outcomes ? `outcomes_${keyLabels.outcomes}` : 'outcomes_key';
    const finalOutcomesNameCandidates = [
        nameCompareConfig.outcomes ? `outcomes_${nameCompareConfig.outcomes}` : '',
        ...outcomesColumns
    ].filter(Boolean);
    const outcomesNameContextKey = finalOutcomesNameCandidates.find(key => key !== outcomesKeyContextKey) || finalOutcomesNameCandidates[0] || 'outcomes_name';
    const wsuKeyContextKey = keyLabels.wsu ? `wsu_${keyLabels.wsu}` : 'wsu_key';
    const finalWsuNameCandidates = [
        nameCompareConfig.wsu ? `wsu_${nameCompareConfig.wsu}` : '',
        ...wsuColumns
    ].filter(Boolean);
    const wsuNameContextKey = finalWsuNameCandidates.find(key => key !== wsuKeyContextKey) || finalWsuNameCandidates[0] || 'wsu_name';
    const finalOutcomesCols = outcomesColumns
        .filter(k => k && k !== outcomesKeyContextKey)
        .map(k => ({ key: k, header: buildHeaders([k], keyLabels)[0] || k }));
    const finalWsuCols = wsuColumns
        .filter(k => k && k !== wsuKeyContextKey)
        .map(k => ({ key: k, header: buildHeaders([k], keyLabels)[0] || k }));
    const finalColumns = [
        { key: 'Review_Row_ID', header: 'Review Row ID' },
        { key: 'Decision', header: 'Decision' },
        { key: 'Source_Sheet', header: 'Source Sheet' },
        { key: 'Error_Type', header: 'Error Type' },
        ...finalOutcomesCols,
        { key: 'translate_input', header: 'Translate Input' },
        { key: 'translate_output', header: 'Translate Output' },
        ...finalWsuCols
    ].filter(Boolean);
    const finalColIndex = {};
    finalColumns.forEach((col, idx) => {
        finalColIndex[col.key] = idx + 1;
    });
    const finalColLetter = {};
    Object.keys(finalColIndex).forEach(key => {
        finalColLetter[key] = columnIndexToLetter(finalColIndex[key]);
    });
    const mapFinalSourceKey = (finalKey) => (
        finalKey === 'translate_input'
            ? 'Final_Input'
            : finalKey === 'translate_output'
                ? 'Final_Output'
                : finalKey
    );
    const getFinalValueFromRow = (row, finalKey) => {
        const sourceKey = mapFinalSourceKey(finalKey);
        if (Object.prototype.hasOwnProperty.call(row, sourceKey)) return row[sourceKey] ?? '';
        if (Object.prototype.hasOwnProperty.call(row, finalKey)) return row[finalKey] ?? '';
        return '';
    };
    const buildFinalAutoRow = (row) => finalColumns.map(col => sanitizeCellValue(getFinalValueFromRow(row, col.key)));
    const reviewPublishedCell = (rowNum) => `${reviewPublishCellRef(rowNum)}=1`;
    const reviewFinalValueFormula = (sourceKey, rowNum) => {
        const letter = reviewColLetter[sourceKey];
        if (!letter) return '';
        return { formula: `IF(${reviewPublishedCell(rowNum)},Review_Workbench!$${letter}$${rowNum},"")` };
    };
    const reviewFinalValueFormulaWithSuggestionFallback = (sourceKey, suggestionKey, suggestionWhenSide, rowNum) => {
        const letter = reviewColLetter[sourceKey];
        const sugLetter = reviewColLetter[suggestionKey];
        const decLetter = reviewColLetter.Decision;
        const sideLetter = reviewColLetter.Key_Update_Side;
        if (!letter) return '';
        const valueExpr = (decLetter && sideLetter && sugLetter)
            ? `IF(AND(Review_Workbench!$${decLetter}$${rowNum}="Use Suggestion",OR(Review_Workbench!$${sideLetter}$${rowNum}="${suggestionWhenSide}",Review_Workbench!$${sideLetter}$${rowNum}="Both")),Review_Workbench!$${sugLetter}$${rowNum},Review_Workbench!$${letter}$${rowNum})`
            : `Review_Workbench!$${letter}$${rowNum}`;
        return { formula: `IF(${reviewPublishedCell(rowNum)},${valueExpr},"")` };
    };
    const roleToSuggestion = { School: 'Suggested_School', City: 'Suggested_City', State: 'Suggested_State', Country: 'Suggested_Country' };
    const buildContextFallbacksForCols = (cols, rolesKey, prefix, side) => {
        const fallbacks = {};
        const roles = columnRoles[rolesKey] || {};
        cols.forEach(col => {
            const baseCol = col.key.replace(new RegExp(`^${prefix}_`), '');
            const role = roles[baseCol];
            const sugKey = role && roleToSuggestion[role];
            if (sugKey) fallbacks[col.key] = [sugKey, side];
        });
        return fallbacks;
    };
    const outcomesContextFallbacks = buildContextFallbacksForCols(finalOutcomesCols, 'outcomes', 'outcomes', 'Input');
    const wsuContextFallbacks = buildContextFallbacksForCols(finalWsuCols, 'wsu_org', 'wsu', 'Output');
    if (outcomesNameContextKey && !outcomesContextFallbacks[outcomesNameContextKey]) {
        outcomesContextFallbacks[outcomesNameContextKey] = ['Suggested_School', 'Input'];
    }
    if (wsuNameContextKey && !wsuContextFallbacks[wsuNameContextKey]) {
        wsuContextFallbacks[wsuNameContextKey] = ['Suggested_School', 'Output'];
    }
    const contextFallbacks = { ...outcomesContextFallbacks, ...wsuContextFallbacks };
    const finalFormulaRows = autoApprovedRows.length + cappedReviewFormulaRows;

    const stagingSheet = workbook.addWorksheet('Final_Staging');
    stagingSheet.state = 'hidden';
    stagingSheet.addRow(finalColumns.map(col => col.header));
    autoApprovedRows.forEach(row => stagingSheet.addRow(buildFinalAutoRow(row)));
    for (let reviewIndex = 1; reviewIndex <= cappedReviewFormulaRows; reviewIndex += 1) {
        const reviewRowNum = reviewIndex + 1;
        const rowValues = finalColumns.map(col => {
            const sourceColKey = mapFinalSourceKey(col.key);
            const fallback = contextFallbacks[col.key];
            if (fallback) {
                return reviewFinalValueFormulaWithSuggestionFallback(sourceColKey, fallback[0], fallback[1], reviewRowNum);
            }
            return reviewFinalValueFormula(sourceColKey, reviewRowNum);
        });
        stagingSheet.addRow(rowValues);
    }

    finalSheet.addRow(finalColumns.map(col => col.header));
    finalSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B7285' } };
    });
    const qaCap = 10000;
    const stagingLastRow = Math.max(2, finalFormulaRows + 1);
    const refLastRow = Math.max(qaCap, stagingLastRow);
    const stagingLastCol = columnIndexToLetter(finalColumns.length);
    const filterCol = finalColLetter.translate_input;
    const filterFormula = `_xlfn._xlws.FILTER(Final_Staging!A2:${stagingLastCol}${stagingLastRow},Final_Staging!$${filterCol}$2:$${filterCol}$${stagingLastRow}<>"","")`;
    finalSheet.getCell('A2').value = {
        formula: filterFormula,
        shareType: 'array',
        ref: `A2:${stagingLastCol}${refLastRow}`
    };

    const finalColumnWidths = {
        translate_input: 24,
        translate_output: 24,
        Review_Row_ID: 56,
        Decision: 22
    };
    sourceContextKeys.forEach(k => { finalColumnWidths[k] = 18; });
    finalSheet.columns = finalColumns.map(col => ({
        width: finalColumnWidths[col.key] || 22,
        hidden: ['Decision', 'Source_Sheet', 'Error_Type'].includes(col.key)
    }));
    finalSheet.autoFilter = {
        from: 'A1',
        to: `${stagingLastCol}${refLastRow}`
    };

    reportProgress('Building Translation_Key_Updates...', 89);
    const updatesSheet = workbook.addWorksheet('Translation_Key_Updates');
    const updatesColumns = [
        { key: 'Review_Row_ID', header: 'Review Row ID' },
        { key: 'Current_Input', header: 'Current Translate Input' },
        { key: 'Current_Output', header: 'Current Translate Output' },
        { key: 'Final_Input', header: 'Final Translate Input' },
        { key: 'Final_Output', header: 'Final Translate Output' },
        { key: 'Decision', header: 'Decision' },
        { key: 'Reason_Code', header: 'Reason Code' },
        { key: 'Source_Sheet', header: 'Source Sheet' },
        { key: 'Owner', header: 'Owner' },
        { key: 'Resolution_Note', header: 'Resolution Note' }
    ];
    updatesSheet.addRow(updatesColumns.map(col => col.header));
    updatesSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7C2D12' } };
    });
    const updateSourceValueFormula = (sourceKey, rowNum, includeExpr) => {
        const letter = reviewColLetter[sourceKey];
        if (!letter) return '';
        return { formula: `IF(${includeExpr},Review_Workbench!$${letter}$${rowNum},"")` };
    };
    // Has_Update is reviewer-driven; auto-approved rows are intentionally excluded (always 0).
    const updateFormulaRows = Math.max(1, cappedReviewFormulaRows);
    for (let outputIndex = 1; outputIndex <= updateFormulaRows; outputIndex += 1) {
        const rowNum = outputIndex + 1;
        const includeExpr = `AND(${reviewPublishedCell(rowNum)},Review_Workbench!$${reviewColLetter.Has_Update}$${rowNum}=1)`;
        updatesSheet.addRow([
            updateSourceValueFormula('Review_Row_ID', rowNum, includeExpr),
            updateSourceValueFormula('Current_Input', rowNum, includeExpr),
            updateSourceValueFormula('Current_Output', rowNum, includeExpr),
            updateSourceValueFormula('Final_Input', rowNum, includeExpr),
            updateSourceValueFormula('Final_Output', rowNum, includeExpr),
            updateSourceValueFormula('Decision', rowNum, includeExpr),
            updateSourceValueFormula('Reason_Code', rowNum, includeExpr),
            updateSourceValueFormula('Source_Sheet', rowNum, includeExpr),
            updateSourceValueFormula('Owner', rowNum, includeExpr),
            updateSourceValueFormula('Resolution_Note', rowNum, includeExpr)
        ]);
    }
    updatesSheet.columns = [
        { width: 56 },
        { width: 24 },
        { width: 24 },
        { width: 24 },
        { width: 24 },
        { width: 22 },
        { width: 24 },
        { width: 22 },
        { width: 18 },
        { width: 46 }
    ];

    reportProgress('Building QA_Checks_Validate...', 91);
    const qaValidateSheet = workbook.addWorksheet('QA_Checks_Validate');
    const decisionRange = `Review_Workbench!$${reviewColLetter.Decision}$2:$${reviewColLetter.Decision}$${reviewLastRow}`;
    const reviewPublishRangeRef = `Review_Workbench!$${reviewColLetter.Publish_Eligible}$2:$${reviewColLetter.Publish_Eligible}$${reviewLastRow}`;
    const reviewFinalInputRange = `Review_Workbench!$${reviewColLetter.Final_Input}$2:$${reviewColLetter.Final_Input}$${reviewLastRow}`;
    const reviewFinalOutputRange = `Review_Workbench!$${reviewColLetter.Final_Output}$2:$${reviewColLetter.Final_Output}$${reviewLastRow}`;
    const reviewStaleRange = `Review_Workbench!$${reviewColLetter.Is_Stale_Key}$2:$${reviewColLetter.Is_Stale_Key}$${reviewLastRow}`;
    const reviewSourceRange = `Review_Workbench!$${reviewColLetter.Source_Sheet}$2:$${reviewColLetter.Source_Sheet}$${reviewLastRow}`;
    const reviewErrorTypeRange = `Review_Workbench!$${reviewColLetter.Error_Type}$2:$${reviewColLetter.Error_Type}$${reviewLastRow}`;
    const reviewSuggestedKeyRange = `Review_Workbench!$${reviewColLetter.Suggested_Key}$2:$${reviewColLetter.Suggested_Key}$${reviewLastRow}`;
    const reviewManualSuggestedKeyRange = `Review_Workbench!$${reviewColLetter.Manual_Suggested_Key}$2:$${reviewColLetter.Manual_Suggested_Key}$${reviewLastRow}`;
    const reviewReasonCodeRange = `Review_Workbench!$${reviewColLetter.Reason_Code}$2:$${reviewColLetter.Reason_Code}$${reviewLastRow}`;
    const reviewDecisionWarningRange = `Review_Workbench!$${reviewColLetter.Decision_Warning}$2:$${reviewColLetter.Decision_Warning}$${reviewLastRow}`;
    const reviewKeyUpdateSideRange = `Review_Workbench!$${reviewColLetter.Key_Update_Side}$2:$${reviewColLetter.Key_Update_Side}$${reviewLastRow}`;
    const finalLastRow = refLastRow;
    const finalInputRange = `Final_Translation_Table!$${finalColLetter.translate_input}$2:$${finalColLetter.translate_input}$${finalLastRow}`;
    const finalOutputRange = `Final_Translation_Table!$${finalColLetter.translate_output}$2:$${finalColLetter.translate_output}$${finalLastRow}`;
    const finalDecisionRange = `Final_Translation_Table!$${finalColLetter.Decision}$2:$${finalColLetter.Decision}$${finalLastRow}`;
    const finalSourceSheetRange = `Final_Translation_Table!$${finalColLetter.Source_Sheet}$2:$${finalColLetter.Source_Sheet}$${finalLastRow}`;
    const finalErrorTypeRange = `Final_Translation_Table!$${finalColLetter.Error_Type}$2:$${finalColLetter.Error_Type}$${finalLastRow}`;
    const getQAEmptyRows = (typeof ValidationExportHelpers !== 'undefined' && ValidationExportHelpers &&
        typeof ValidationExportHelpers.getQAValidateRowsForEmptyQueue === 'function')
        ? ValidationExportHelpers.getQAValidateRowsForEmptyQueue
        : () => [
            ['Check', 'Count', 'Status', 'Detail'],
            ['Unresolved actions', 0, 'PASS', 'Rows without a decision'],
            ['Approved for update', 0, 'PASS', 'Keep As-Is or Use Suggestion decisions'],
            ['Stale-key rows lacking decision', 0, 'PASS', 'Likely stale key rows without decision'],
            ['Duplicate conflict rows lacking decision', 0, 'PASS', 'One-to-many rows without decision']
        ];
    const qaRows = actionQueueRows.length > 0
        ? [
            ['Check', 'Count', 'Status', 'Detail'],
            ['Unresolved actions', `=COUNTIF(${decisionRange},"")+COUNTIF(${decisionRange},"Ignore")`, '=IF(B2=0,"PASS","CHECK")', 'Blank or Ignore'],
            ['Approved review rows', `=COUNTIF(${decisionRange},"Keep As-Is")+COUNTIF(${decisionRange},"Use Suggestion")+COUNTIF(${decisionRange},"Allow One-to-Many")`, '=IF(B3>0,"PASS","CHECK")', 'Rows approved from Review_Workbench'],
            ['Approved rows beyond formula capacity', `=MAX(0,B3-${cappedReviewFormulaRows})`, '=IF(B4=0,"PASS","CHECK")', `Rows above ${cappedReviewFormulaRows} approved review decisions exceed formula row capacity`],
            ['Blank final keys on publish-eligible rows (sanity)', `=COUNTIFS(${reviewPublishRangeRef},1,${reviewFinalInputRange},"")+COUNTIFS(${reviewPublishRangeRef},1,${reviewFinalOutputRange},"")`, '=IF(B5=0,"PASS","FAIL")', 'Sanity check: publish-eligible rows should already enforce non-blank finals'],
            ['Use Suggestion without effective key', `=COUNTIFS(${decisionRange},"Use Suggestion",${reviewManualSuggestedKeyRange},"",${reviewSuggestedKeyRange},"")`, '=IF(B6=0,"PASS","FAIL")', 'Use Suggestion needs Manual Key or Selected Candidate ID + Suggested Key'],
            ['Use Suggestion with invalid Update Side', `=COUNTIFS(${decisionRange},"Use Suggestion",${reviewKeyUpdateSideRange},"None")`, '=IF(B7=0,"PASS","FAIL")', 'Use Suggestion chosen but Update Side is None; fix or change decision'],
            ['Use Suggestion with invalid manual key', `=COUNTIF(${reviewDecisionWarningRange},"*Invalid manual key*")`, '=IF(B8=0,"PASS","FAIL")', 'Manual key not found in valid myWSU/Outcomes keys'],
            ['Use Suggestion no-op (key equals current value)', `=COUNTIF(${reviewDecisionWarningRange},"*no change*")`, '=IF(B9=0,"PASS","FAIL")', 'Use Suggestion chosen but effective key equals current; fix or change decision'],
            ['Risky decisions without reason code', `=SUMPRODUCT((1*(((${decisionRange}="Use Suggestion")*(TRIM(${reviewManualSuggestedKeyRange})<>""))+(${decisionRange}="Allow One-to-Many")+(((${decisionRange}="Keep As-Is")*(${reviewSourceRange}="One_to_Many")*(${reviewErrorTypeRange}="Duplicate_Target")))>0))*(TRIM(${reviewReasonCodeRange})=""))`, '=IF(B10=0,"PASS","FAIL")', 'Reason Code required for: Use Suggestion+Manual Key, Allow One-to-Many, Duplicate_Target+Keep As-Is'],
            ['Stale-key rows lacking decision', `=COUNTIFS(${reviewStaleRange},1,${decisionRange},"")`, '=IF(B11=0,"PASS","CHECK")', 'Likely stale key rows without decision (advisory)'],
            ['One-to-many rows lacking decision', `=COUNTIFS(${reviewSourceRange},"One_to_Many",${decisionRange},"")`, '=IF(B12=0,"PASS","CHECK")', 'One-to-many rows without decision (advisory)'],
            ['Duplicate final input keys (excluding Allow One-to-Many)', `=SUMPRODUCT((${finalInputRange}<>"")*(${finalDecisionRange}<>"Allow One-to-Many")*(COUNTIFS(${finalInputRange},${finalInputRange},${finalDecisionRange},"<>Allow One-to-Many")>1))/2`, '=IF(B13=0,"PASS","CHECK")', 'Duplicates in Final_Translation_Table input keys excluding approved one-to-many rows'],
            ['Duplicate final output keys (excluding Allow One-to-Many)', `=SUMPRODUCT((${finalOutputRange}<>"")*(${finalDecisionRange}<>"Allow One-to-Many")*(((${finalSourceSheetRange}<>"One_to_Many")+(${finalErrorTypeRange}<>"Duplicate_Target")+(${finalDecisionRange}<>"Keep As-Is"))>0)*(COUNTIFS(${finalOutputRange},${finalOutputRange},${finalDecisionRange},"<>Allow One-to-Many")-COUNTIFS(${finalOutputRange},${finalOutputRange},${finalSourceSheetRange},"One_to_Many",${finalErrorTypeRange},"Duplicate_Target",${finalDecisionRange},"Keep As-Is")>1))/2`, '=IF(B14=0,"PASS","CHECK")', 'Duplicates in Final_Translation_Table output keys excluding Allow One-to-Many and Duplicate_Target+Keep As-Is'],
            ['Duplicate (input, output) pairs in Final_Translation_Table', `=SUMPRODUCT((${finalInputRange}<>"")*(${finalOutputRange}<>"")*(COUNTIFS(${finalInputRange},${finalInputRange},${finalOutputRange},${finalOutputRange})>1))/2`, '=IF(B15=0,"PASS","FAIL")', 'Duplicate (input,output) pairs; set one to Ignore before publish'],
            ...(reimportSummary.applied > 0 ? [['Prior decisions applied', reimportSummary.applied, '', 'Re-import from prior Validate workbook']] : []),
            ['Publish gate', `=IF(AND(B2=0,B4=0,B5=0,B6=0,B7=0,B8=0,B9=0,B10=0,B13=0,B14=0,B15=0),"PASS","HOLD")`, '', 'Final publish gate (B11/B12 advisory). Diagnostic tabs are hidden; right-click tab bar and choose Unhide if needed.']
        ]
        : getQAEmptyRows();
    qaRows.forEach(row => qaValidateSheet.addRow(row));
    qaValidateSheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
    });
    qaValidateSheet.columns = [{ width: 42 }, { width: 26 }, { width: 14 }, { width: 70 }];
    if (canSuggestNames && typeof console !== 'undefined' && typeof console.log === 'function') {
        const forwardFallbackRate = suggestionBlockStats.forwardQueries > 0
            ? suggestionBlockStats.forwardFallbacks / suggestionBlockStats.forwardQueries
            : 0;
        const reverseFallbackRate = suggestionBlockStats.reverseQueries > 0
            ? suggestionBlockStats.reverseFallbacks / suggestionBlockStats.reverseQueries
            : 0;
        console.log(
            'Suggestion blocking stats:',
            JSON.stringify({
                forwardQueries: suggestionBlockStats.forwardQueries,
                reverseQueries: suggestionBlockStats.reverseQueries,
                forwardFallbacks: suggestionBlockStats.forwardFallbacks,
                reverseFallbacks: suggestionBlockStats.reverseFallbacks,
                forwardFallbackRate: Number(forwardFallbackRate.toFixed(4)),
                reverseFallbackRate: Number(reverseFallbackRate.toFixed(4))
            })
        );
    }

    // Open workbook focused on Review_Workbench so it is the first visible reviewer tab.
    const workbookSheets = Array.isArray(workbook.worksheets)
        ? workbook.worksheets
        : (Array.isArray(workbook._worksheets) ? workbook._worksheets : []);
    const reviewSheetIndex = workbookSheets.findIndex(sheet => sheet && sheet.name === 'Review_Workbench');
    if (reviewSheetIndex >= 0) {
        workbook.views = [{ activeTab: reviewSheetIndex, firstSheet: reviewSheetIndex }];
    }

    reportProgress('Finalizing Excel file...', 92);
    const buffer = toArrayBuffer(await workbook.xlsx.writeBuffer());
    reportProgress('Saving file...', 100);
    const result = {
        buffer,
        filename: downloadSafeFileName(options.fileName, 'WSU_Mapping_Validation_Report.xlsx')
    };
    if (priorDecisions) result.reimportSummary = reimportSummary;
    return result;
}

async function buildJoinPreviewExport(payload) {
    const { selectedCols = {}, options = {}, context = {} } = payload || {};
    const loadedData = context.loadedData || { outcomes: [], translate: [], wsu_org: [] };
    const keyConfig = context.keyConfig || {};
    const keyLabels = context.keyLabels || {};

    const outcomesRows = loadedData.outcomes || [];
    const translateRows = loadedData.translate || [];
    const wsuRows = loadedData.wsu_org || [];

    const outcomesKeyCol = keyConfig.outcomes || '';
    const translateInputCol = keyConfig.translateInput || '';
    const translateOutputCol = keyConfig.translateOutput || '';
    const wsuKeyCol = keyConfig.wsu || '';

    const outcomesCols = (selectedCols.outcomes || []).filter(Boolean);
    const wsuCols = (selectedCols.wsu_org || []).filter(Boolean);

    const outcomesMap = new Map();
    outcomesRows.forEach((row, idx) => {
        const raw = row[outcomesKeyCol];
        const norm = typeof normalizeKeyValue === 'function' ? normalizeKeyValue(raw) : String(raw || '').trim();
        if (!norm) return;
        if (!outcomesMap.has(norm)) outcomesMap.set(norm, { row, idx });
    });

    const wsuMap = new Map();
    wsuRows.forEach((row, idx) => {
        const raw = row[wsuKeyCol];
        const norm = typeof normalizeKeyValue === 'function' ? normalizeKeyValue(raw) : String(raw || '').trim();
        if (!norm) return;
        if (!wsuMap.has(norm)) wsuMap.set(norm, { row, idx });
    });

    const outcomesColKeys = outcomesCols.map(c => `outcomes_${c}`);
    const wsuColKeys = wsuCols.map(c => `wsu_${c}`);
    const outputColumns = [...outcomesColKeys, 'translate_input', 'translate_output', ...wsuColKeys];

    const joinRows = translateRows.map(tr => {
        const inputRaw = tr[translateInputCol];
        const outputRaw = tr[translateOutputCol];
        const inputNorm = typeof normalizeKeyValue === 'function' ? normalizeKeyValue(inputRaw) : String(inputRaw || '').trim();
        const outputNorm = typeof normalizeKeyValue === 'function' ? normalizeKeyValue(outputRaw) : String(outputRaw || '').trim();

        const outcomesEntry = inputNorm ? outcomesMap.get(inputNorm) : null;
        const wsuEntry = outputNorm ? wsuMap.get(outputNorm) : null;

        const outcomesRow = outcomesEntry ? outcomesEntry.row : {};
        const wsuRow = wsuEntry ? wsuEntry.row : {};

        const row = {};
        outcomesCols.forEach(col => {
            row[`outcomes_${col}`] = outcomesRow[col] ?? '';
        });
        row.translate_input = inputRaw ?? '';
        row.translate_output = outputRaw ?? '';
        wsuCols.forEach(col => {
            row[`wsu_${col}`] = wsuRow[col] ?? '';
        });
        return row;
    });

    const workbook = new ExcelJS.Workbook();
    workbook.calcProperties = { fullCalcOnLoad: true };

    const baseStyle = {
        defaultHeaderColor: 'FF1E3A8A',
        defaultBodyColor: 'FFF3F4F6',
        outcomesHeaderColor: 'FF1E3A8A',
        outcomesBodyColor: 'FFF3F4F6',
        wsuHeaderColor: 'FF1E3A8A',
        wsuBodyColor: 'FFF3F4F6',
        translateHeaderColor: 'FF0F766E',
        translateBodyColor: 'FFD1FAE5'
    };

    const headers = outputColumns.map(col => {
        if (col === 'translate_input') return `${keyLabels.translateInput || 'Source key'} (Translate Input)`;
        if (col === 'translate_output') return `${keyLabels.translateOutput || 'Target key'} (Translate Output)`;
        return buildHeaders([col], keyLabels)[0] || col;
    });

    addSheetWithRows(workbook, {
        sheetName: 'Join_Preview',
        outputColumns,
        rows: joinRows,
        style: baseStyle,
        headers
    });

    const buffer = toArrayBuffer(await workbook.xlsx.writeBuffer());
    return {
        buffer,
        filename: downloadSafeFileName(options.fileName, 'WSU_Join_Preview.xlsx')
    };
}

self.onmessage = async (event) => {
    const { type, payload } = event.data || {};
    try {
        if (type === 'build_generation_export') {
            const result = await buildGenerationExport(payload || {});
            self.postMessage({ type: 'result', result }, [result.buffer]);
            return;
        }
        if (type === 'build_validation_export') {
            const result = await buildValidationExport(payload || {});
            if (result.actionQueueRows && !result.buffer) {
                self.postMessage({ type: 'result', result });
            } else {
                self.postMessage({ type: 'result', result }, [result.buffer]);
            }
            return;
        }
        if (type === 'get_action_queue') {
            const result = await buildValidationExport({ ...(payload || {}), returnActionQueueOnly: true });
            self.postMessage({ type: 'result', result });
            return;
        }
        if (type === 'build_join_preview_export') {
            const result = await buildJoinPreviewExport(payload || {});
            self.postMessage({ type: 'result', result }, [result.buffer]);
            return;
        }
        throw new Error(`Unknown export task: ${type}`);
    } catch (error) {
        let msg = error?.message || String(error);
        if (msg.includes("reading '0'")) {
            msg += ' (cell values may contain objects; try exporting with fewer columns or check source data)';
        }
        self.postMessage({
            type: 'error',
            message: msg,
            stack: error?.stack || ''
        });
    }
};
