/* global mergeData, validateMappings, detectMissingMappings, generateSummaryStats, normalizeKeyValue, calculateNameSimilarity, isHighConfidenceNameMatch, statesMatch, countriesMatch */
importScripts('validation.js');

const WORKER_NAME_MATCH_AMBIGUITY_GAP = 0.03;
const PROGRESS_MIN_INTERVAL_MS = 400;
const progressState = {};

function reportProgress(stage, processed, total) {
    const now = Date.now();
    const last = progressState[stage] || 0;
    if (processed === 0 || processed === total || now - last >= PROGRESS_MIN_INTERVAL_MS) {
        self.postMessage({ type: 'progress', stage, processed, total });
        progressState[stage] = now;
    }
}

function buildKeyValueMap(rows, keyField, datasetLabel) {
    const map = new Map();
    const duplicateCounts = new Map();
    rows.forEach(row => {
        const raw = row[keyField];
        const normalized = normalizeKeyValue(raw);
        if (!normalized) {
            return;
        }
        if (map.has(normalized)) {
            duplicateCounts.set(normalized, (duplicateCounts.get(normalized) || 1) + 1);
            return;
        }
        map.set(normalized, row);
    });

    if (duplicateCounts.size > 0) {
        const duplicateKeys = Array.from(duplicateCounts.entries())
            .sort((a, b) => b[1] - a[1]);
        const sample = duplicateKeys
            .slice(0, 5)
            .map(([key, count]) => `${key} (${count})`)
            .join(', ');
        throw new Error(
            `${datasetLabel} has duplicate key values in "${keyField}" (${duplicateCounts.size} duplicate keys). ` +
            `Examples: ${sample}`
        );
    }

    return map;
}

function resolveLocationValue(row, fieldName) {
    if (!row || !fieldName) return '';
    return row[fieldName] ?? '';
}

function buildLocationContext(row, locationFields) {
    return {
        state: resolveLocationValue(row, locationFields.state),
        city: resolveLocationValue(row, locationFields.city),
        country: resolveLocationValue(row, locationFields.country)
    };
}

function hasLocationSignal(context) {
    return Boolean(context.state || context.city || context.country);
}

function hasComparableState(state1, state2) {
    const normalized1 = String(state1 || '').trim().toLowerCase();
    const normalized2 = String(state2 || '').trim().toLowerCase();
    return Boolean(normalized1 && normalized2 && normalized1 !== 'ot' && normalized2 !== 'ot');
}

function shouldBlockByLocation(sourceContext, targetContext) {
    if (
        sourceContext.country && targetContext.country &&
        !countriesMatch(sourceContext.country, targetContext.country)
    ) {
        return true;
    }

    if (
        hasComparableState(sourceContext.state, targetContext.state) &&
        sourceContext.country &&
        targetContext.country &&
        countriesMatch(sourceContext.country, targetContext.country) &&
        !statesMatch(sourceContext.state, targetContext.state)
    ) {
        return true;
    }

    return false;
}

function passesNameGate(sourceName, targetName, sourceContext, targetContext, score, threshold) {
    const hasContext = hasLocationSignal(sourceContext) || hasLocationSignal(targetContext);
    const highConfidence = isHighConfidenceNameMatch(
        sourceName,
        targetName,
        sourceContext.state,
        targetContext.state,
        sourceContext.city,
        targetContext.city,
        sourceContext.country,
        targetContext.country,
        score,
        threshold
    );
    if (highConfidence) {
        return true;
    }
    if (score < threshold) {
        return false;
    }
    if (!hasContext) {
        return true;
    }

    const strictScore = Math.max(0.9, threshold + 0.08);
    return score >= strictScore;
}

function normalizeSimilarityPercent(score) {
    if (!Number.isFinite(score)) {
        return '';
    }
    return Math.round(Math.max(0, Math.min(1, score)) * 1000) / 10;
}

function resolveConfidenceTier(similarityPercent, threshold, highConfidence = false) {
    if (!Number.isFinite(similarityPercent)) {
        return '';
    }
    const thresholdPercent = Math.max(0, Math.min(1, threshold)) * 100;
    if (highConfidence || similarityPercent >= 90) {
        return 'high';
    }
    if (similarityPercent >= 80) {
        return 'medium';
    }
    if (similarityPercent >= thresholdPercent) {
        return 'low';
    }
    return 'low';
}

function findBestNameMatch(
    sourceRow,
    sourceNameField,
    sourceLocationFields,
    targetEntries,
    targetNameField,
    targetLocationFields,
    threshold,
    usedKeys,
    minGap = WORKER_NAME_MATCH_AMBIGUITY_GAP
) {
    const sourceName = sourceRow?.[sourceNameField];
    if (!sourceName) return null;
    const sourceContext = buildLocationContext(sourceRow, sourceLocationFields);
    let best = null;
    let bestScore = -1;
    let secondBestScore = -1;
    const topCandidates = [];
    targetEntries.forEach(({ key, row }) => {
        if (usedKeys.has(key)) return;
        const targetName = row[targetNameField];
        if (!targetName) return;
        const targetContext = buildLocationContext(row, targetLocationFields);
        if (shouldBlockByLocation(sourceContext, targetContext)) return;
        const score = calculateNameSimilarity(sourceName, targetName);
        if (!passesNameGate(sourceName, targetName, sourceContext, targetContext, score, threshold)) {
            return;
        }
        topCandidates.push({ key, row, score });
        topCandidates.sort((a, b) => b.score - a.score);
        if (topCandidates.length > 3) {
            topCandidates.pop();
        }
        if (score > bestScore) {
            secondBestScore = bestScore;
            bestScore = score;
            best = { key, row, score };
        } else if (score > secondBestScore) {
            secondBestScore = score;
        }
    });
    if (!best) {
        return null;
    }
    const effectiveGap = bestScore >= 0.9 ? minGap : (minGap + 0.01);
    if (secondBestScore >= 0 && bestScore - secondBestScore < effectiveGap) {
        return {
            ambiguous: true,
            bestScore,
            secondBestScore,
            candidate: best,
            topCandidates
        };
    }
    return best;
}

function collectTopNameCandidates(
    sourceRows,
    targetRows,
    sourceNameField,
    targetNameField,
    sourceLocationFields,
    targetLocationFields,
    threshold,
    maxPerSource = 3,
    minGap = WORKER_NAME_MATCH_AMBIGUITY_GAP,
    onProgress = null
) {
    const candidates = [];
    const ambiguousSources = new Set();
    const ambiguousCandidatesBySource = new Map();

    sourceRows.forEach((sourceRow, sourceIndex) => {
        if (onProgress) {
            onProgress(sourceIndex + 1, sourceRows.length);
        }
        const sourceName = sourceRow[sourceNameField];
        if (!sourceName) {
            return;
        }
        const sourceContext = buildLocationContext(sourceRow, sourceLocationFields);

        const topMatches = [];
        targetRows.forEach((targetRow, targetIndex) => {
            const targetName = targetRow[targetNameField];
            if (!targetName) {
                return;
            }
            const targetContext = buildLocationContext(targetRow, targetLocationFields);
            if (shouldBlockByLocation(sourceContext, targetContext)) {
                return;
            }
            const score = calculateNameSimilarity(sourceName, targetName);
            if (!passesNameGate(sourceName, targetName, sourceContext, targetContext, score, threshold)) {
                return;
            }

            topMatches.push({ sourceIndex, targetIndex, score });
            topMatches.sort((a, b) => b.score - a.score);
            if (topMatches.length > maxPerSource) {
                topMatches.pop();
            }
        });

        const effectiveGap = topMatches[0]?.score >= 0.9 ? minGap : (minGap + 0.01);
        if (
            topMatches.length > 1 &&
            (topMatches[0].score - topMatches[1].score) < effectiveGap
        ) {
            ambiguousSources.add(sourceIndex);
            ambiguousCandidatesBySource.set(
                sourceIndex,
                topMatches.map(match => ({
                    sourceIndex: match.sourceIndex,
                    targetIndex: match.targetIndex,
                    score: match.score
                }))
            );
            return;
        }
        topMatches.forEach(match => candidates.push(match));
    });

    candidates.sort(
        (a, b) => b.score - a.score || a.sourceIndex - b.sourceIndex || a.targetIndex - b.targetIndex
    );
    return { candidates, ambiguousSources, ambiguousCandidatesBySource };
}

function assignGlobalNameMatches(
    sourceRows,
    targetRows,
    sourceNameField,
    targetNameField,
    sourceLocationFields,
    targetLocationFields,
    threshold,
    minGap = WORKER_NAME_MATCH_AMBIGUITY_GAP,
    onProgress = null
) {
    const { candidates, ambiguousSources, ambiguousCandidatesBySource } = collectTopNameCandidates(
        sourceRows,
        targetRows,
        sourceNameField,
        targetNameField,
        sourceLocationFields,
        targetLocationFields,
        threshold,
        4,
        minGap,
        onProgress
    );

    const matchesBySource = new Map();
    const usedTargets = new Set();

    candidates.forEach(candidate => {
        if (matchesBySource.has(candidate.sourceIndex)) {
            return;
        }
        if (usedTargets.has(candidate.targetIndex)) {
            return;
        }
        matchesBySource.set(candidate.sourceIndex, {
            targetIndex: candidate.targetIndex,
            score: candidate.score
        });
        usedTargets.add(candidate.targetIndex);
    });

    return { matchesBySource, usedTargets, ambiguousSources, ambiguousCandidatesBySource };
}

function generateTranslationTableWorker(outcomes, wsuOrg, keyConfig, nameCompare, options, selectedColumns, keyLabels) {
    const nameCompareEnabled = Boolean(nameCompare.enabled);
    const outcomesNameField = nameCompare.outcomes_column || '';
    const wsuNameField = nameCompare.wsu_column || '';
    const threshold = typeof nameCompare.threshold === 'number' ? nameCompare.threshold : 0.8;
    const ambiguityGap = typeof nameCompare.ambiguity_gap === 'number'
        ? nameCompare.ambiguity_gap
        : WORKER_NAME_MATCH_AMBIGUITY_GAP;
    const outcomesLocationFields = {
        state: nameCompare.state_outcomes || '',
        city: nameCompare.city_outcomes || '',
        country: nameCompare.country_outcomes || ''
    };
    const wsuLocationFields = {
        state: nameCompare.state_wsu || '',
        city: nameCompare.city_wsu || '',
        country: nameCompare.country_wsu || ''
    };
    const canNameMatch = nameCompareEnabled && outcomesNameField && wsuNameField;
    const forceNameMatch = Boolean(options.forceNameMatch);
    const outcomesKeyField = keyConfig.outcomes || '';
    const wsuKeyField = keyConfig.wsu || '';
    const outcomesDisplayField = outcomesNameField || outcomesKeyField || selectedColumns.outcomes[0] || '';
    const wsuDisplayField = wsuNameField || wsuKeyField || selectedColumns.wsu_org[0] || '';

    const headerLabels = {
        input: forceNameMatch
            ? (outcomesNameField || keyLabels.outcomes || 'Outcomes Name')
            : (keyLabels.outcomes || outcomesNameField || 'Outcomes Key'),
        output: forceNameMatch
            ? (wsuNameField || keyLabels.wsu || 'myWSU Name')
            : (keyLabels.wsu || wsuNameField || 'myWSU Key')
    };
    const generationConfig = {
        threshold: Math.max(0, Math.min(1, threshold)),
        outcomesNameField,
        wsuNameField,
        outcomesKeyField,
        wsuKeyField
    };

    const cleanRows = [];
    const errorRows = [];
    const outcomesRowToIndex = new Map();
    const wsuRowToIndex = new Map();
    outcomes.forEach((row, index) => outcomesRowToIndex.set(row, index));
    wsuOrg.forEach((row, index) => wsuRowToIndex.set(row, index));

    const toText = (value) => (value === null || value === undefined ? '' : String(value).trim());

    const buildAlternateCandidatesFromIndexMatches = (matches, targetRows, targetNameField, targetKeyField) => (
        (matches || []).map((candidate, idx) => {
            const targetRow = targetRows[candidate.targetIndex];
            return {
                rank: idx + 1,
                key: targetKeyField ? (targetRow?.[targetKeyField] ?? '') : '',
                name: targetNameField ? (targetRow?.[targetNameField] ?? '') : '',
                similarity: normalizeSimilarityPercent(candidate.score)
            };
        })
    );

    const buildAlternateCandidatesFromKeyMatches = (matches, targetNameField, targetKeyField) => (
        (matches || []).map((candidate, idx) => ({
            rank: idx + 1,
            key: targetKeyField ? (candidate?.row?.[targetKeyField] ?? '') : (candidate?.key ?? ''),
            name: targetNameField ? (candidate?.row?.[targetNameField] ?? '') : '',
            similarity: normalizeSimilarityPercent(candidate?.score)
        }))
    );

    const addSelectedColumnValues = (rowData, outcomesRow, wsuRow) => {
        selectedColumns.outcomes.forEach(col => {
            rowData[`outcomes_${col}`] = outcomesRow ? outcomesRow[col] ?? '' : '';
        });
        selectedColumns.wsu_org.forEach(col => {
            rowData[`wsu_${col}`] = wsuRow ? wsuRow[col] ?? '' : '';
        });
    };

    const buildMetadata = (outcomesRow, wsuRow, similarityScore = null) => {
        const outcomesRowIndex = outcomesRowToIndex.has(outcomesRow)
            ? outcomesRowToIndex.get(outcomesRow)
            : '';
        const wsuRowIndex = wsuRowToIndex.has(wsuRow)
            ? wsuRowToIndex.get(wsuRow)
            : '';

        const outcomesDisplayName = outcomesRow
            ? toText(outcomesRow[outcomesDisplayField])
            : '';
        const wsuDisplayName = wsuRow
            ? toText(wsuRow[wsuDisplayField])
            : '';
        const outcomesIdentifier = outcomesRow
            ? (toText(outcomesKeyField ? outcomesRow[outcomesKeyField] : '') || outcomesDisplayName || `outcomes_row_${Number(outcomesRowIndex) + 1}`)
            : '';
        const wsuIdentifier = wsuRow
            ? (toText(wsuKeyField ? wsuRow[wsuKeyField] : '') || wsuDisplayName || (wsuRowIndex === '' ? '' : `wsu_row_${Number(wsuRowIndex) + 1}`))
            : '';

        const similarityPercent = Number.isFinite(similarityScore)
            ? normalizeSimilarityPercent(similarityScore)
            : '';
        let highConfidence = false;
        if (Number.isFinite(similarityScore) && outcomesRow && wsuRow && canNameMatch) {
            const outcomesContext = buildLocationContext(outcomesRow, outcomesLocationFields);
            const wsuContext = buildLocationContext(wsuRow, wsuLocationFields);
            highConfidence = isHighConfidenceNameMatch(
                outcomesRow[outcomesNameField],
                wsuRow[wsuNameField],
                outcomesContext.state,
                wsuContext.state,
                outcomesContext.city,
                wsuContext.city,
                outcomesContext.country,
                wsuContext.country,
                similarityScore,
                threshold
            );
        }

        return {
            outcomes_row_index: outcomesRowIndex,
            wsu_row_index: wsuRowIndex,
            outcomes_record_id: outcomesIdentifier,
            outcomes_display_name: outcomesDisplayName,
            wsu_record_id: wsuIdentifier,
            wsu_display_name: wsuDisplayName,
            proposed_wsu_key: wsuIdentifier,
            proposed_wsu_name: wsuDisplayName,
            match_similarity: similarityPercent,
            confidence_tier: resolveConfidenceTier(similarityPercent, threshold, highConfidence)
        };
    };

    const buildMatchedRow = (outcomesRow, wsuRow, similarityScore = null) => {
        const rowData = {
            ...buildMetadata(outcomesRow, wsuRow, similarityScore),
            alternate_candidates: []
        };
        addSelectedColumnValues(rowData, outcomesRow, wsuRow);
        return rowData;
    };

    const buildErrorRow = (outcomesRow, wsuRow, missingIn, normalizedKey, alternateCandidates = []) => {
        const metadata = buildMetadata(outcomesRow, wsuRow, null);
        const fallbackProposed = alternateCandidates[0] || {};
        const rowData = {
            ...metadata,
            normalized_key: normalizedKey,
            missing_in: missingIn,
            alternate_candidates: alternateCandidates,
            proposed_wsu_key: metadata.proposed_wsu_key || fallbackProposed.key || '',
            proposed_wsu_name: metadata.proposed_wsu_name || fallbackProposed.name || ''
        };
        addSelectedColumnValues(rowData, outcomesRow, wsuRow);
        return rowData;
    };

    if (forceNameMatch || !keyConfig.outcomes || !keyConfig.wsu) {
        if (!canNameMatch) {
            return { cleanRows, errorRows, selectedColumns, headerLabels, generationConfig };
        }

        reportProgress('match_candidates', 0, outcomes.length);
        const { matchesBySource, usedTargets, ambiguousSources, ambiguousCandidatesBySource } = assignGlobalNameMatches(
            outcomes,
            wsuOrg,
            outcomesNameField,
            wsuNameField,
            outcomesLocationFields,
            wsuLocationFields,
            threshold,
            ambiguityGap,
            (processed, total) => reportProgress('match_candidates', processed, total)
        );

        const totalRows = outcomes.length + wsuOrg.length;
        let processedRows = 0;
        outcomes.forEach((outcomesRow, outcomesIndex) => {
            const match = matchesBySource.get(outcomesIndex);
            const wsuRow = match ? wsuOrg[match.targetIndex] : null;
            processedRows += 1;
            reportProgress('build_rows', processedRows, totalRows);

            if (wsuRow) {
                cleanRows.push(buildMatchedRow(outcomesRow, wsuRow, match?.score));
            } else {
                const isAmbiguous = ambiguousSources.has(outcomesIndex);
                const alternateCandidates = isAmbiguous
                    ? buildAlternateCandidatesFromIndexMatches(
                        ambiguousCandidatesBySource.get(outcomesIndex) || [],
                        wsuOrg,
                        wsuNameField,
                        wsuKeyField
                    )
                    : [];
                errorRows.push(
                    buildErrorRow(
                        outcomesRow,
                        null,
                        isAmbiguous ? 'Ambiguous Match' : 'myWSU',
                        outcomesRow[outcomesNameField] ?? '',
                        alternateCandidates
                    )
                );
            }
        });

        wsuOrg.forEach((wsuRow, wsuIndex) => {
            if (usedTargets.has(wsuIndex)) {
                return;
            }
            processedRows += 1;
            reportProgress('build_rows', processedRows, totalRows);
            errorRows.push(
                buildErrorRow(
                    null,
                    wsuRow,
                    'Outcomes',
                    wsuRow[wsuNameField] ?? ''
                )
            );
        });

        return { cleanRows, errorRows, selectedColumns, headerLabels, generationConfig };
    }

    const outcomesMap = buildKeyValueMap(outcomes, keyConfig.outcomes, 'Outcomes source');
    const wsuMap = buildKeyValueMap(wsuOrg, keyConfig.wsu, 'myWSU source');
    const allKeys = new Set([...outcomesMap.keys(), ...wsuMap.keys()]);
    const outcomesEntries = Array.from(outcomesMap.entries()).map(([key, row]) => ({ key, row }));
    const wsuEntries = Array.from(wsuMap.entries()).map(([key, row]) => ({ key, row }));
    const usedOutcomes = new Set();
    const usedWsu = new Set();

    const allKeyList = Array.from(allKeys)
        .sort((a, b) => String(a).localeCompare(String(b)))
    reportProgress('build_rows', 0, allKeyList.length);
    allKeyList.forEach((key, index) => {
            let outcomesRow = outcomesMap.get(key) || null;
            let wsuRow = wsuMap.get(key) || null;
            let handledAsAmbiguous = false;

            if (outcomesRow) usedOutcomes.add(key);
            if (wsuRow) usedWsu.add(key);

            if (!outcomesRow && wsuRow && canNameMatch) {
                const match = findBestNameMatch(
                    wsuRow,
                    wsuNameField,
                    wsuLocationFields,
                    outcomesEntries,
                    outcomesNameField,
                    outcomesLocationFields,
                    threshold,
                    usedOutcomes,
                    ambiguityGap
                );
                if (match?.ambiguous) {
                    const alternateCandidates = buildAlternateCandidatesFromKeyMatches(
                        match.topCandidates,
                        outcomesNameField,
                        outcomesKeyField
                    );
                    errorRows.push(
                        buildErrorRow(
                            null,
                            wsuRow,
                            'Ambiguous Match',
                            wsuRow[wsuNameField] ?? '',
                            alternateCandidates
                        )
                    );
                    handledAsAmbiguous = true;
                } else if (match) {
                    outcomesRow = match.row;
                    usedOutcomes.add(match.key);
                }
            }
            if (outcomesRow && !wsuRow && canNameMatch) {
                const match = findBestNameMatch(
                    outcomesRow,
                    outcomesNameField,
                    outcomesLocationFields,
                    wsuEntries,
                    wsuNameField,
                    wsuLocationFields,
                    threshold,
                    usedWsu,
                    ambiguityGap
                );
                if (match?.ambiguous) {
                    const alternateCandidates = buildAlternateCandidatesFromKeyMatches(
                        match.topCandidates,
                        wsuNameField,
                        wsuKeyField
                    );
                    errorRows.push(
                        buildErrorRow(
                            outcomesRow,
                            null,
                            'Ambiguous Match',
                            outcomesRow[outcomesNameField] ?? '',
                            alternateCandidates
                        )
                    );
                    handledAsAmbiguous = true;
                } else if (match) {
                    wsuRow = match.row;
                    usedWsu.add(match.key);
                }
            }

            if (outcomesRow && wsuRow) {
                let nameScore = null;
                if (canNameMatch) {
                    nameScore = calculateNameSimilarity(
                        outcomesRow[outcomesNameField],
                        wsuRow[wsuNameField]
                    );
                }
                cleanRows.push(buildMatchedRow(outcomesRow, wsuRow, nameScore));
            } else if ((!outcomesRow || !wsuRow) && !handledAsAmbiguous) {
                errorRows.push(
                    buildErrorRow(
                        outcomesRow,
                        wsuRow,
                        outcomesRow ? 'myWSU' : 'Outcomes',
                        key
                    )
                );
            }
            reportProgress('build_rows', index + 1, allKeyList.length);
        });

    return { cleanRows, errorRows, selectedColumns, headerLabels, generationConfig };
}

self.onmessage = (event) => {
    const { type, payload } = event.data || {};
    try {
        if (type === 'validate') {
            const { outcomes, translate, wsu_org, keyConfig, nameCompare } = payload;
            self.postMessage({ type: 'progress', stage: 'merge' });
            const merged = mergeData(outcomes, translate, wsu_org, keyConfig);
            self.postMessage({ type: 'progress', stage: 'validate' });
            const validatedData = validateMappings(
                merged,
                translate,
                outcomes,
                wsu_org,
                keyConfig,
                nameCompare,
                (processed, total) => {
                    self.postMessage({
                        type: 'progress',
                        stage: 'validate',
                        processed,
                        total
                    });
                }
            );
            const missingData = detectMissingMappings(outcomes, translate, keyConfig);
            const stats = generateSummaryStats(validatedData, outcomes, translate, wsu_org);
            self.postMessage({
                type: 'result',
                result: { validatedData, missingData, stats }
            });
            return;
        }
        if (type === 'generate') {
            const result = generateTranslationTableWorker(
                payload.outcomes,
                payload.wsu_org,
                payload.keyConfig,
                payload.nameCompare,
                payload.options,
                payload.selectedColumns,
                payload.keyLabels
            );
            self.postMessage({ type: 'result', result });
        }
    } catch (error) {
        self.postMessage({
            type: 'error',
            message: error?.message || String(error)
        });
    }
};
