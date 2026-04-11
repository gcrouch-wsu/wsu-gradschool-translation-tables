let outcomesData = [];
let translateData = [];
let wsuOrgData = [];

function getSheetHeaderRows(sheet, maxRows = 25) {
    if (!sheet || !sheet['!ref']) return [];
    let range = sheet['!ref'];
    try {
        const decoded = XLSX.utils.decode_range(sheet['!ref']);
        decoded.e.r = Math.min(decoded.e.r, Math.max(0, maxRows - 1));
        range = XLSX.utils.encode_range(decoded);
    } catch (_) {
        range = sheet['!ref'];
    }
    return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false,
        range
    });
}

function scoreSheetForHeaderHints(sheet, normalizedHints) {
    if (!sheet || !normalizedHints.length) return 0;
    const rows = getSheetHeaderRows(sheet, 25);
    if (!rows.length) return 0;
    let best = 0;
    rows.forEach((row) => {
        const cells = (row || [])
            .map(cell => normalizeHeader(cell))
            .filter(Boolean);
        if (!cells.length) return;
        const cellSet = new Set(cells);
        let matched = 0;
        normalizedHints.forEach((hint) => {
            if (cellSet.has(hint)) {
                matched += 1;
            }
        });
        if (matched > best) {
            best = matched;
        }
    });
    return best;
}

function selectBestSheet(workbook, expectedHeaders = [], sheetHeaderHints = []) {
    const sheetNames = workbook && Array.isArray(workbook.SheetNames)
        ? workbook.SheetNames
        : [];
    if (!sheetNames.length) return null;

    const normalizedHints = Array.from(
        new Set(
            [...(expectedHeaders || []), ...(sheetHeaderHints || [])]
                .map(header => normalizeHeader(header))
                .filter(Boolean)
        )
    );
    if (!normalizedHints.length) {
        return workbook.Sheets[sheetNames[0]];
    }

    let bestSheetName = sheetNames[0];
    let bestScore = -1;
    sheetNames.forEach((sheetName) => {
        const score = scoreSheetForHeaderHints(workbook.Sheets[sheetName], normalizedHints);
        if (score > bestScore) {
            bestScore = score;
            bestSheetName = sheetName;
        }
    });

    return workbook.Sheets[bestSheetName];
}

async function loadFile(file, options = {}) {
    return new Promise((resolve, reject) => {
        const fileName = file.name.toLowerCase();
        const expectedHeaders = options.expectedHeaders || [];
        const sheetHeaderHints = options.sheetHeaderHints || [];

        if (fileName.endsWith('.csv')) {
            const userEncoding = options.encoding || 'auto';

            const parseText = (text) => {
                try {
                    return parseCSV(text);
                } catch (parseError) {
                    if (typeof XLSX !== 'undefined') {
                        const workbook = XLSX.read(text, { type: 'string' });
                        const selectedSheet = selectBestSheet(workbook, expectedHeaders, sheetHeaderHints);
                        return sheetToJsonWithHeaderDetection(selectedSheet, expectedHeaders);
                    }
                    throw parseError;
                }
            };

            const attemptRead = (encoding, allowFallback) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const text = e.target.result;
                        const data = parseText(text);
                        resolve(data);
                    } catch (error) {
                        if (allowFallback) {
                            attemptRead('iso-8859-1', false);
                            return;
                        }
                        reject(new Error(`Error parsing CSV: ${error.message}`));
                    }
                };
                reader.onerror = () => reject(new Error('Error reading CSV file'));
                if (encoding) {
                    reader.readAsText(file, encoding);
                } else {
                    reader.readAsText(file);
                }
            };

            if (userEncoding !== 'auto') {
                attemptRead(userEncoding, false);
            } else {
                attemptRead(null, true);
            }
        } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const selectedSheet = selectBestSheet(workbook, expectedHeaders, sheetHeaderHints);
                    const jsonData = sheetToJsonWithHeaderDetection(selectedSheet, expectedHeaders);
                    resolve(jsonData);
                } catch (error) {
                    reject(new Error(`Error parsing Excel: ${error.message}`));
                }
            };
            reader.onerror = () => reject(new Error('Error reading Excel file'));
            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('Unsupported file format'));
        }
    });
}

function normalizeHeader(value) {
    return String(value || '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
}

function sheetToJsonWithHeaderDetection(sheet, expectedHeaders) {
    let jsonData = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: '' });
    if (expectedHeaders.length && hasExpectedHeaders(jsonData, expectedHeaders)) {
        return jsonData;
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    if (!rows.length) {
        return jsonData;
    }

    const expectedNormalized = expectedHeaders.map(normalizeHeader);
    const headerRowIndex = expectedHeaders.length
        ? rows.findIndex(row => {
            const normalizedRow = row.map(normalizeHeader);
            return expectedNormalized.every(header => normalizedRow.includes(header));
        })
        : detectHeaderRowIndex(rows);

    if (headerRowIndex === -1 || headerRowIndex === rows.length - 1) {
        throw new Error('Header row not found.');
    }

    const headers = rows[headerRowIndex];
    const dataRows = rows.slice(headerRowIndex + 1);
    return rowsToObjects(headers, dataRows);
}

function hasExpectedHeaders(jsonData, expectedHeaders) {
    if (!jsonData.length) {
        return false;
    }
    const keys = Object.keys(jsonData[0]).map(normalizeHeader);
    const expectedNormalized = expectedHeaders.map(normalizeHeader);
    return expectedNormalized.every(header => keys.includes(header));
}

function rowsToObjects(headers, dataRows) {
    const cleanedHeaders = headers.map(header => String(header || '').trim());
    return dataRows.map(row => {
        const rowData = {};
        cleanedHeaders.forEach((header, index) => {
            if (header) {
                rowData[header] = row[index] !== undefined ? row[index] : '';
            }
        });
        return rowData;
    });
}

function detectHeaderRowIndex(rows) {
    const maxScan = Math.min(rows.length, 10);
    const scored = [];

    for (let i = 0; i < maxScan; i++) {
        const row = rows[i] || [];
        const cells = row.map(cell => String(cell || '').trim());
        const nonEmpty = cells.filter(cell => cell.length > 0);
        if (nonEmpty.length === 0) {
            continue;
        }
        const alphaCount = nonEmpty.filter(cell => /[A-Za-z]/.test(cell)).length;
        const numericCount = nonEmpty.filter(cell => /^[0-9]+$/.test(cell)).length;
        const uniqueCount = new Set(nonEmpty.map(normalizeHeader)).size;

        scored.push({
            index: i,
            nonEmpty: nonEmpty.length,
            alphaCount,
            numericCount,
            uniqueCount
        });
    }

    if (!scored.length) {
        return -1;
    }

    scored.sort((a, b) => {
        if (b.nonEmpty !== a.nonEmpty) return b.nonEmpty - a.nonEmpty;
        if (b.alphaCount !== a.alphaCount) return b.alphaCount - a.alphaCount;
        if (b.uniqueCount !== a.uniqueCount) return b.uniqueCount - a.uniqueCount;
        return a.index - b.index;
    });

    const best = scored[0];
    const first = scored.find(row => row.index === 0);
    if (first && first.nonEmpty >= best.nonEmpty && first.alphaCount >= 2) {
        return 0;
    }

    if (best.nonEmpty < 2 || best.alphaCount === 0 || best.numericCount === best.nonEmpty) {
        return -1;
    }

    return best.index;
}

function parseCSV(text) {
    const rows = parseCSVRows(text);
    if (rows.length === 0) return [];

    const headers = rows[0].map(h => String(h || '').trim());
    if (headers.length && headers[0]) {
        headers[0] = headers[0].replace(/^\uFEFF/, '');
    }
    const headerCells = headers.filter(cell => cell.length > 0);
    const headerHasAlpha = headerCells.some(cell => /[A-Za-z]/.test(cell));
    if (headerCells.length === 0 || !headerHasAlpha) {
        throw new Error('Header row not found.');
    }

    const data = [];
    for (let i = 1; i < rows.length; i++) {
        const values = rows[i];
        if (!values.length) continue;
        const row = {};
        headers.forEach((header, idx) => {
            if (header) {
                row[header] = values[idx] !== undefined ? values[idx] : '';
            }
        });
        data.push(row);
    }

    return data;
}

function parseCSVRows(text) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];

        if (char === '"') {
            const nextChar = text[i + 1];
            if (inQuotes && nextChar === '"') {
                field += '"';
                i += 1;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }

        if (char === ',' && !inQuotes) {
            row.push(field);
            field = '';
            continue;
        }

        if ((char === '\n' || char === '\r') && !inQuotes) {
            if (char === '\r' && text[i + 1] === '\n') {
                i += 1;
            }
            row.push(field);
            const hasData = row.some(cell => String(cell || '').trim().length > 0);
            if (hasData) {
                rows.push(row.map(cell => String(cell || '').trim()));
            }
            row = [];
            field = '';
            continue;
        }

        field += char;
    }

    if (field.length || row.length) {
        row.push(field);
        const hasData = row.some(cell => String(cell || '').trim().length > 0);
        if (hasData) {
            rows.push(row.map(cell => String(cell || '').trim()));
        }
    }

    return rows;
}

const NAME_TOKEN_STOPWORDS = new Set([
    'university', 'college', 'school', 'academy', 'institute', 'institution',
    'faculty', 'department', 'center', 'centre', 'program', 'programs',
    'campus', 'of', 'the', 'and', 'for', 'at', 'in', 'on', 'de', 'la', 'le',
    'du', 'da', 'di', 'del'
]);
const NAME_MATCH_AMBIGUITY_GAP = 0.03;
const OUTPUT_NOT_FOUND_SUBTYPE = {
    LIKELY_STALE_KEY: 'Output_Not_Found_Likely_Stale_Key',
    AMBIGUOUS_REPLACEMENT: 'Output_Not_Found_Ambiguous_Replacement',
    NO_REPLACEMENT: 'Output_Not_Found_No_Replacement'
};
const STALE_KEY_REPLACEMENT_GAP = 0.04;

const NAME_TOKEN_ALIASES = {
    univ: 'university',
    uni: 'university',
    universiti: 'university',
    universitat: 'university',
    universite: 'university',
    universita: 'university',
    universidad: 'university',
    universidade: 'university',
    coll: 'college',
    col: 'college',
    intl: 'international',
    int: 'international',
    intnl: 'international',
    tech: 'technology',
    sci: 'science',
    engr: 'engineering',
    engg: 'engineering',
    stds: 'studies',
    st: 'saint',
    mt: 'mount',
    ag: 'agricultural',
    med: 'medical',
    dept: 'department',
    ctr: 'center',
    sch: 'school',
    cc: 'community',
    comm: 'community',
    cmty: 'community',
    jr: 'junior',
    u: 'university',
    ft: 'fort',
    tech: 'technology',
    inst: 'institute',
    poly: 'polytechnic',
    polytech: 'polytechnic',
    cal: 'california',
    umass: 'massachusetts',
    ucla: 'california',
    usc: 'southern',
    ucsd: 'california',
    ucd: 'california',
    ucsb: 'california',
    uci: 'california',
    ucr: 'california',
    ucm: 'california',
    ucb: 'california',
    asu: 'state',
    fsu: 'state',
    osu: 'state',
    isu: 'state',
    wsu: 'state',
    sjsu: 'state',
    sdsu: 'state',
    sfsu: 'state',
    unc: 'north',
    uab: 'alabama',
    uofl: 'louisville',
    natl: 'national'
};

function normalizeNameForCompare(value) {
    return String(value || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\buniv\b/g, 'university')
        .replace(/\bcal\b/g, 'california')
        .replace(/\buniveristy\b/g, 'university')
        .replace(/\buniversiti\b/g, 'university')
        .replace(/\bcal state university\b/g, 'california state')
        .replace(/\bcalifornia state university\b/g, 'california state')
        .replace(/\bcal state\b/g, 'california state')
        .replace(/\buniversity of california\b/g, 'california')
        .replace(/\buc\b/g, 'california')
        .replace(/\bstate u\b/g, 'state university')
        .replace(/\bstate univ\b/g, 'state university')
        .replace(/\bumass\b/g, 'university of massachusetts')
        .replace(/\btexas a&m\b/g, 'texas agricultural and mechanical')
        .replace(/\ba&m\b/g, 'agricultural and mechanical')
        .replace(/\bgeorgia tech\b/g, 'georgia institute of technology')
        .replace(/\bga tech\b/g, 'georgia institute of technology')
        .replace(/\bjunior college\b/g, 'community college')
        .replace(/\bjr college\b/g, 'community college')
        .replace(/\bcc\b/g, 'community college')
        .replace(/\bcmty\b/g, 'community')
        .replace(/\bcol\b/g, 'college')
        .replace(/\bnatl\b/g, 'national')
        .replace(/\bengr\b/g, 'engineering')
        .replace(/\bengg\b/g, 'engineering')
        .replace(/\bstds\b/g, 'studies')
        .replace(/\bintn'?l\b/g, 'international')
        .replace(/\blowcountry\b/g, 'low country')
        .replace(/\bft\b/g, 'fort')
        .replace(/\bu of\b/g, 'university of')
        .replace(/\bla verne\b/g, 'laverne')
        .replace(/&/g, ' and ')
        .replace(/[^a-z0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function tokenizeNormalizedName(normalizedValue) {
    const normalized = String(normalizedValue || '').trim();
    if (!normalized) {
        return [];
    }
    return normalized
        .split(' ')
        .map(token => NAME_TOKEN_ALIASES[token] || token)
        .filter(Boolean);
}

function tokenizeName(value) {
    return tokenizeNormalizedName(normalizeNameForCompare(value));
}

function buildTokenIDF(allNames) {
    const names = Array.isArray(allNames) ? allNames : [];
    const docCount = names.length;
    if (!docCount) return {};
    const df = {};
    names.forEach(name => {
        const tokens = new Set(getInformativeTokens(tokenizeName(name)));
        tokens.forEach(token => {
            df[token] = (df[token] || 0) + 1;
        });
    });
    const idf = {};
    Object.keys(df).forEach(token => {
        idf[token] = Math.log((docCount + 1) / (df[token] + 1)) + 1;
    });
    const idfValues = Object.values(idf).sort((a, b) => a - b);
    const medianIDF = idfValues.length
        ? idfValues[Math.floor(idfValues.length / 2)]
        : 0;
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
}

function getMedianIDF(idfTable) {
    if (!idfTable || typeof idfTable !== 'object') {
        return 0;
    }
    if (typeof idfTable.__median === 'number' && Number.isFinite(idfTable.__median)) {
        return idfTable.__median;
    }
    const idfValues = Object.keys(idfTable)
        .filter(key => !key.startsWith('__'))
        .map(key => idfTable[key])
        .filter(value => typeof value === 'number' && Number.isFinite(value))
        .sort((a, b) => a - b);
    const medianIDF = idfValues.length
        ? idfValues[Math.floor(idfValues.length / 2)]
        : 0;
    try {
        Object.defineProperty(idfTable, '__median', {
            value: medianIDF,
            writable: true,
            enumerable: false,
            configurable: true
        });
    } catch (error) {
        idfTable.__median = medianIDF;
    }
    return medianIDF;
}

function getInformativeTokens(tokens) {
    return tokens.filter(token => token.length > 1 && !NAME_TOKEN_STOPWORDS.has(token));
}

function jaroWinkler(value1, value2) {
    const s1 = String(value1 || '');
    const s2 = String(value2 || '');
    if (s1 === s2) return 1.0;
    const len1 = s1.length;
    const len2 = s2.length;
    if (!len1 || !len2) return 0.0;

    const matchWindow = Math.max(0, Math.floor(Math.max(len1, len2) / 2) - 1);
    const s1Matches = new Array(len1).fill(false);
    const s2Matches = new Array(len2).fill(false);
    let matches = 0;

    for (let i = 0; i < len1; i++) {
        const start = Math.max(0, i - matchWindow);
        const end = Math.min(i + matchWindow + 1, len2);
        for (let j = start; j < end; j++) {
            if (s2Matches[j] || s1[i] !== s2[j]) continue;
            s1Matches[i] = true;
            s2Matches[j] = true;
            matches += 1;
            break;
        }
    }

    if (!matches) return 0.0;

    let transpositions = 0;
    let k = 0;
    for (let i = 0; i < len1; i++) {
        if (!s1Matches[i]) continue;
        while (!s2Matches[k]) k += 1;
        if (s1[i] !== s2[k]) transpositions += 1;
        k += 1;
    }

    const jaro = (matches / len1 + matches / len2 + (matches - transpositions / 2) / matches) / 3;
    let prefixLen = 0;
    const maxPrefix = Math.min(4, len1, len2);
    for (let i = 0; i < maxPrefix; i++) {
        if (s1[i] !== s2[i]) break;
        prefixLen += 1;
    }
    return jaro + prefixLen * 0.1 * (1 - jaro);
}

function tokenSimilarity(token1, token2) {
    if (token1 === token2) return 1.0;
    return jaroWinkler(token1, token2);
}

function tokensMatch(token1, token2) {
    return tokenSimilarity(token1, token2) >= 0.8;
}

const CONTRADICTORY_NAME_PAIRS = [
    [['dental', 'dentistry'], ['medical', 'medicine']],
    [['dental'], ['veterinary', 'vet']],
    [['law'], ['medical', 'business', 'engineering']],
    [['medical'], ['law', 'business', 'engineering']],
    [['graduate'], ['undergraduate']],
    [['community college', 'cc'], ['university']]
];

function tokenOverlapStats(tokens1, tokens2, idfTable = null) {
    if (!tokens1.length || !tokens2.length) {
        return { score: 0, matches: 0, forwardBest: [], backwardBest: [] };
    }

    const mongeElkanScore = (queryTokens, targetTokens, idf = null) => {
        let weightedSum = 0;
        let weightSum = 0;
        let softMatches = 0;
        const bestByQuery = [];
        for (const queryToken of queryTokens) {
            let bestSimilarity = 0;
            for (const targetToken of targetTokens) {
                const similarity = tokenSimilarity(queryToken, targetToken);
                if (similarity > bestSimilarity) {
                    bestSimilarity = similarity;
                }
            }
            const weight = (idf && idf[queryToken]) ? idf[queryToken] : 1;
            weightedSum += bestSimilarity * weight;
            weightSum += weight;
            if (bestSimilarity >= 0.8) {
                softMatches += 1;
            }
            bestByQuery.push({ token: queryToken, bestSimilarity });
        }
        return {
            score: weightSum > 0 ? weightedSum / weightSum : 0,
            matches: softMatches,
            bestByQuery
        };
    };

    const forward = mongeElkanScore(tokens1, tokens2, idfTable);
    const backward = mongeElkanScore(tokens2, tokens1, idfTable);

    return {
        score: Math.max(forward.score, backward.score),
        matches: Math.max(forward.matches, backward.matches),
        forwardBest: forward.bestByQuery,
        backwardBest: backward.bestByQuery
    };
}

const STATE_PROVINCE_MAP = {
    al: 'alabama', ak: 'alaska', az: 'arizona', ar: 'arkansas', ca: 'california',
    co: 'colorado', ct: 'connecticut', de: 'delaware', fl: 'florida', ga: 'georgia',
    hi: 'hawaii', id: 'idaho', il: 'illinois', in: 'indiana', ia: 'iowa',
    ks: 'kansas', ky: 'kentucky', la: 'louisiana', ma: 'massachusetts', md: 'maryland',
    me: 'maine', mi: 'michigan', mn: 'minnesota', mo: 'missouri', ms: 'mississippi',
    mt: 'montana', nc: 'north carolina', nd: 'north dakota', ne: 'nebraska', nh: 'new hampshire',
    nj: 'new jersey', nm: 'new mexico', nv: 'nevada', ny: 'new york', oh: 'ohio',
    ok: 'oklahoma', or: 'oregon', pa: 'pennsylvania', ri: 'rhode island', sc: 'south carolina',
    sd: 'south dakota', tn: 'tennessee', tx: 'texas', ut: 'utah', va: 'virginia',
    vt: 'vermont', wa: 'washington', wi: 'wisconsin', wv: 'west virginia', wy: 'wyoming',
    dc: 'district of columbia',
    ab: 'alberta', bc: 'british columbia', mb: 'manitoba', nb: 'new brunswick',
    nl: 'newfoundland and labrador', ns: 'nova scotia', nt: 'northwest territories',
    nu: 'nunavut', on: 'ontario', pe: 'prince edward island', qc: 'quebec',
    sk: 'saskatchewan', yt: 'yukon'
};

const COUNTRY_CODE_MAP = {
    us: 'usa', usa: 'usa', 'united states': 'usa', 'united states of america': 'usa',
    ca: 'can', can: 'can', canada: 'can',
    uk: 'gbr', gb: 'gbr', gbr: 'gbr', 'united kingdom': 'gbr', 'great britain': 'gbr',
    in: 'ind', ind: 'ind', india: 'ind',
    bd: 'bgd', bgd: 'bgd', bangladesh: 'bgd',
    pk: 'pak', pak: 'pak', pakistan: 'pak',
    my: 'mys', mys: 'mys', malaysia: 'mys',
    ng: 'nga', nga: 'nga', nigeria: 'nga',
    th: 'tha', tha: 'tha', thailand: 'tha',
    np: 'npl', npl: 'npl', nepal: 'npl',
    kr: 'kor', kor: 'kor', korea: 'kor', 'south korea': 'kor'
};

const NAME_JOINER_STOPWORDS = new Set([
    'of', 'the', 'and', 'for', 'at', 'in', 'on', 'de', 'la', 'le', 'du', 'da', 'di', 'del'
]);

function normalizeStateValue(value) {
    return String(value || '').trim().toLowerCase();
}

function normalizeCountryValue(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\./g, '')
        .replace(/\s+/g, ' ');
}

function statesMatch(value1, value2) {
    const state1 = normalizeStateValue(value1);
    const state2 = normalizeStateValue(value2);
    if (!state1 || !state2) return false;
    const normalized1 = state1.replace(/\./g, '');
    const normalized2 = state2.replace(/\./g, '');
    if (normalized1 === normalized2) return true;
    if (['dc', 'washington dc', 'washington d c', 'district of columbia'].includes(normalized1) &&
        ['dc', 'washington dc', 'washington d c', 'district of columbia'].includes(normalized2)
    ) {
        return true;
    }
    const mapped1 = STATE_PROVINCE_MAP[normalized1] || normalized1;
    const mapped2 = STATE_PROVINCE_MAP[normalized2] || normalized2;
    return mapped1 === mapped2;
}

function countriesMatch(value1, value2) {
    const country1 = normalizeCountryValue(value1);
    const country2 = normalizeCountryValue(value2);
    if (!country1 || !country2) return false;
    const mapped1 = COUNTRY_CODE_MAP[country1] || country1;
    const mapped2 = COUNTRY_CODE_MAP[country2] || country2;
    return mapped1 === mapped2;
}

function extractParentheticalChunks(nameValue) {
    if (!nameValue) return [];
    return Array.from(String(nameValue).matchAll(/\(([^)]+)\)/g))
        .map(match => String(match[1] || '').trim())
        .filter(Boolean);
}

function extractParentheticalTokens(nameValue) {
    const chunks = extractParentheticalChunks(nameValue);
    if (!chunks.length) return [];
    return chunks
        .flatMap(chunk => normalizeNameForCompare(chunk).split(' '))
        .filter(Boolean);
}

function extractHyphenSuffixTokens(nameValue) {
    if (!nameValue) return [];
    const normalizedName = String(nameValue).replace(/[–—]/g, '-');
    const safeParts = normalizedName.split('-');
    if (safeParts.length >= 2) {
        const safeSuffix = safeParts[safeParts.length - 1].trim();
        if (safeSuffix) {
            return normalizeNameForCompare(safeSuffix).split(' ').filter(Boolean);
        }
    }
    const parts = String(nameValue).split(/[-–—]/);
    if (parts.length < 2) return [];
    const suffix = parts[parts.length - 1].trim();
    if (!suffix) return [];
    return normalizeNameForCompare(suffix).split(' ').filter(Boolean);
}

function extractAliasParentheticalCandidates(nameValue) {
    const chunks = extractParentheticalChunks(nameValue);
    if (!chunks.length) return [];
    return chunks
        .filter(chunk => /\b(now|formerly|aka|previously)\b/i.test(chunk))
        .map(chunk => chunk.replace(/\b(now|formerly|aka|previously)\b/gi, ' '))
        .map(chunk => normalizeNameForCompare(chunk))
        .map(chunk => chunk.replace(/\bclosed\s+[0-9]{4}\b/g, '').trim())
        .filter(Boolean);
}

function aliasParentheticalMatches(nameWithAlias, otherName) {
    if (!nameWithAlias || !otherName) return false;
    const aliasCandidates = extractAliasParentheticalCandidates(nameWithAlias);
    if (!aliasCandidates.length) return false;

    const normalizedOther = normalizeNameForCompare(otherName);
    const otherTokens = getInformativeTokens(tokenizeName(otherName));
    if (!normalizedOther || !otherTokens.length) return false;

    for (const candidate of aliasCandidates) {
        if (candidate === normalizedOther) {
            return true;
        }
        const aliasTokens = getInformativeTokens(tokenizeName(candidate));
        if (!aliasTokens.length) {
            continue;
        }
        const overlap = tokenOverlapStats(aliasTokens, otherTokens);
        const required = Math.min(2, aliasTokens.length, otherTokens.length);
        if (required > 0 && overlap.matches >= required && overlap.score >= 0.6) {
            return true;
        }
    }
    return false;
}

function parentheticalAcronymMatches(nameWithAcronym, otherName) {
    if (!nameWithAcronym || !otherName) return false;
    const chunks = extractParentheticalChunks(nameWithAcronym);
    if (!chunks.length) return false;

    const tokens = tokenizeName(otherName).filter(
        token => token.length > 1 && !NAME_JOINER_STOPWORDS.has(token)
    );
    if (tokens.length < 2) return false;
    const initials = tokens.map(token => token[0]).join('');
    if (!initials) return false;

    return chunks.some(chunk => {
        const acronym = normalizeNameForCompare(chunk).replace(/\s+/g, '');
        if (!/^[a-z]{3,6}$/.test(acronym)) return false;
        return initials === acronym || initials.includes(acronym);
    });
}

function normalizeUniversityRemainder(nameValue) {
    const normalized = normalizeNameForCompare(nameValue);
    if (!normalized) return '';
    if (normalized.startsWith('university of ')) {
        return normalized.slice('university of '.length).trim();
    }
    if (normalized.endsWith(' university')) {
        return normalized.slice(0, -' university'.length).trim();
    }
    return '';
}

function universityWordOrderVariant(name1, name2) {
    const remainder1 = normalizeUniversityRemainder(name1);
    const remainder2 = normalizeUniversityRemainder(name2);
    if (!remainder1 || !remainder2) return false;
    if (remainder1.length < 4 || remainder2.length < 4) return false;
    return remainder1 === remainder2;
}

function hasInternationalStateGap(state1, state2) {
    const normalized1 = normalizeStateValue(state1);
    const normalized2 = normalizeStateValue(state2);
    return !normalized1 || !normalized2 || normalized1 === 'ot' || normalized2 === 'ot';
}

function hasComparableStateValues(state1, state2) {
    const normalized1 = normalizeStateValue(state1);
    const normalized2 = normalizeStateValue(state2);
    return Boolean(normalized1 && normalized2 && normalized1 !== 'ot' && normalized2 !== 'ot');
}

function locationInNameMatches(nameValue, cityValue, stateValue) {
    const tokens = [
        ...extractParentheticalTokens(nameValue),
        ...extractHyphenSuffixTokens(nameValue)
    ];
    if (!tokens.length) return false;
    const normalizedCity = normalizeNameForCompare(cityValue);
    for (const token of tokens) {
        if (normalizedCity) {
            if (token === normalizedCity) {
                return true;
            }
            if (token.length >= 3 && (
                normalizedCity.includes(token) || token.includes(normalizedCity)
            )) {
                return true;
            }
        }
        if (stateValue && statesMatch(token, stateValue)) {
            return true;
        }
    }
    return false;
}

function isHighConfidenceNameMatch(
    name1,
    name2,
    state1,
    state2,
    city1,
    city2,
    country1,
    country2,
    similarity,
    threshold,
    idfTable = null
) {
    if (!name1 || !name2) return false;
    const stateOkay = statesMatch(state1, state2);
    const cityOkay = cityInName(name1, city2) || cityInName(name2, city1);
    const parenOkay = locationInNameMatches(name1, city2, state2) || locationInNameMatches(name2, city1, state1);
    const countryOkay = countriesMatch(country1, country2);
    const aliasParenOkay = aliasParentheticalMatches(name1, name2) || aliasParentheticalMatches(name2, name1);
    const acronymOkay = parentheticalAcronymMatches(name1, name2) || parentheticalAcronymMatches(name2, name1);
    const universityVariantOkay = universityWordOrderVariant(name1, name2);

    const tokens1 = getInformativeTokens(tokenizeName(name1));
    const tokens2 = getInformativeTokens(tokenizeName(name2));
    const overlap = tokenOverlapStats(tokens1, tokens2, idfTable);

    if ((aliasParenOkay || acronymOkay || universityVariantOkay) && (countryOkay || stateOkay || cityOkay || parenOkay)) {
        return true;
    }

    if (!stateOkay && !cityOkay && !parenOkay && !countryOkay) return false;

    if (typeof similarity === 'number' && similarity >= Math.max(0, threshold - 0.05)) {
        return true;
    }

    if (parenOkay || stateOkay) {
        return overlap.matches >= 1;
    }
    if (cityOkay) {
        return overlap.matches >= 2;
    }
    if (countryOkay && hasInternationalStateGap(state1, state2)) {
        return overlap.matches >= 2 && overlap.score >= 0.6;
    }
    return overlap.matches >= 1;
}

function fieldEvidence(
    name1,
    name2,
    state1,
    state2,
    city1,
    city2,
    country1,
    country2,
    code1 = '',
    code2 = ''
) {
    let logOdds = 0;

    if (countriesMatch(country1, country2)) {
        logOdds += 1.5;
    } else if (country1 && country2) {
        logOdds -= 2.0;
    }

    if (statesMatch(state1, state2)) {
        logOdds += 2.0;
    } else if (hasComparableStateValues(state1, state2)) {
        logOdds -= 1.0;
    }

    if (cityInName(name1, city2) || cityInName(name2, city1)) {
        logOdds += 1.5;
    }

    const normalizedCode1 = normalizeKeyValue(code1);
    const normalizedCode2 = normalizeKeyValue(code2);
    if (normalizedCode1 && normalizedCode2 && normalizedCode1 === normalizedCode2) {
        logOdds += 5.0;
    }

    return 1 / (1 + Math.exp(-logOdds));
}

function cityInName(nameValue, cityValue) {
    if (!nameValue || !cityValue) return false;
    const normalizedName = normalizeNameForCompare(nameValue);
    const normalizedCity = normalizeNameForCompare(cityValue);
    if (!normalizedName || !normalizedCity) return false;
    if (normalizedName.includes(normalizedCity)) return true;
    const tokens = normalizedCity.split(' ').filter(Boolean);
    if (!tokens.length) return false;
    return tokens.every(token => normalizedName.includes(token));
}

function getTopTwoNameScores(sourceName, targetRows, targetField, threshold = 0, scoreCache = null, idfTable = null) {
    if (!sourceName) {
        return { bestScore: -1, secondBestScore: -1 };
    }
    const normalizedSource = normalizeNameForCompare(sourceName);
    if (!normalizedSource) {
        return { bestScore: -1, secondBestScore: -1 };
    }
    const cacheKey = scoreCache
        ? `${targetField}|${threshold}|${normalizedSource}|${Boolean(idfTable)}`
        : '';
    if (scoreCache && scoreCache.has(cacheKey)) {
        return scoreCache.get(cacheKey);
    }

    let bestScore = -1;
    let secondBestScore = -1;
    targetRows.forEach(row => {
        const targetName = row[targetField];
        if (!targetName) {
            return;
        }
        const score = calculateNameSimilarity(sourceName, targetName, idfTable);
        if (score < threshold) {
            return;
        }
        if (score > bestScore) {
            secondBestScore = bestScore;
            bestScore = score;
        } else if (score > secondBestScore) {
            secondBestScore = score;
        }
    });
    const result = { bestScore, secondBestScore };
    if (scoreCache) {
        scoreCache.set(cacheKey, result);
    }
    return result;
}

function isAmbiguousNameMatch(
    sourceName,
    targetName,
    targetRows,
    targetField,
    threshold,
    gap = NAME_MATCH_AMBIGUITY_GAP,
    scoreCache = null,
    mappedScoreValue = null,
    idfTable = null
) {
    if (!sourceName || !targetName) {
        return false;
    }
    const mappedScore = typeof mappedScoreValue === 'number'
        ? mappedScoreValue
        : calculateNameSimilarity(sourceName, targetName, idfTable);
    if (mappedScore < threshold) {
        return false;
    }
    const { bestScore, secondBestScore } = getTopTwoNameScores(
        sourceName,
        targetRows,
        targetField,
        threshold,
        scoreCache,
        idfTable
    );
    if (bestScore < 0 || secondBestScore < 0) {
        return false;
    }
    if (bestScore - secondBestScore >= gap) {
        return false;
    }
    return (bestScore - mappedScore) <= gap;
}

function classifyMissingOutputReplacement(
    sourceName,
    sourceState,
    sourceCity,
    sourceCountry,
    wsuRows,
    options = {}
) {
    const nameField = options.nameField || '';
    const keyField = options.keyField || '';
    const stateField = options.stateField || '';
    const cityField = options.cityField || '';
    const countryField = options.countryField || '';
    const threshold = typeof options.threshold === 'number' ? options.threshold : 0.8;
    const gap = typeof options.ambiguityGap === 'number'
        ? Math.max(0.01, options.ambiguityGap)
        : STALE_KEY_REPLACEMENT_GAP;
    const idfTable = options.idfTable || null;

    if (!sourceName || !nameField || !keyField || !Array.isArray(wsuRows) || !wsuRows.length) {
        return { subtype: OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT, bestCandidate: null };
    }

    const candidates = [];
    wsuRows.forEach(targetRow => {
        const targetName = targetRow[nameField];
        if (!targetName) {
            return;
        }
        const targetState = stateField ? (targetRow[stateField] ?? '') : '';
        const targetCity = cityField ? (targetRow[cityField] ?? '') : '';
        const targetCountry = countryField ? (targetRow[countryField] ?? '') : '';

        if (sourceCountry && targetCountry && !countriesMatch(sourceCountry, targetCountry)) {
            return;
        }
        if (!sourceCountry && targetCountry) {
            return;
        }

        if (
            hasComparableStateValues(sourceState, targetState) &&
            sourceCountry &&
            targetCountry &&
            countriesMatch(sourceCountry, targetCountry) &&
            !statesMatch(sourceState, targetState)
        ) {
            return;
        }

        const similarity = calculateNameSimilarity(sourceName, targetName, idfTable);
        if (!isHighConfidenceNameMatch(
            sourceName,
            targetName,
            sourceState,
            targetState,
            sourceCity,
            targetCity,
            sourceCountry,
            targetCountry,
            similarity,
            threshold,
            idfTable
        )) {
            return;
        }

        candidates.push({
            key: targetRow[keyField] ?? '',
            name: targetName,
            city: targetCity,
            state: targetState,
            country: targetCountry,
            score: similarity
        });
    });

    if (!candidates.length) {
        return { subtype: OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT, bestCandidate: null };
    }

    candidates.sort((a, b) => b.score - a.score);
    const bestCandidate = candidates[0];
    const secondCandidate = candidates.length > 1 ? candidates[1] : null;

    if (secondCandidate && (bestCandidate.score - secondCandidate.score) < gap) {
        return {
            subtype: OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT,
            bestCandidate,
            secondCandidate
        };
    }

    return {
        subtype: OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY,
        bestCandidate,
        secondCandidate
    };
}

function calculateNameSimilarity(name1, name2, idfTable = null) {
    if (!name1 || !name2) return 0.0;

    const normalized1 = normalizeNameForCompare(name1);
    const normalized2 = normalizeNameForCompare(name2);

    if (!normalized1 || !normalized2) {
        return 0.0;
    }
    if (normalized1 === normalized2) {
        return 1.0;
    }

    const tokensAll1 = tokenizeNormalizedName(normalized1);
    const tokensAll2 = tokenizeNormalizedName(normalized2);
    const informative1 = getInformativeTokens(tokensAll1);
    const informative2 = getInformativeTokens(tokensAll2);
    if (!informative1.length && !informative2.length) {
        return 0.0;
    }
    const overlap = tokenOverlapStats(informative1, informative2, idfTable);
    const tokenSimilarityScore = overlap.score;
    const baseSimilarity = similarityRatio(normalized1, normalized2);

    const hasWord = (word, normalized, tokens) => {
        if (!word) return false;
        if (word.includes(' ')) {
            return normalized.includes(word);
        }
        if (word.length <= 3) {
            return tokens.includes(word);
        }
        return normalized.includes(word);
    };

    for (const [group1, group2] of CONTRADICTORY_NAME_PAIRS) {
        const hasGroup1 = group1.some(word => hasWord(word, normalized1, tokensAll1));
        const hasGroup2 = group2.some(word => hasWord(word, normalized2, tokensAll2));

        if (hasGroup1 && hasGroup2) return 0.0;

        const hasGroup1In2 = group1.some(word => hasWord(word, normalized2, tokensAll2));
        const hasGroup2In1 = group2.some(word => hasWord(word, normalized1, tokensAll1));

        if (hasGroup1In2 && hasGroup2In1) return 0.0;
    }

    if (informative1.length >= 2 && informative2.length >= 2 && overlap.matches === 0) {
        return 0.0;
    }
    if (informative1.length && informative2.length && tokenSimilarityScore === 0) {
        return Math.min(baseSimilarity * 0.6, 0.45);
    }

    let score = (baseSimilarity * 0.55) + (tokenSimilarityScore * 0.45);
    if (tokenSimilarityScore < 0.2 && baseSimilarity < 0.85) {
        score *= 0.85;
    }
    const medianIDF = idfTable ? getMedianIDF(idfTable) : 0;
    if (idfTable && informative1.length && informative2.length) {
        const countRareMismatches = (bestByQuery) => {
            let mismatches = 0;
            bestByQuery.forEach(({ token, bestSimilarity }) => {
                if (bestSimilarity < 0.8 && (idfTable[token] || 0) >= medianIDF) {
                    mismatches += 1;
                }
            });
            return mismatches;
        };
        const rareMismatches = Math.max(
            countRareMismatches(overlap.forwardBest || []),
            countRareMismatches(overlap.backwardBest || [])
        );
        if (rareMismatches > 0) {
            score -= Math.min(0.2, rareMismatches * 0.06);
        }
    }

    const TRUNCATION_LENGTH = 30;
    const originalLen1 = String(name1 || '').trim().length;
    const originalLen2 = String(name2 || '').trim().length;
    if ((originalLen1 === TRUNCATION_LENGTH || originalLen2 === TRUNCATION_LENGTH) && idfTable) {
        const useForward = originalLen1 === TRUNCATION_LENGTH;
        const truncTokens = useForward ? informative1 : informative2;
        const bestMatches = useForward ? (overlap.forwardBest || []) : (overlap.backwardBest || []);
        let containedCount = 0;
        let hasRareAlignment = false;
        for (const truncToken of truncTokens) {
            const tokenMatch = bestMatches.find(entry => entry.token === truncToken);
            const bestSimilarity = tokenMatch ? tokenMatch.bestSimilarity : 0;
            if (bestSimilarity >= 0.8) {
                containedCount += 1;
                if ((idfTable[truncToken] || 0) > medianIDF) {
                    hasRareAlignment = true;
                }
            }
        }
        const containment = truncTokens.length > 0 ? containedCount / truncTokens.length : 0;
        if (containment >= 0.9 && hasRareAlignment) {
            score = Math.min(1.0, score + 0.08);
        }
    }

    return Math.max(0, Math.min(1, score));
}

function similarityRatio(str1, str2) {
    const matrix = [];
    const len1 = str1.length;
    const len2 = str2.length;

    for (let i = 0; i <= len1; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
            if (str1[i - 1] === str2[j - 1]) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }

    const distance = matrix[len1][len2];
    const maxLen = Math.max(len1, len2);
    return maxLen === 0 ? 1.0 : (maxLen - distance) / maxLen;
}

function normalizeKeyValue(value) {
    if (value === null || value === undefined) {
        return '';
    }
    const raw = String(value).trim();
    if (raw === '') {
        return '';
    }
    if (/^\d+$/.test(raw)) {
        const cleaned = raw.replace(/^0+/, '');
        return cleaned === '' ? '0' : cleaned;
    }
    return raw.toLowerCase();
}

function buildUniqueKeyMap(rows, keyField, datasetLabel) {
    const map = new Map();
    const duplicateCounts = new Map();

    rows.forEach(row => {
        const normalized = normalizeKeyValue(row[keyField]);
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

function mergeData(outcomes, translate, wsuOrg, keyConfig) {
    const merged = [];
    const outcomesMap = buildUniqueKeyMap(outcomes, keyConfig.outcomes, 'Outcomes source');
    const wsuOrgMap = buildUniqueKeyMap(wsuOrg, keyConfig.wsu, 'myWSU source');

    for (const tRow of translate) {
        const inputRaw = tRow[keyConfig.translateInput] ?? '';
        const outputRaw = tRow[keyConfig.translateOutput] ?? '';
        const inputNormalized = normalizeKeyValue(inputRaw);
        const outputNormalized = normalizeKeyValue(outputRaw);

        const outcomesMatch = outcomesMap.get(inputNormalized);
        const wsuMatch = wsuOrgMap.get(outputNormalized);

        const mergedRow = {
            translate_input: inputRaw,
            translate_output: outputRaw,
            translate_input_norm: inputNormalized,
            translate_output_norm: outputNormalized
        };

        if (outcomesMatch) {
            for (const key in outcomesMatch) {
                mergedRow[`outcomes_${key}`] = outcomesMatch[key];
            }
        }

        if (wsuMatch) {
            for (const key in wsuMatch) {
                mergedRow[`wsu_${key}`] = wsuMatch[key];
            }
        }

        merged.push(mergedRow);
    }

    console.log(`[OK] Merged data: ${merged.length} rows`);
    return merged;
}

function detectDuplicateTargets(translate, keyConfig) {
    const targetMap = {};

    for (const row of translate) {
        const targetKey = normalizeKeyValue(row[keyConfig.translateOutput]);
        const sourceKey = normalizeKeyValue(row[keyConfig.translateInput]);
        if (!targetKey || !sourceKey) {
            continue;
        }
        if (!targetMap[targetKey]) {
            targetMap[targetKey] = new Set();
        }
        targetMap[targetKey].add(sourceKey);
    }

    const duplicates = {};
    for (const [target, sourceSet] of Object.entries(targetMap)) {
        const sources = Array.from(sourceSet);
        if (sources.length > 1) {
            duplicates[target] = sources;
        }
    }

    const totalDuplicateRows = Object.values(duplicates).reduce((sum, codes) => sum + codes.length, 0);
    console.log(`[OK] Found ${Object.keys(duplicates).length} target keys with duplicates (${totalDuplicateRows} total rows)`);

    return duplicates;
}

function detectDuplicateSources(translate, keyConfig) {
    const sourceMap = {};

    for (const row of translate) {
        const sourceKey = normalizeKeyValue(row[keyConfig.translateInput]);
        const targetKey = normalizeKeyValue(row[keyConfig.translateOutput]);
        if (!sourceKey) {
            continue;
        }
        if (!sourceMap[sourceKey]) {
            sourceMap[sourceKey] = new Set();
        }
        if (targetKey) {
            sourceMap[sourceKey].add(targetKey);
        }
    }

    const duplicates = {};
    for (const [source, targetSet] of Object.entries(sourceMap)) {
        const targets = Array.from(targetSet);
        if (targets.length > 1) {
            duplicates[source] = targets;
        }
    }

    console.log(`[OK] Found ${Object.keys(duplicates).length} source keys with duplicate mappings`);
    return duplicates;
}

function detectOrphanedMappings(translate, outcomes, keyConfig) {
    const outcomesKeys = new Set(outcomes.map(o => normalizeKeyValue(o[keyConfig.outcomes])));
    const orphaned = [];

    for (const row of translate) {
        const sourceKey = normalizeKeyValue(row[keyConfig.translateInput]);
        if (sourceKey && !outcomesKeys.has(sourceKey)) {
            orphaned.push(sourceKey);
        }
    }

    console.log(`[OK] Found ${orphaned.length} orphaned mappings`);
    return orphaned;
}

function detectMissingMappings(outcomes, translate, keyConfig) {
    const translateKeys = new Set(translate.map(t => normalizeKeyValue(t[keyConfig.translateInput])));
    const missing = outcomes.filter(o => {
        const key = normalizeKeyValue(o[keyConfig.outcomes]);
        return key && !translateKeys.has(key);
    });

    console.log(`[OK] Found ${missing.length} missing mappings`);
    return missing;
}

function validateMappings(merged, translate, outcomes, wsuOrg, keyConfig, nameCompare = {}, onProgress = null) {
    console.log('\n=== Running Validation ===');

    const duplicateTargetsDict = detectDuplicateTargets(translate, keyConfig);
    const duplicateSourcesDict = detectDuplicateSources(translate, keyConfig);

    const outcomesKeys = new Set(
        outcomes
            .map(row => normalizeKeyValue(row[keyConfig.outcomes]))
            .filter(Boolean)
    );
    const wsuKeys = new Set(
        wsuOrg
            .map(row => normalizeKeyValue(row[keyConfig.wsu]))
            .filter(Boolean)
    );

    const duplicateGroups = {};
    const sortedTargetDups = Object.entries(duplicateTargetsDict)
        .sort((a, b) => b[1].length - a[1].length);

    sortedTargetDups.forEach(([targetKey, sourceCodes], idx) => {
        const groupName = idx < 26 ? `Tgt_Group_${String.fromCharCode(65 + idx)}` : `Tgt_Group_${idx + 1}`;
        duplicateGroups[`tgt:${targetKey}`] = groupName;
    });

    const sortedSourceDups = Object.entries(duplicateSourcesDict)
        .sort((a, b) => b[1].length - a[1].length);

    sortedSourceDups.forEach(([sourceKey, targetCodes], idx) => {
        const groupName = idx < 26 ? `Src_Group_${String.fromCharCode(65 + idx)}` : `Src_Group_${idx + 1}`;
        duplicateGroups[`src:${sourceKey}`] = groupName;
    });

    const nameCompareEnabled = Boolean(nameCompare.enabled);
    const outcomesColumn = nameCompare.outcomes_column || '';
    const wsuColumn = nameCompare.wsu_column || '';
    const threshold = typeof nameCompare.threshold === 'number' ? nameCompare.threshold : 0.8;
    const ambiguityGap = typeof nameCompare.ambiguity_gap === 'number'
        ? nameCompare.ambiguity_gap
        : NAME_MATCH_AMBIGUITY_GAP;
    const outcomesKey = outcomesColumn ? `outcomes_${outcomesColumn}` : '';
    const wsuKey = wsuColumn ? `wsu_${wsuColumn}` : '';
    const outcomesStateKey = nameCompare.state_outcomes ? `outcomes_${nameCompare.state_outcomes}` : '';
    const wsuStateKey = nameCompare.state_wsu ? `wsu_${nameCompare.state_wsu}` : '';
    const outcomesCityKey = nameCompare.city_outcomes ? `outcomes_${nameCompare.city_outcomes}` : '';
    const wsuCityKey = nameCompare.city_wsu ? `wsu_${nameCompare.city_wsu}` : '';
    const outcomesCountryKey = nameCompare.country_outcomes ? `outcomes_${nameCompare.country_outcomes}` : '';
    const wsuCountryKey = nameCompare.country_wsu ? `wsu_${nameCompare.country_wsu}` : '';
    const canCompareNames = nameCompareEnabled && outcomesKey && wsuKey;
    const idfTable = canCompareNames
        ? buildTokenIDF([
            ...outcomes.map(row => row[outcomesColumn] || ''),
            ...wsuOrg.map(row => row[wsuColumn] || '')
        ].filter(Boolean))
        : null;
    const ambiguityScoreCache = canCompareNames ? new Map() : null;
    const missingOutputReplacementCache = canCompareNames ? new Map() : null;

    const validated = [];
    const totalRows = merged.length;
    let processed = 0;
    const reportEvery = 200;

    for (const row of merged) {
        const result = {
            ...row,
            Error_Type: 'Valid',
            Error_Description: 'Mapping is valid',
            Duplicate_Group: '',
            Error_Subtype: '',
            Suggested_Key: '',
            Suggested_School: '',
            Suggested_City: '',
            Suggested_State: '',
            Suggested_Country: '',
            Suggestion_Score: ''
        };

        if (!row.translate_input_norm) {
            result.Error_Type = 'Missing_Input';
            result.Error_Description = 'Translation input is missing';
        }

        if (result.Error_Type === 'Valid' && !row.translate_output_norm) {
            result.Error_Type = 'Missing_Output';
            result.Error_Description = 'Translation output is missing';
        }

        if (result.Error_Type === 'Valid' && row.translate_input_norm && !outcomesKeys.has(row.translate_input_norm)) {
            result.Error_Type = 'Input_Not_Found';
            result.Error_Description = 'Translation input does not exist in Outcomes data';
        }

        if (result.Error_Type === 'Valid' && row.translate_output_norm && !wsuKeys.has(row.translate_output_norm)) {
            result.Error_Type = 'Output_Not_Found';
            result.Error_Description = 'Translation output does not exist in myWSU data';
            result.Error_Subtype = OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT;

            if (canCompareNames && row[outcomesKey]) {
                const sourceName = row[outcomesKey];
                const sourceState = outcomesStateKey ? row[outcomesStateKey] : '';
                const sourceCity = outcomesCityKey ? row[outcomesCityKey] : '';
                const sourceCountry = outcomesCountryKey ? row[outcomesCountryKey] : '';
                const cacheKey = [
                    normalizeNameForCompare(sourceName),
                    normalizeStateValue(sourceState),
                    normalizeNameForCompare(sourceCity),
                    normalizeCountryValue(sourceCountry)
                ].join('|');

                let replacement = missingOutputReplacementCache.get(cacheKey);
                if (!replacement) {
                    replacement = classifyMissingOutputReplacement(
                        sourceName,
                        sourceState,
                        sourceCity,
                        sourceCountry,
                        wsuOrg,
                        {
                            nameField: wsuColumn,
                            keyField: keyConfig.wsu,
                            stateField: nameCompare.state_wsu || '',
                            cityField: nameCompare.city_wsu || '',
                            countryField: nameCompare.country_wsu || '',
                            idfTable,
                            threshold,
                            ambiguityGap: Math.max(STALE_KEY_REPLACEMENT_GAP, ambiguityGap)
                        }
                    );
                    missingOutputReplacementCache.set(cacheKey, replacement);
                }

                const suggestedKeyNorm = replacement.bestCandidate
                    ? normalizeKeyValue(replacement.bestCandidate.key || '')
                    : '';
                const sameAsCurrent = suggestedKeyNorm && row.translate_output_norm && suggestedKeyNorm === row.translate_output_norm;
                result.Error_Subtype = (sameAsCurrent ? OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT : replacement.subtype) || OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT;
                if (
                    replacement.subtype === OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY &&
                    replacement.bestCandidate &&
                    !sameAsCurrent
                ) {
                    result.Suggested_Key = replacement.bestCandidate.key || '';
                    result.Suggested_School = replacement.bestCandidate.name || '';
                    result.Suggested_City = replacement.bestCandidate.city || '';
                    result.Suggested_State = replacement.bestCandidate.state || '';
                    result.Suggested_Country = replacement.bestCandidate.country || '';
                    result.Suggestion_Score = replacement.bestCandidate.score;
                    const scorePct = Math.round((replacement.bestCandidate.score || 0) * 100);
                    const locationParts = [
                        replacement.bestCandidate.city,
                        replacement.bestCandidate.state,
                        replacement.bestCandidate.country
                    ].filter(Boolean).join(', ');
                    const locationSuffix = locationParts ? ` (${locationParts})` : '';
                    result.Error_Description = `Translation output does not exist in myWSU data. Likely stale key; suggested replacement ${result.Suggested_Key}: "${result.Suggested_School}"${locationSuffix} (score: ${scorePct}%).`;
                } else if (replacement.subtype === OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT) {
                    if (replacement.bestCandidate && !sameAsCurrent) {
                        result.Suggested_Key = replacement.bestCandidate.key || '';
                        result.Suggested_School = replacement.bestCandidate.name || '';
                        result.Suggested_City = replacement.bestCandidate.city || '';
                        result.Suggested_State = replacement.bestCandidate.state || '';
                        result.Suggested_Country = replacement.bestCandidate.country || '';
                        result.Suggestion_Score = replacement.bestCandidate.score;
                        result.Error_Description = 'Translation output does not exist in myWSU data. Multiple high-confidence replacement candidates were found; top candidate pre-filled for Use Suggestion.';
                    } else {
                        result.Error_Description = 'Translation output does not exist in myWSU data. Multiple high-confidence replacement candidates were found; review manually.';
                    }
                } else {
                    result.Error_Description = 'Translation output does not exist in myWSU data. No high-confidence replacement candidate was found.';
                }
            }
        }

        const duplicateSourceCount = duplicateSourcesDict[row.translate_input_norm]?.length || 0;
        if (result.Error_Type === 'Valid' && row.translate_input_norm && duplicateSourceCount > 1) {
            result.Error_Type = 'Duplicate_Source';
            result.Error_Description = `Source key maps to ${duplicateSourceCount} different target keys (one-to-many)`;
            result.Duplicate_Group = duplicateGroups[`src:${row.translate_input_norm}`] || '';
        }

        const duplicateTargetCount = duplicateTargetsDict[row.translate_output_norm]?.length || 0;
        if (result.Error_Type === 'Valid' && row.translate_output_norm && duplicateTargetCount > 1) {
            result.Error_Type = 'Duplicate_Target';
            result.Error_Description = `Target key maps to ${duplicateTargetCount} different source keys (many-to-one)`;
            result.Duplicate_Group = duplicateGroups[`tgt:${row.translate_output_norm}`] || '';
        }

        if (result.Error_Type === 'Valid' && canCompareNames && row[outcomesKey] && row[wsuKey]) {
            const stateValue1 = outcomesStateKey ? row[outcomesStateKey] : '';
            const stateValue2 = wsuStateKey ? row[wsuStateKey] : '';
            const cityValue1 = outcomesCityKey ? row[outcomesCityKey] : '';
            const cityValue2 = wsuCityKey ? row[wsuCityKey] : '';
            const countryValue1 = outcomesCountryKey ? row[outcomesCountryKey] : '';
            const countryValue2 = wsuCountryKey ? row[wsuCountryKey] : '';
            const evidence = fieldEvidence(
                row[outcomesKey],
                row[wsuKey],
                stateValue1,
                stateValue2,
                cityValue1,
                cityValue2,
                countryValue1,
                countryValue2
            );
            const effectiveThreshold = Math.max(
                0,
                Math.min(
                    1,
                    threshold + (evidence > 0.7 ? -0.03 : evidence < 0.3 ? 0.03 : 0)
                )
            );
            const similarity = calculateNameSimilarity(row[outcomesKey], row[wsuKey], idfTable);

            if (similarity < effectiveThreshold) {
                if (isHighConfidenceNameMatch(
                    row[outcomesKey],
                    row[wsuKey],
                    stateValue1,
                    stateValue2,
                    cityValue1,
                    cityValue2,
                    countryValue1,
                    countryValue2,
                    similarity,
                    effectiveThreshold,
                    idfTable
                )) {
                    result.Error_Type = 'High_Confidence_Match';
                    result.Error_Description = 'High confidence match based on name + state';
                } else {
                    result.Error_Type = 'Name_Mismatch';
                    result.Error_Description = `Names do not match (similarity: ${Math.round(similarity * 100)}%). "${row[outcomesKey]}" mapped to "${row[wsuKey]}" - verify this is correct`;
                }
            } else if (
                isAmbiguousNameMatch(
                    row[outcomesKey],
                    row[wsuKey],
                    wsuOrg,
                    wsuColumn,
                    effectiveThreshold,
                    ambiguityGap,
                    ambiguityScoreCache,
                    similarity,
                    idfTable
                )
            ) {
                if (!isHighConfidenceNameMatch(
                    row[outcomesKey],
                    row[wsuKey],
                    stateValue1,
                    stateValue2,
                    cityValue1,
                    cityValue2,
                    countryValue1,
                    countryValue2,
                    similarity,
                    effectiveThreshold,
                    idfTable
                )) {
                    result.Error_Type = 'Ambiguous_Match';
                    result.Error_Description = `Ambiguous name match (similarity: ${Math.round(similarity * 100)}%). Another candidate is within ${Math.round(ambiguityGap * 100)}% - review alternatives`;
                }
            }
        }

        validated.push(result);

        processed += 1;
        if (onProgress && (processed % reportEvery === 0 || processed === totalRows)) {
            onProgress(processed, totalRows);
        }
    }

    console.log('\n=== Validation Complete ===');
    console.log(`Total rows validated: ${validated.length}`);

    const errorCounts = {};
    validated.forEach(row => {
        errorCounts[row.Error_Type] = (errorCounts[row.Error_Type] || 0) + 1;
    });

    console.log('\nError Type Breakdown:');
    for (const [type, count] of Object.entries(errorCounts)) {
        console.log(`  ${type}: ${count}`);
    }

    return validated;
}

function generateSummaryStats(validated, outcomes, translate, wsuOrg) {
    const totalMappings = validated.length;
    const validCount = validated.filter(r => (
        r.Error_Type === 'Valid' || r.Error_Type === 'High_Confidence_Match'
    )).length;
    const errorCount = totalMappings - validCount;
    const validPct = totalMappings > 0
        ? Math.round((validCount / totalMappings) * 1000) / 10
        : 0;
    const errorPct = totalMappings > 0
        ? Math.round((errorCount / totalMappings) * 1000) / 10
        : 0;

    const errorCounts = {};
    validated.forEach(row => {
        errorCounts[row.Error_Type] = (errorCounts[row.Error_Type] || 0) + 1;
    });
    const outputNotFoundRows = validated.filter(row => row.Error_Type === 'Output_Not_Found');
    const outputNotFoundLikelyStale = outputNotFoundRows.filter(
        row => row.Error_Subtype === OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY
    ).length;
    const outputNotFoundAmbiguous = outputNotFoundRows.filter(
        row => row.Error_Subtype === OUTPUT_NOT_FOUND_SUBTYPE.AMBIGUOUS_REPLACEMENT
    ).length;
    const outputNotFoundNoReplacement = outputNotFoundRows.filter(
        row => row.Error_Subtype === OUTPUT_NOT_FOUND_SUBTYPE.NO_REPLACEMENT
    ).length;

    const stats = {
        timestamp: new Date().toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        }),
        files: {
            outcomes_rows: outcomes.length,
            translate_rows: translate.length,
            wsu_org_rows: wsuOrg.length
        },
        validation: {
            total_mappings: totalMappings,
            valid_count: validCount,
            valid_percentage: validPct,
            error_count: errorCount,
            error_percentage: errorPct
        },
        errors: {
            missing_inputs: errorCounts.Missing_Input || 0,
            missing_outputs: errorCounts.Missing_Output || 0,
            input_not_found: errorCounts.Input_Not_Found || 0,
            output_not_found: errorCounts.Output_Not_Found || 0,
            output_not_found_likely_stale_key: outputNotFoundLikelyStale,
            output_not_found_ambiguous_replacement: outputNotFoundAmbiguous,
            output_not_found_no_replacement: outputNotFoundNoReplacement,
            duplicate_targets: errorCounts.Duplicate_Target || 0,
            duplicate_sources: errorCounts.Duplicate_Source || 0,
            name_mismatches: errorCounts.Name_Mismatch || 0,
            ambiguous_matches: errorCounts.Ambiguous_Match || 0,
            high_confidence_matches: errorCounts.High_Confidence_Match || 0
        }
    };

    return stats;
}

function getErrorSamples(validated, limit = 10) {
    const samples = {};
    const errorTypes = [
        'Missing_Input',
        'Missing_Output',
        'Input_Not_Found',
        'Output_Not_Found',
        'Duplicate_Target',
        'Duplicate_Source',
        'Name_Mismatch',
        'Ambiguous_Match'
    ];

    const resolvedLimit = limit && limit > 0 ? limit : null;

    errorTypes.forEach(errorType => {
        const rows = validated.filter(r => r.Error_Type === errorType);
        const showing = resolvedLimit ? Math.min(rows.length, resolvedLimit) : rows.length;
        samples[errorType] = {
            count: rows.length,
            showing,
            rows: rows.slice(0, showing).map(r => ({
                translate_input: r.translate_input,
                translate_output: r.translate_output,
                Error_Description: r.Error_Description
            }))
        };
    });

    const staleKeyRows = validated.filter(r => (
        r.Error_Type === 'Output_Not_Found' &&
        r.Error_Subtype === OUTPUT_NOT_FOUND_SUBTYPE.LIKELY_STALE_KEY
    ));
    const staleShowing = resolvedLimit ? Math.min(staleKeyRows.length, resolvedLimit) : staleKeyRows.length;
    samples.Output_Not_Found_Likely_Stale_Key = {
        count: staleKeyRows.length,
        showing: staleShowing,
        rows: staleKeyRows.slice(0, staleShowing).map(r => ({
            translate_input: r.translate_input,
            translate_output: r.translate_output,
            Error_Description: r.Error_Description
        }))
    };

    return samples;
}
