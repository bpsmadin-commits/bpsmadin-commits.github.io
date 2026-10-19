// Lazy-load library berat: baru diambil pas benar-benar dipakai
const CDN = {
    xlsx: 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
    jspdf: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
    html2canvas: 'https://html2canvas.hertzen.com/dist/html2canvas.min.js'
};
const loadedScripts = {};

function loadScript(url) {
    if (loadedScripts[url]) return loadedScripts[url];
    loadedScripts[url] = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = url;
        s.onload = resolve;
        s.onerror = () => reject(new Error('Gagal memuat modul: ' + url));
        document.head.appendChild(s);
    });
    return loadedScripts[url];
}

function mergeCachedSheetData(sheetList, sheetsData) {
    let subjects = [];
    let scores = [];
    let studentNames = [];
    let studentMeta = [];
    let studentCount = 0;
    for (const name of sheetList) {
        const data = sheetsData[name];
        if (!data || !data.subjects.length || !data.scores.length) continue;
        if (!subjects.length) subjects = data.subjects.slice();
        if (data.subjects.length === subjects.length) {
            scores = scores.concat(data.scores);
            studentNames = studentNames.concat(data.studentNames || data.scores.map(() => ''));
            studentMeta = studentMeta.concat(data.studentMeta || data.scores.map(() => ({})));
            studentCount += data.studentCount || 0;
        }
    }
    return { subjects, averages: computeAverages(subjects, scores), studentCount, scores, studentNames, studentMeta };
}

async function ensureXLSX() {
    if (typeof XLSX === 'undefined') await loadScript(CDN.xlsx);
}

async function ensurePDFLibs() {
    if (typeof window.jspdf === 'undefined') await loadScript(CDN.jspdf);
    if (typeof html2canvas === 'undefined') await loadScript(CDN.html2canvas);
}

// Global Variables
let workbookData = null;
let currentSheetName = "";
let currentAverages = [];
let currentSubjectNames = [];
let currentStudentCount = 0;
let KKM = 50;
const DEFAULT_KKM = 50;

let waliKelas = {};

function loadWaliKelas() {
    try {
        const raw = localStorage.getItem(WALI_STORAGE_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch (e) {
        return {};
    }
}

function saveWaliKelas() {
    try {
        localStorage.setItem(WALI_STORAGE_KEY, JSON.stringify(waliKelas));
    } catch (e) {
        console.warn('Gagal simpan wali kelas:', e);
    }
}

let groupedData = {};
let allSheetsData = {};
let originalSheetNames = [];

let currentQuartal = 'Kuartal_I';
let quartalCache = {};
const QUARTAL_LABELS = {
    Kuartal_I: 'Kuartal I',
    Kuartal_II: 'Kuartal II',
    Kuartal_III: 'Kuartal III'
};

const STORAGE_PREFIX = 'dashboard_akademik_';
const WALI_STORAGE_KEY = STORAGE_PREFIX + 'wali_kelas';

// DOM References
const quartalTabs = document.getElementById('quartalTabs');
const fileInput = document.getElementById('file-input');
const uploadBtnText = document.getElementById('upload-btn-text');
const filterBar = document.getElementById('filterBar');
const sheetSelect = document.getElementById('sheet-select');
const jenjangDropdownBar = document.getElementById('jenjangDropdownBar');
const jenjangSelect = document.getElementById('jenjang-select');
const mapelSelect = document.getElementById('mapelSelect');
const selectAllMapelBtn = document.getElementById('selectAllMapel');
const deselectAllMapelBtn = document.getElementById('deselectAllMapel');
const grafikSelectAllMapelBtn = document.getElementById('grafikSelectAllMapel');
const grafikDeselectAllMapelBtn = document.getElementById('grafikDeselectAllMapel');
const tableBodyKelas = document.getElementById('table-body-kelas');
const kelasKetuntasanSummaryBox = document.getElementById('kelasKetuntasanSummaryBox');
const kelasKetuntasanBox = document.getElementById('kelasKetuntasanBox');
const kelasMiniStats = document.getElementById('kelasMiniStats');
const kelasQuickSummary = document.getElementById('kelasQuickSummary');
const legerSheetSelect = document.getElementById('legerSheetSelect');
const legerContainer = document.getElementById('legerContainer');
const jenjangStudentSearchInput = document.getElementById('jenjangStudentSearchInput');
const jenjangStudentSearchResults = document.getElementById('jenjangStudentSearchResults');
const metaDataStatus = document.getElementById('meta-data-status');
const metaSheet = document.getElementById('meta-sheet');
const kelasTableTitle = document.getElementById('kelasTableTitle');
const themeBtns = document.querySelectorAll('.theme-btn');
const pdfContainer = document.getElementById('pdf-export-container');
const kkmInput = document.getElementById('kkmInput');
const resetKkmBtn = document.getElementById('resetKkmBtn');
const jenjangContainer = document.getElementById('jenjangContainer');
const grafikJenjangSelect = document.getElementById('grafikJenjangSelect');
const grafikMapelSelect = document.getElementById('grafikMapelSelect');
const grafikPerjenjangContainer = document.getElementById('grafikPerjenjangContainer');

// UI State Management
function getUIState() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_PREFIX + 'ui_state')) || {};
    } catch (e) {
        return {};
    }
}

function saveUIState() {
    const state = {
        quartal: currentQuartal,
        sheet: currentSheetName,
        kkm: KKM,
        kkmScale: 100
    };
    localStorage.setItem(STORAGE_PREFIX + 'ui_state', JSON.stringify(state));
}

function saveToLocalStorage() {
    const data = { groupedData, allSheetsData, originalSheetNames, waliKelas };
    try {
        localStorage.setItem(STORAGE_PREFIX + currentQuartal, JSON.stringify(data));
    } catch (e) {
        console.warn('Gagal simpan ke localStorage:', e);
    }
}

function loadFromLocalStorage(quartal) {
    try {
        const raw = localStorage.getItem(STORAGE_PREFIX + quartal);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch (e) {
        return null;
    }
}

function clearLocalStorage(quartal) {
    localStorage.removeItem(STORAGE_PREFIX + quartal);
}

// Data Processing
function normalizeCellText(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/[\uFEFF\u00A0]/g, ' ').trim();
}

function parseAsuhanText(raw) {
    const text = normalizeCellText(raw || '');
    if (!text) return '';
    const splitIdx = text.search(/[:\-]/);
    if (splitIdx !== -1 && splitIdx > 0) {
        const candidate = text.substring(splitIdx + 1).trim();
        if (candidate) return candidate;
    }
    return text;
}

function detectAsuhanFromRows(rows) {
    const labelPattern = /(?:ASUHAN|WALI\s*KELAS|GURU\s*KELAS)\s*[:\-]?\s*(.*)/i;
    const scanEnd = Math.min(rows.length, 15);

    for (let r = 0; r < scanEnd; r++) {
        const row = rows[r];
        if (!row) continue;

        for (let c = 0; c < row.length; c++) {
            const cell = normalizeCellText(row[c]);
            const match = cell.match(labelPattern);
            if (!match) continue;

            const inlineName = normalizeCellText(match[1]);
            if (inlineName) return inlineName;

            for (let next = c + 1; next < row.length; next++) {
                const value = normalizeCellText(row[next]);
                if (value) return value;
            }
        }
    }

    return '';
}

function normalizeWaliClassKey(value) {
    const roman = { I: '1', II: '2', III: '3', IV: '4', V: '5', VI: '6', VII: '7', VIII: '8', IX: '9' };
    const parts = normalizeCellText(value).toUpperCase()
        .replace(/\bKLS?\b|\bKELAS\b/g, ' ')
        .replace(/\bMDTA\b/g, 'MID')
        .replace(/\bMTSD\b/g, 'MDTW')
        .replace(/[^A-Z0-9]+/g, ' ')
        .trim().split(/\s+/).filter(Boolean);
    parts.forEach((part, index) => {
        if (roman[part]) parts[index] = roman[part];
    });
    return parts.join(' ');
}

function readWaliKelasSheet(worksheet) {
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const headerIndex = rows.findIndex(row => {
        const headers = row.map(normalizeCellText).map(value => value.toUpperCase());
        return headers.includes('NAMA') && headers.some(value => value.includes('KELAS'));
    });
    if (headerIndex < 0) return {};

    const headers = rows[headerIndex].map(normalizeCellText);
    const nameIndex = headers.findIndex(value => /^NAMA(?: WALI)?$/i.test(value));
    const classIndexes = headers.reduce((indexes, value, index) => {
        if (/KELAS|KLS/i.test(value)) indexes.push(index);
        return indexes;
    }, []);
    const positionIndex = headers.findIndex(value => /JABATAN/i.test(value));
    const mapping = {};

    rows.slice(headerIndex + 1).forEach(row => {
        const name = normalizeCellText(row[nameIndex]);
        const position = positionIndex >= 0 ? normalizeCellText(row[positionIndex]) : '';
        if (!name || (position && !/WALI\s*KELAS/i.test(position))) return;
        classIndexes.forEach(index => {
            const classKey = normalizeWaliClassKey(row[index]);
            if (classKey) mapping[classKey] = name;
        });
    });
    return mapping;
}

function findWaliKelas(sheetName, data, mapping) {
    const candidates = [sheetName, formatClassLabel(sheetName), ...(data?.studentMeta || []).map(meta => meta.kelas)];
    for (const candidate of candidates) {
        const key = normalizeWaliClassKey(candidate);
        if (mapping[key]) return mapping[key];
        const matchingKey = Object.keys(mapping).find(mappedKey =>
            key === mappedKey || key.endsWith(` ${mappedKey}`) || mappedKey.endsWith(` ${key}`));
        if (matchingKey) return mapping[matchingKey];
    }
    return '';
}

function detectHeaderAndData(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z1000');
    const rows = [];

    for (let R = range.s.r; R <= range.e.r; R++) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];
            let val = cell ? cell.v : "";
            if (cell && cell.t === 'n') val = cell.v;
            else if (cell && cell.t === 's') val = cell.v;
            else val = (val === undefined) ? "" : val;
            row.push(val);
        }
        rows.push(row);
    }

    const subjectKeywords = [
        "فتح القريب", "تفسير الجلالين", "شرح ام البراهين", "لتبيان فى علوم القران", "علم الفرائض",
        "الخلفاء الراشدون", "الكواكب الدرية", "ستين مسئلة", "بلوغ المرام", "إملاء", "إيضاح المبهم",
        "قــرأة الكتاب", "مبادئ الفقه", "خريدة البهية", "هداية المستفيد", "اللغة العربية",
        "قاعد مختصر جدا", "خلاصة نور اليقين", "رابطة الحروف", "عقيدة التوحيد", "تحفة الاطفال",
        "الديانة", "هداية الصبيان", "عقيدة العوام", "جواهر الكلامية", "التصريف", "رياض البديعة"
    ];

    let headerRowIndex = -1,
        dataStartRow = -1;

    for (let i = 0; i < rows.length && i < 50; i++) {
        for (let j = 0; j < rows[i].length && j < 15; j++) {
            const cell = String(rows[i][j]).trim().toUpperCase();
            if (cell === 'NAMA' || cell === 'NAMA SISWA') {
                headerRowIndex = i;
                dataStartRow = i + 1;
                break;
            }
        }
        if (headerRowIndex !== -1) break;
    }

    if (headerRowIndex === -1) {
        for (let i = 0; i < rows.length && i < 50; i++) {
            let matchCount = 0;
            for (let j = 0; j < rows[i].length && j < 30; j++) {
                const cell = String(rows[i][j]).trim();
                if (cell.length > 0 && subjectKeywords.some(kw => cell.includes(kw))) matchCount++;
            }
            if (matchCount >= 3) { headerRowIndex = i;
                dataStartRow = i + 1; break; }
        }
    }

    if (headerRowIndex === -1 && rows.length > 6) { headerRowIndex = 5;
        dataStartRow = 6; }
    if (headerRowIndex === -1 || dataStartRow === -1) return { subjects: [], scores: [], studentNames: [] };

    const headerRow = rows[headerRowIndex];
    const subjectIndices = [],
        subjectNames = [];
    const nameColIdx = headerRow.findIndex(cell => ['NAMA', 'NAMA SISWA'].includes(normalizeCellText(cell).toUpperCase()));
    const metaColumnIndices = {
        nopes: headerRow.findIndex(cell => /^(NO\s*PES|NOPES|NOMOR PESERTA|NO PESERTA)$/i.test(normalizeCellText(cell))),
        nis: headerRow.findIndex(cell => /^(NIS|NISN)$/i.test(normalizeCellText(cell))),
        asr: headerRow.findIndex(cell => /^ASR$/i.test(normalizeCellText(cell))),
        kelas: headerRow.findIndex(cell => /KELAS|KLS/i.test(normalizeCellText(cell))),
    };
    const ignoreKeywords = ["NO", "NOPES", "NIS", "NAMA", "ASR", "KELAS", "KLS", "RUANG", "ASUHAN", "الترتيب",
        "جملة", "ألتسوية", "فائز", "RANK", "SUM", "jml"
    ];

    let asuhanColIdx = -1;
    for (let c = 0; c < headerRow.length; c++) {
        let cellVal = headerRow[c];
        if (cellVal && typeof cellVal === 'string') {
            let cleaned = cellVal.trim().toUpperCase();
            if (cleaned === 'ASUHAN' || cleaned.includes('ASUHAN')) {
                asuhanColIdx = c;
                break;
            }
        }
    }

    for (let c = 0; c < headerRow.length; c++) {
        let cellVal = headerRow[c];
        if (cellVal && typeof cellVal === 'string') {
            let cleaned = cellVal.trim();
            if (!ignoreKeywords.some(ig => cleaned.toUpperCase().includes(ig)) && cleaned.length > 1 && !/^\d+(\.\d+)?$/
                .test(cleaned)) {
                subjectIndices.push(c);
                subjectNames.push(cleaned);
            }
        }
    }

    if (subjectNames.length < 3 && headerRow.length > 6) {
        subjectIndices.length = 0;
        subjectNames.length = 0;
        for (let c = 6; c < Math.min(headerRow.length, 30); c++) {
            let name = headerRow[c];
            if (name && typeof name === 'string' && name.trim().length > 0 &&
                !ignoreKeywords.some(ig => name.trim().toUpperCase().includes(ig))) {
                subjectIndices.push(c);
                subjectNames.push(name.trim());
            }
        }
    }

    const studentScores = [];
    const studentNames = [];
    const studentMeta = [];
    for (let r = dataStartRow; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length === 0) continue;
        let hasNumeric = false;
        const scoreRow = [];
        for (let idx of subjectIndices) {
            let val = row[idx];
            let num = null;
            if (val !== undefined && val !== "") {
                if (typeof val === 'number') { num = val;
                    hasNumeric = true; } else if (typeof val === 'string') { let p = parseFloat(val); if (!isNaN(
                    p)) { num = p;
                    hasNumeric = true; } } else if (typeof val === 'boolean') { num = val ? 1 : 0;
                hasNumeric = true; }
            }
            scoreRow.push(num);
        }
        if (hasNumeric) {
            studentScores.push(scoreRow);
            studentNames.push(normalizeCellText(row[nameColIdx]));
            studentMeta.push(Object.fromEntries(Object.entries(metaColumnIndices)
                .filter(([, colIndex]) => colIndex !== -1)
                .map(([key, colIndex]) => [key, normalizeCellText(row[colIndex])])));
        }
    }

    const asuhanRaw = (asuhanColIdx !== -1 && headerRowIndex !== -1 && rows[headerRowIndex]) ?
        rows[headerRowIndex][asuhanColIdx] : '';

    let asuhanVal = detectAsuhanFromRows(rows);
    if (!asuhanVal && asuhanRaw !== '') {
        asuhanVal = parseAsuhanText(asuhanRaw);
    }

    return { subjects: subjectNames, scores: studentScores, studentNames, studentMeta,
        studentCount: studentScores.length, asuhan: asuhanVal };
}

function computeAverages(subjects, scores) {
    if (!scores.length || !subjects.length) return [];
    const sums = new Array(subjects.length).fill(0);
    const counts = new Array(subjects.length).fill(0);
    for (const student of scores) {
        for (let i = 0; i < subjects.length; i++) {
            const val = student[i];
            if (val !== null && typeof val === 'number' && !isNaN(val)) { sums[i] += val;
                counts[i]++; }
        }
    }
    return sums.map((sum, idx) => counts[idx] ? parseFloat((sum / counts[idx]).toFixed(2)) : 0);
}

function sortSubjectsByAverageDesc(subjects, averages) {
    const pairs = subjects.map((s, i) => ({ subject: s, avg: averages[i] }));
    pairs.sort((a, b) => b.avg - a.avg);
    return { subjects: pairs.map(p => p.subject), averages: pairs.map(p => p.avg) };
}

function mergeSheetData(sheetList) {
    let allSubjects = [],
        allScores = [],
        allStudentNames = [],
        allStudentMeta = [],
        totalStudents = 0;
    for (const sheetName of sheetList) {
        const ws = workbookData.Sheets[sheetName];
        if (!ws) continue;
        const { subjects, scores, studentNames, studentMeta, studentCount } = detectHeaderAndData(ws);
        if (subjects.length === 0 || scores.length === 0) continue;
        if (allSubjects.length === 0) allSubjects = subjects.slice();
        for (let i = 0; i < scores.length; i++) {
            if (scores[i].length === allSubjects.length) {
                allScores.push(scores[i].slice());
                allStudentNames.push(studentNames[i] || '');
                allStudentMeta.push(studentMeta?.[i] || {});
            }
        }
        totalStudents += studentCount;
    }
    return { subjects: allSubjects, averages: computeAverages(allSubjects, allScores), studentCount: totalStudents,
        scores: allScores, studentNames: allStudentNames, studentMeta: allStudentMeta };
}

function computeStudentAverages(scores) {
    if (!scores || !scores.length) return [];
    return scores.map(student => {
        let sum = 0;
        for (let val of student) { if (typeof val === 'number' && !isNaN(val)) { sum += val; } }
        return student.length > 0 ? sum / student.length : 0;
    });
}

function renderLeger(sheetName = legerSheetSelect.value) {
    const data = allSheetsData[sheetName];
    if (!data?.scores?.length) {
        legerContainer.innerHTML = '<div class="leger-empty">Belum ada data nilai untuk dibuat leger.</div>';
        return;
    }

    const metaFields = [
        ['nopes', 'NOPES'],
        ['nis', 'NIS'],
        ['asr', 'ASR']
    ].filter(([key]) => (data.studentMeta || []).some(meta => meta?.[key]));
    const headers = ['No', ...metaFields.map(([, label]) => label), 'Nama', 'Kelas',
        ...data.subjects, 'Jumlah', 'Rata-rata', 'Status'];
    const rows = data.scores.map((scores, index) => {
        const meta = data.studentMeta?.[index] || {};
        const studentName = data.studentNames?.[index] || 'Nama tidak terbaca';
        const average = computeStudentAverages([scores])[0] || 0;
        const total = scores.reduce((sum, score) => sum + (typeof score === 'number' && Number.isFinite(score) ? score : 0), 0);
        const values = [
            String(index + 1),
            ...metaFields.map(([key]) => meta[key] || ''),
            studentName,
            meta.kelas || formatClassLabel(sheetName),
            ...scores.map(score => typeof score === 'number' && Number.isFinite(score) ? score.toFixed(0) : ''),
            total.toFixed(1),
            average.toFixed(1),
            average >= KKM ? 'Tuntas' : 'Tidak Tuntas'
        ];
        return { values, name: studentName, average };
    }).sort((a, b) => b.average - a.average);

    const tableHTML = `
        <div class="leger-table-wrap">
            <table class="leger-table">
                <thead><tr>${headers.map(header => `<th>${arabicTextHTML(header)}</th>`).join('')}</tr></thead>
                <tbody>${rows.map(row => `<tr class="${row.average >= KKM ? 'leger-passed' : 'leger-failed'}">${row.values.map(value => `<td>${arabicTextHTML(value)}</td>`).join('')}</tr>`).join('')}</tbody>
            </table>
        </div>`;
    const wali = waliKelas[sheetName] || data.asuhan || '';
    legerContainer.innerHTML = `
        <div class="leger-heading">
            <div>
                <h2>Leger Nilai Ujian — ${escapeHTML(formatClassLabel(sheetName))}</h2>
                <p>${rows.length} siswa${wali ? ` · Wali Kelas: ${arabicTextHTML(wali)}` : ''}</p>
            </div>
        </div>
        ${tableHTML}`;
}

function safeExcelSheetName(name, usedNames) {
    const base = String(name || 'Kelas').replace(/[\\/*?:\[\]]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'Kelas';
    let result = base;
    let suffix = 2;
    while (usedNames.has(result)) {
        const tail = ` (${suffix++})`;
        result = `${base.slice(0, 31 - tail.length)}${tail}`;
    }
    usedNames.add(result);
    return result;
}

async function exportAllLegerToExcel() {
    const sheetNames = originalSheetNames.filter(name => allSheetsData[name]?.scores?.length);
    if (!sheetNames.length) {
        alert('Tidak ada data leger untuk diekspor.');
        return;
    }

    const button = document.getElementById('exportExcelLeger');
    button.disabled = true;
    try {
        await ensureXLSX();
        const workbook = XLSX.utils.book_new();
        const usedNames = new Set();

        sheetNames.forEach(sheetName => {
            const data = allSheetsData[sheetName];
            const wali = waliKelas[sheetName] || data.asuhan || '';
            const headers = ['NO', 'NOPES', 'NAMA', 'ASR', 'KLS',
                ...data.subjects, 'جملة', 'التسوية', 'الترتيب', 'STATUS'];
            const averages = computeStudentAverages(data.scores || []);
            const rows = (data.scores || []).map((scores, index) => {
                const meta = data.studentMeta?.[index] || {};
                const total = scores.reduce((sum, score) => sum + (typeof score === 'number' && Number.isFinite(score) ? score : 0), 0);
                const average = averages[index] || 0;
                return {
                    average,
                    values: [index + 1, meta.nopes || '',
                        data.studentNames?.[index] || 'Nama tidak terbaca',
                        meta.asr || '', meta.kelas || formatClassLabel(sheetName),
                        ...scores.map(score => typeof score === 'number' && Number.isFinite(score) ? score : ''),
                        total, average, index + 1, average >= KKM ? 'Tuntas' : 'Tidak Tuntas']
                };
            }).sort((a, b) => b.average - a.average);

            const titleRows = [
                ['LEGER NILAI UJI TULIS KUARTAL'],
                ['MADRASAH DINIYAH TAKMILIYAH (MDT) MIFTAHUL ULUM'],
                ['TAHUN PELAJARAN 1447 - 1448 H.'],
                [wali ? `WALI KELAS: ${wali}` : '']
            ];
            const headerRow = titleRows.length;
            const worksheet = XLSX.utils.aoa_to_sheet([...titleRows, headers, ...rows.map(row => row.values)]);
            const firstDataRow = headerRow + 2;
            const subjectStartCol = 5;
            const totalCol = subjectStartCol + data.subjects.length;
            const averageCol = totalCol + 1;
            const rankingCol = averageCol + 1;
            const statusCol = rankingCol + 1;
            rows.forEach((row, index) => {
                const excelRow = firstDataRow + index;
                const firstSubject = XLSX.utils.encode_cell({ r: excelRow - 1, c: subjectStartCol });
                const lastSubject = XLSX.utils.encode_cell({ r: excelRow - 1, c: totalCol - 1 });
                const totalCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: totalCol });
                const averageCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: averageCol });
                const statusCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: statusCol });
                const rankingCell = XLSX.utils.encode_cell({ r: excelRow - 1, c: rankingCol });
                worksheet[totalCell] = { t: 'n', f: `ROUND(SUM(${firstSubject}:${lastSubject}),0)`, v: Math.round(row.values[5 + data.subjects.length]) };
                worksheet[averageCell] = { t: 'n', f: `IFERROR(${totalCell}/${data.subjects.length},0)`, v: row.values[6 + data.subjects.length] };
                worksheet[rankingCell] = { t: 'n', f: `RANK(${averageCell},$${XLSX.utils.encode_col(averageCol)}$${firstDataRow}:$${XLSX.utils.encode_col(averageCol)}$${firstDataRow + rows.length - 1},0)`, v: row.values[7 + data.subjects.length] };
                worksheet[statusCell] = { t: 's', f: `IF(${averageCell}>=${KKM},"Tuntas","Tidak Tuntas")`, v: row.values[8 + data.subjects.length] };
            });
            const lastCol = headers.length - 1;
            const lastRow = firstDataRow + rows.length - 1;
            const border = { style: 'thin', color: { rgb: '808080' } };
            const headerStyle = {
                fill: { fgColor: { rgb: '9BBB59' } },
                font: { bold: true, color: { rgb: '000000' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: { top: border, bottom: border, left: border, right: border }
            };
            const subjectHeaderStyle = {
                ...headerStyle,
                font: { name: 'Traditional Arabic', sz: 12, bold: true, color: { rgb: '000000' } },
                alignment: { horizontal: 'center', vertical: 'center', textRotation: 90, wrapText: true }
            };
            const cellStyle = {
                alignment: { horizontal: 'center', vertical: 'center' },
                border: { top: border, bottom: border, left: border, right: border }
            };
            const nameStyle = {
                alignment: { horizontal: 'left', vertical: 'center' },
                border: { top: border, bottom: border, left: border, right: border }
            };
            for (let col = 0; col <= lastCol; col++) {
                worksheet[XLSX.utils.encode_cell({ r: headerRow, c: col })].s =
                    col >= subjectStartCol && col <= rankingCol ? subjectHeaderStyle : headerStyle;
            }
            for (let rowIndex = firstDataRow - 1; rowIndex <= lastRow - 1; rowIndex++) {
                for (let col = 0; col <= lastCol; col++) {
                    const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: col })];
                    if (cell) cell.s = cellStyle;
                }
                const nameCell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 2 })];
                if (nameCell) nameCell.s = nameStyle;
            }
            for (let rowIndex = firstDataRow - 1; rowIndex <= lastRow - 1; rowIndex++) {
                const noCell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })];
                if (noCell) noCell.z = '0';
                for (let col = subjectStartCol; col < totalCol; col++) {
                    const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: col })];
                    if (cell) cell.z = '0';
                }
                worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: totalCol })].z = '0';
                worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: averageCol })].z = '0.0';
                worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: rankingCol })].z = '0';
                if (rows[rowIndex - (firstDataRow - 1)].average < KKM) {
                    const failedStyle = {
                        fill: { fgColor: { rgb: 'F4CCCC' } },
                        font: { color: { rgb: 'C00000' }, bold: true },
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: { top: border, bottom: border, left: border, right: border }
                    };
                    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: averageCol })].s = failedStyle;
                    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: statusCol })].s = failedStyle;
                }
            }
            worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: lastCol } },
                { s: { r: 2, c: 0 }, e: { r: 2, c: lastCol } },
                { s: { r: 3, c: 0 }, e: { r: 3, c: lastCol } }];
            worksheet['!cols'] = headers.map((header, index) => {
                if (index === 0) return { wch: 4 };
                if (index === 2) return { wch: 25 };
                if (index === 4) return { wch: 12 };
                if (index >= subjectStartCol && index < totalCol) return { wch: 4 };
                if (index === totalCol) return { wch: 6 };
                if (index === rankingCol) return { wch: 4 };
                return { wch: Math.max(8, Math.min(18, String(header).length + 3)) };
            });
            worksheet['!rows'] = [{ hpt: 22 }, { hpt: 20 }, { hpt: 18 }, { hpt: 18 }, { hpt: 88 }];
            worksheet['!autofilter'] = { ref: `${XLSX.utils.encode_cell({ r: headerRow, c: 0 })}:${XLSX.utils.encode_cell({ r: lastRow - 1, c: lastCol })}` };
            worksheet['!freeze'] = { xSplit: 6, ySplit: headerRow + 1 };
            XLSX.utils.book_append_sheet(workbook, worksheet, safeExcelSheetName(sheetName, usedNames));
        });

        XLSX.writeFile(workbook, `Leger_Semua_Kelas_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (error) {
        console.error('Gagal export Excel leger:', error);
        alert('Gagal membuat file Excel leger.');
    } finally {
        button.disabled = false;
    }
}

function buildLegerDropdown() {
    const selectedSheet = legerSheetSelect.value || currentSheetName;
    legerSheetSelect.innerHTML = '';
    const sheetNames = originalSheetNames.filter(name => allSheetsData[name]?.scores?.length);
    if (!sheetNames.length) {
        legerSheetSelect.disabled = true;
        legerContainer.innerHTML = '<div class="leger-empty">Upload data nilai untuk membuat leger.</div>';
        return;
    }

    legerSheetSelect.disabled = false;
    sheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = formatClassLabel(name);
        legerSheetSelect.appendChild(option);
    });
    legerSheetSelect.value = sheetNames.includes(selectedSheet) ? selectedSheet : sheetNames[0];
    renderLeger(legerSheetSelect.value);
}

function getClassNumber(sheetName) {
    const digits = sheetName.match(/\d+/);
    if (digits) return parseInt(digits[0], 10);

    const roman = sheetName.match(/\b(XII|XI|IX|VIII|VII|VI|X|IV|V|III|II|I)\b/i);
    if (!roman) return 0;
    return { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10, XI: 11, XII: 12 }[roman[0].toUpperCase()];
}

function getClassType(sheetName) {
    const specificType = sheetName.match(/MDTA|MID|MDTW|MTSD/i);
    if (specificType) {
        const type = specificType[0].toUpperCase();
        return type === 'MDTA' || type === 'MID' ? 'MID' : 'MDTW';
    }
    return /MTS/i.test(sheetName) ? 'MDTW' : 'MID';
}

function getClassGroupKey(sheetName) {
    return `${getClassNumber(sheetName)}_${getClassType(sheetName)}`;
}

function formatClassLabel(sheetName) {
    return sheetName.replace(/^KLS\b/i, 'Kelas').replace(/[-_]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function getScoreScale() {
    if (!Number.isFinite(KKM)) return 100;
    return KKM > 10 ? 100 : 10;
}

function normalizeScoreToPercent(avg) {
    const scale = getScoreScale();
    const safeAvg = Number.isFinite(avg) ? avg : 0;
    if (scale <= 0) return 0;
    return Math.min(100, (safeAvg / scale) * 100);
}

function getTuntasCount(studentAverages) {
    return studentAverages.filter(avg => avg >= KKM).length;
}

function sortJenjangKeys(keys) {
    return keys.slice().sort((a, b) => {
        const [numA, jenisA] = a.split('_');
        const [numB, jenisB] = b.split('_');
        if (jenisA !== jenisB) return jenisA.localeCompare(jenisB);
        return Number(numA) - Number(numB);
    });
}

// UI Helpers
function getProgressBarClass(avg) {
    const scale = getScoreScale();
    const goodThreshold = scale > 10 ? 75 : scale * 0.75;
    if (avg >= goodThreshold) return 'good';
    if (avg >= KKM) return 'ok';
    return 'bad';
}

function getProgressBarHTML(avg) {
    const pct = normalizeScoreToPercent(avg);
    const cls = getProgressBarClass(avg);
    return `<div class="progress-cell"><div class="progress-bar-wrap"><div class="progress-bar-fill ${cls}" style="width:0%" data-target-width="${pct.toFixed(1)}%"></div></div><strong class="count-value" data-count-to="${avg}" data-decimals="1">0.0</strong></div>`;
}

// Counter animation dibuang supaya render lebih ringan di dataset besar.
// Angka langsung ditampilkan final, bar masih punya transisi CSS yang ringan.
function animateCountValue(el, target, decimals, suffix) {
    el.textContent = target.toFixed(decimals) + suffix;
}

// Find every [data-count-to] / progress bar inside `scope` and set them to final state.
// Ini menghindari requestAnimationFrame yang berat di banyak elemen.
function runCountUps(scope) {
    const root = scope || document;

    const bars = root.querySelectorAll('.progress-bar-fill[data-target-width]');
    bars.forEach(bar => {
        bar.style.transition = 'none';
        bar.style.width = bar.dataset.targetWidth || '0%';
    });

    const jcBars = root.querySelectorAll('.jc-bar-fill[data-target-height]');
    jcBars.forEach(bar => {
        bar.style.transition = 'none';
        bar.style.height = bar.dataset.targetHeight || '0%';
    });

    root.querySelectorAll('.count-value[data-count-to]').forEach(el => {
        const target = parseFloat(el.dataset.countTo);
        if (isNaN(target)) return;
        const decimals = el.dataset.decimals !== undefined ? parseInt(el.dataset.decimals, 10) : 2;
        const suffix = el.dataset.suffix || '';
        animateCountValue(el, target, decimals, suffix);
    });

    requestAnimationFrame(() => {
        bars.forEach(bar => { bar.style.transition = ''; });
        jcBars.forEach(bar => { bar.style.transition = ''; });
    });
}

function getStatusBadge(avg) {
    const scale = getScoreScale();
    const goodThreshold = scale > 10 ? 75 : scale * 0.75;
    if (avg >= goodThreshold) return '<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Baik</span>';
    if (avg >= KKM) return '<span class="badge badge-warning"><i class="fa-solid fa-exclamation-triangle"></i> Cukup</span>';
    return '<span class="badge badge-danger"><i class="fa-solid fa-circle-exclamation"></i> Perlu Bimbingan</span>';
}

function getStatusText(avg) {
    const scale = getScoreScale();
    const goodThreshold = scale > 10 ? 75 : scale * 0.75;
    if (avg >= goodThreshold) return "Baik";
    if (avg >= KKM) return "Cukup";
    return "Perlu Bimbingan";
}

function isBelowKKM(avg) { return avg < KKM && avg > 0; }

function buildMiniStatsHTML(subjects, averages, wali = '', showWaliBadge = true) {
    if (!subjects.length || !averages.length) return '';
    const maxVal = Math.max(...averages),
        minVal = Math.min(...averages);
    const maxIdx = averages.indexOf(maxVal),
        minIdx = averages.indexOf(minVal);
    const overall = (averages.reduce((a, b) => a + b, 0) / averages.length).toFixed(1);
    const waliBadge = wali && showWaliBadge ? `
            <div class="mini-wali" style="margin:0 0 12px; padding:8px 12px; border-radius:999px; background:rgba(255,255,255,0.06); border:1px solid rgba(255,255,255,0.12); color:var(--muted); font-size:0.8rem; display:inline-flex; align-items:center; gap:8px;">
                <i class="fa-solid fa-user-tie"></i><span>Wali Kelas: <strong style="color:var(--text);">${arabicTextHTML(wali)}</strong></span>
            </div>` : '';
    return `
            ${waliBadge}
            <div class="mini-stats">
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-trophy"></i> Tertinggi</div><div class="mini-value count-value" data-count-to="${maxVal}" data-decimals="1">0.0</div><div class="mini-sub">${arabicTextHTML(subjects[maxIdx])}</div></div>
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-arrow-trend-down"></i> Terendah</div><div class="mini-value count-value" data-count-to="${minVal}" data-decimals="1">0.0</div><div class="mini-sub">${arabicTextHTML(subjects[minIdx])}</div></div>
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-bullseye"></i> Rata-rata</div><div class="mini-value count-value" data-count-to="${overall}" data-decimals="1">0.0</div><div class="mini-sub">Keseluruhan</div></div>
            </div>`;
}

// Render Per Kelas
function renderKelasTable(subjects, averages, sheetName, studentCount = 0, sourceData = null) {
    const title = sheetName || currentSheetName || 'Kelas';
    const wali = waliKelas[sheetName || currentSheetName] || (sourceData?.asuhan) || '';
    kelasTableTitle.innerHTML = `Tabel Analisis — ${arabicTextHTML(title)}` + (studentCount ? ` (${studentCount} siswa)` : '') + (wali ? ` · Wali Kelas: ${arabicTextHTML(wali)}` : '');

    if (!subjects.length || !averages.length) {
        kelasQuickSummary.innerHTML = '';
        tableBodyKelas.innerHTML =
            `<tr><td colspan="4" style="padding:0;"><div class="empty-dropzone"><div class="empty-dropzone-icon"><i class="fa-solid fa-file-excel"></i></div><div class="empty-dropzone-title">Belum ada data</div><div class="empty-dropzone-text">Silakan upload file Excel nilai siswa untuk mulai melihat analisis performa akademik.</div><label for="file-input" class="empty-dropzone-btn"><i class="fa-solid fa-cloud-arrow-up"></i> Upload Data Sekarang</label></div></td></tr>`;
        kelasKetuntasanSummaryBox.innerHTML = '';
        kelasKetuntasanBox.innerHTML = '';
        kelasMiniStats.innerHTML = '';
        return;
    }

    const sorted = sortSubjectsByAverageDesc(subjects, averages);
    currentSubjectNames = sorted.subjects;
    currentAverages = sorted.averages;
    currentStudentCount = studentCount;
    // Status & dataset di hero sudah dihapus, jadi tidak perlu update elemen itu lagi

    kelasMiniStats.innerHTML = buildMiniStatsHTML(currentSubjectNames, currentAverages, wali, false);

    tableBodyKelas.innerHTML = '';
    currentSubjectNames.forEach((sub, idx) => {
        const avg = currentAverages[idx];
        const kkmClass = isBelowKKM(avg) ? 'below-kkm' : '';
        const row = tableBodyKelas.insertRow();
        row.className = kkmClass;
        row.insertCell(0).innerText = idx + 1;
        row.insertCell(1).innerHTML = `<span class="mapel-name">${arabicTextHTML(sub)}</span>`;
        row.insertCell(2).innerHTML = getProgressBarHTML(avg);
        row.insertCell(3).innerHTML = getStatusBadge(avg);
    });

    const ketuntasan = sourceData ? buildKetuntasanHTML(sourceData, null, false, null, true, true) : null;
    kelasKetuntasanSummaryBox.innerHTML = ketuntasan?.summary || '';
    kelasKetuntasanBox.innerHTML = ketuntasan?.students || '';
    const studentAverages = sourceData?.scores?.length ? computeStudentAverages(sourceData.scores) : [];
    if (studentAverages.length) {
        const tuntas = getTuntasCount(studentAverages);
        const total = studentAverages.length;
        const percent = (tuntas / total) * 100;
        kelasQuickSummary.innerHTML = `
            <i class="fa-solid fa-circle-check" aria-hidden="true"></i>
            <div class="kelas-quick-copy">
                <span>Ketuntasan kelas</span>
                <strong>${tuntas}/${total} siswa <span>${percent.toFixed(1)}%</span></strong>
            </div>
            <div class="kelas-quick-progress" role="img" aria-label="${percent.toFixed(1)}% tuntas">
                <span style="width:${percent.toFixed(1)}%"></span>
            </div>`;
    } else {
        kelasQuickSummary.innerHTML = '';
    }

    runCountUps(document.getElementById('pane-kelas'));
}

// Render Single Jenjang (dengan filter mapel)
function renderSingleJenjang(key, selectedSubjects) {
    let data = null;
    let title = '';

    if (groupedData[key]) {
        data = groupedData[key];
        title = key.replace('_', ' ');
    } else if (allSheetsData[key]) {
        data = allSheetsData[key];
        title = key;
    }

    if (!data || !data.subjects.length) {
        jenjangContainer.innerHTML = '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Tidak ada data</p>';
        return;
    }

    let subjectsToShow = data.subjects;
    let averagesToShow = data.averages;
    let selectedIndices = null; // null = pakai semua mapel (default)
    if (Array.isArray(selectedSubjects) && selectedSubjects.length === 0) {
        subjectsToShow = [];
        averagesToShow = [];
    } else if (selectedSubjects && selectedSubjects.length) {
        const indices = [];
        selectedSubjects.forEach(name => {
            const idx = data.subjects.indexOf(name);
            if (idx !== -1) indices.push(idx);
        });
        if (indices.length) {
            subjectsToShow = indices.map(i => data.subjects[i]);
            averagesToShow = indices.map(i => data.averages[i]);
            selectedIndices = indices;
        }
    }

    if (!subjectsToShow.length) {
        jenjangContainer.innerHTML = '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Tidak ada mapel dipilih.</p>';
        return;
    }

    const sorted = sortSubjectsByAverageDesc(subjectsToShow, averagesToShow);
    let tableRows = '';
    sorted.subjects.forEach((sub, i) => {
        const avg = sorted.averages[i];
        const kkmClass = isBelowKKM(avg) ? 'below-kkm' : '';
        tableRows +=
            `<tr class="${kkmClass}"><td>${i+1}</td><td><span class="mapel-name">${arabicTextHTML(sub)}</span></td><td>${getProgressBarHTML(avg)}</td><td>${getStatusBadge(avg)}</td></tr>`;
    });

    const sTitle = title + (data.studentCount ? ` (${data.studentCount} siswa)` : '');
    const wali = data?.asuhan || waliKelas[key] || '';
    const miniStatsHTML = buildMiniStatsHTML(sorted.subjects, sorted.averages, wali);
    const ketuntasanHTML = buildKetuntasanHTML(data, selectedIndices, false, null, false);

    jenjangContainer.innerHTML = `
            <div class="section-title" style="font-size:1.2rem; justify-content:space-between; flex-wrap:wrap;">
                <span style="display:flex; align-items:center; gap:10px;"><i class="fa-solid fa-table-list"></i> ${sTitle}</span>
                <div class="export-row" style="margin-top:0;">
                    <button id="exportPDFJenjang" class="btn btn-pdf btn-sm"><i class="fa-solid fa-file-pdf"></i> Export PDF</button>
                    <button id="exportCSVJenjang" class="btn btn-csv btn-sm"><i class="fa-solid fa-file-csv"></i> Export CSV</button>
                </div>
            </div>
            ${miniStatsHTML}
            <div class="table-wrap">
                <table><thead><tr><th>#</th><th>Mapel</th><th>Rata2</th><th>Status</th></tr></thead><tbody>${tableRows}</tbody></table>
            </div>
            ${ketuntasanHTML}`;

    document.getElementById('exportPDFJenjang').addEventListener('click', () => exportPDF(sorted.subjects, sorted
        .averages, title, jenjangContainer));
    document.getElementById('exportCSVJenjang').addEventListener('click', () => exportCSV(sorted.subjects, sorted
        .averages, title));

    runCountUps(jenjangContainer);
}

// Build jenjang dropdown options and populate mapel select
function buildJenjangDropdown() {
    jenjangSelect.innerHTML = '';
    mapelSelect.innerHTML = '';

    const groupedKeys = Object.keys(groupedData);
    const sheetKeys = Object.keys(allSheetsData);

    let options = [];
    if (groupedKeys.length) {
        options = sortJenjangKeys(groupedKeys);
    } else if (sheetKeys.length) {
        options = sheetKeys;
    }

    if (!options.length) {
        jenjangSelect.innerHTML = '<option value="">Tidak ada data</option>';
        jenjangDropdownBar.style.display = 'none';
        jenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Menunggu data komparatif...</p>';
        renderJenjangStudentSearch();
        return;
    }

    options.forEach(key => {
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = key.replace('_', ' ');
        jenjangSelect.appendChild(opt);
    });

    jenjangDropdownBar.style.display = 'flex';
    const firstKey = options[0];
    jenjangSelect.value = firstKey;
    populateMapelSelect(firstKey);
    renderSingleJenjang(firstKey, getSelectedMapels());
    renderJenjangStudentSearch();
}

// Populate mapel chip toggle based on jenjang key
function populateMapelSelect(key) {
    mapelSelect.innerHTML = '';
    let data = null;
    if (groupedData[key]) data = groupedData[key];
    else if (allSheetsData[key]) data = allSheetsData[key];
    if (!data || !data.subjects.length) return;

    data.subjects.forEach(sub => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'mapel-chip-btn active-chip';
        btn.dataset.value = sub;
        btn.innerHTML = `${arabicTextHTML(sub)} <span class="mapel-chip-x">&times;</span>`;
        btn.addEventListener('click', () => {
            btn.classList.toggle('active-chip');
            btn.innerHTML = btn.classList.contains('active-chip')
                ? `${arabicTextHTML(sub)} <span class="mapel-chip-x">&times;</span>`
                : arabicTextHTML(sub);
            const key2 = jenjangSelect.value;
            if (key2) renderSingleJenjang(key2, getSelectedMapels());
        });
        mapelSelect.appendChild(btn);
    });
}

function getSelectedMapels() {
    const selected = [];
    mapelSelect.querySelectorAll('.mapel-chip-btn.active-chip').forEach(btn => selected.push(btn.dataset.value));
    return selected;
}

// Ambil daftar nama sheet (kelas) yang tergabung dalam satu key jenjang,
// memakai logika pengelompokan yang sama persis dengan processAllSheets.
function getSheetsInJenjang(key) {
    if (!originalSheetNames.length) return [];
    const list = [];
    originalSheetNames.forEach(name => {
        const k = getClassGroupKey(name);
        if (k === key) list.push(name);
    });
    return list;
}

function renderJenjangStudentSearch() {
    const query = jenjangStudentSearchInput.value.trim().toLocaleLowerCase();
    const key = jenjangSelect.value;
    if (!query || !key) {
        jenjangStudentSearchResults.hidden = true;
        jenjangStudentSearchResults.innerHTML = '';
        return;
    }

    const sheets = getSheetsInJenjang(key);
    const datasets = sheets.length ? sheets.map(name => [name, allSheetsData[name]]) :
        (allSheetsData[key] ? [[key, allSheetsData[key]]] : []);
    const matches = [];

    datasets.forEach(([sheetName, data]) => {
        if (!data?.scores?.length) return;
        data.scores.forEach((scores, index) => {
            const name = data.studentNames?.[index] || '';
            if (!name.toLocaleLowerCase().includes(query)) return;
            const average = computeStudentAverages([scores])[0] || 0;
            const grades = (data.subjects || []).map((subject, subjectIndex) => ({
                subject,
                score: scores[subjectIndex]
            })).filter(grade => typeof grade.score === 'number' && Number.isFinite(grade.score));
            matches.push({ name, sheetName, average, grades });
        });
    });

    const resultItems = matches.map(student => `
        <article class="jenjang-student-result">
            <div class="jenjang-student-result-heading">
                <strong>${arabicTextHTML(student.name)}</strong>
                <span class="jenjang-student-result-meta">
                    <span>${escapeHTML(formatClassLabel(student.sheetName))}</span>
                    <span>Rerata ${student.average.toFixed(1)}</span>
                    <span class="${student.average >= KKM ? 'passed' : 'not-passed'}">${student.average >= KKM ? 'Tuntas' : 'Belum Tuntas'}</span>
                </span>
            </div>
            <div class="jenjang-student-result-grades">
                ${student.grades.map(grade => `
                    <span class="jenjang-student-grade ${grade.score >= KKM ? 'passed' : 'not-passed'}">
                        ${arabicTextHTML(grade.subject)}<strong>${grade.score.toFixed(1)}</strong>
                    </span>`).join('')}
            </div>
        </article>`).join('');

    jenjangStudentSearchResults.innerHTML = matches.length
        ? `<div class="jenjang-student-search-count">${matches.length} siswa ditemukan</div>${resultItems}`
        : '<div class="jenjang-student-search-empty">Nama siswa tidak ditemukan di jenjang ini.</div>';
    jenjangStudentSearchResults.hidden = false;
}

// Bangun dropdown jenjang khusus untuk tab Grafik Perjenjang
function buildGrafikJenjangDropdown() {
    grafikJenjangSelect.innerHTML = '';

    const groupedKeys = Object.keys(groupedData);
    const sheetKeys = Object.keys(allSheetsData);

    let options = [];
    if (groupedKeys.length) {
        options = sortJenjangKeys(groupedKeys);
    } else if (sheetKeys.length) {
        options = sheetKeys;
    }

    if (!options.length) {
        grafikJenjangSelect.innerHTML = '<option value="">Tidak ada data</option>';
        grafikPerjenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Menunggu data komparatif...</p>';
        return;
    }

    options.forEach(key => {
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = key.replace('_', ' ');
        grafikJenjangSelect.appendChild(opt);
    });

    grafikJenjangSelect.value = options[0];
    populateGrafikMapelSelect(options[0]);
    renderJenjangCompareChart(options[0], getSelectedGrafikMapels());
}

// Populate chip mapel khusus tab Grafik Perjenjang, persis sama polanya kayak populateMapelSelect
function populateGrafikMapelSelect(key) {
    grafikMapelSelect.innerHTML = '';
    let data = null;
    if (groupedData[key]) data = groupedData[key];
    else if (allSheetsData[key]) data = allSheetsData[key];
    if (!data || !data.subjects.length) return;

    data.subjects.forEach(sub => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'mapel-chip-btn active-chip';
        btn.dataset.value = sub;
        btn.innerHTML = `${arabicTextHTML(sub)} <span class="mapel-chip-x">&times;</span>`;
        btn.addEventListener('click', () => {
            btn.classList.toggle('active-chip');
            btn.innerHTML = btn.classList.contains('active-chip')
                ? `${arabicTextHTML(sub)} <span class="mapel-chip-x">&times;</span>`
                : arabicTextHTML(sub);
            const key2 = grafikJenjangSelect.value;
            if (key2) renderJenjangCompareChart(key2, getSelectedGrafikMapels());
        });
        grafikMapelSelect.appendChild(btn);
    });
}

function getSelectedGrafikMapels() {
    const selected = [];
    grafikMapelSelect.querySelectorAll('.mapel-chip-btn.active-chip').forEach(btn => selected.push(btn.dataset.value));
    return selected;
}

// Render chart batang: persentase siswa tuntas per kelas, dalam satu jenjang terpilih
// selectedSubjects: null = semua mapel; array kosong = tidak ada mapel dipilih.
function renderJenjangCompareChart(key, selectedSubjects) {
    if (!key) {
        grafikPerjenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Tidak ada data</p>';
        return;
    }

    let sheetList = getSheetsInJenjang(key);
    // Fallback: kalau tidak ada pengelompokan (mode single-sheet), anggap key itu satu kelas
    if (!sheetList.length && allSheetsData[key]) sheetList = [key];

    if (!sheetList.length) {
        grafikPerjenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Tidak ada data kelas untuk jenjang ini.</p>';
        return;
    }

    const hasFilter = Array.isArray(selectedSubjects);

    if (hasFilter) {
        // Cek dulu apakah ada minimal satu mapel terpilih yang cocok di salah satu kelas
        const anyMatch = sheetList.some(name => {
            const subs = (allSheetsData[name] && allSheetsData[name].subjects) || [];
            return selectedSubjects.some(s => subs.indexOf(s) !== -1);
        });
        if (!anyMatch) {
            grafikPerjenjangContainer.innerHTML =
                '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Tidak ada mapel dipilih.</p>';
            return;
        }
    }

    const bars = sheetList.map(name => {
        const d = allSheetsData[name];
        let scores = d && d.scores ? d.scores : [];
        // Cocokkan nama mapel terpilih ke kolom mapel di masing-masing kelas (urutan/mapel bisa beda tiap kelas)
        if (hasFilter && d && d.subjects && d.subjects.length) {
            const indices = [];
            selectedSubjects.forEach(sub => {
                const idx = d.subjects.indexOf(sub);
                if (idx !== -1) indices.push(idx);
            });
            scores = indices.length ? scores.map(student => indices.map(i => student[i])) : [];
        }
        const studentAverages = computeStudentAverages(scores);
        const total = studentAverages.length;
        const tuntas = getTuntasCount(studentAverages);
        const pct = total > 0 ? (tuntas / total) * 100 : 0;
        return { name, pct, tuntas, total };
    });

    let barsHTML = '';
    bars.forEach(b => {
        const pctStr = b.pct.toFixed(1);
        const targetH = b.total ? pctStr : 0;
        barsHTML += `
            <div class="jc-bar-col" title="${b.name}: ${b.tuntas}/${b.total} siswa tuntas">
                <div class="jc-bar-value count-value" data-count-to="${targetH}" data-decimals="1" data-suffix="%">0.0%</div>
                <div class="jc-bar-track">
                    <div class="jc-bar-fill" data-target-height="${targetH}%" style="height:0%"></div>
                </div>
                <div class="jc-bar-label">${arabicTextHTML(formatClassLabel(b.name))}</div>
            </div>`;
    });

    // Ketuntasan per-kelas (bukan digabung), satu blok per kelas di jenjang ini, ikut mapel yang difilter
    let ketuntasanHTML = '';
    sheetList.forEach(name => {
        const d = allSheetsData[name];
        if (!d) return;
        let selectedIndices = null;
        if (hasFilter && d.subjects && d.subjects.length) {
            const indices = [];
            selectedSubjects.forEach(sub => {
                const idx = d.subjects.indexOf(sub);
                if (idx !== -1) indices.push(idx);
            });
            if (indices.length) selectedIndices = indices;
        }
        const block = buildKetuntasanHTML(d, selectedIndices, false, name);
        if (block) {
            ketuntasanHTML += `<div class="ketuntasan-perkelas-block">${block}</div>`;
        }
    });

    grafikPerjenjangContainer.innerHTML = `
            <div class="jenjang-compare-chart-wrap">
                <div class="jenjang-compare-chart">${barsHTML}</div>
            </div>
            ${ketuntasanHTML ? `<div class="ketuntasan-perkelas-grid">${ketuntasanHTML}</div>` : ''}`;

    runCountUps(grafikPerjenjangContainer);
}

// Bangun blok Ketuntasan untuk satu jenjang (ditempel langsung di layout Per Jenjang)
function escapeHTML(value) {
    return String(value).replace(/[&<>"']/g, char => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    })[char]);
}

function arabicTextHTML(value) {
    const text = String(value);
    return /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(text)
        ? `<span class="arabic-text" lang="ar">${escapeHTML(text)}</span>`
        : escapeHTML(text);
}

function buildKetuntasanHTML(data, selectedIndices, showChart, titleSuffix, showFailedStudents = true, splitPanels = false) {
    const allScores = data.scores || [];
    // Jika ada mapel yang dipilih (selectedIndices), hitung ketuntasan hanya dari kolom mapel itu.
    // Jika tidak (null/kosong), pakai semua mapel seperti semula.
    const subjectIndices = selectedIndices && selectedIndices.length ?
        selectedIndices : (data.subjects || []).map((_, index) => index);
    const scores = (selectedIndices && selectedIndices.length && allScores.length) ?
        allScores.map(student => selectedIndices.map(i => student[i])) :
        allScores;

    let totalStudents = data.studentCount || 0;
    let tuntas = 0,
        tidakTuntas = 0,
        persentaseTuntas = 0;
    const studentAverages = computeStudentAverages(scores);
    const rankedStudents = studentAverages.map((average, index) => ({
        name: data.studentNames?.[index] || 'Nama tidak terbaca',
        average
    })).sort((a, b) => b.average - a.average || a.name.localeCompare(b.name, 'id'));

    if (scores.length) {
        totalStudents = studentAverages.length;
        tuntas = getTuntasCount(studentAverages);
        tidakTuntas = totalStudents - tuntas;
        persentaseTuntas = totalStudents > 0 ? (tuntas / totalStudents) * 100 : 0;
    }

    if (!totalStudents) return '';

    const avgs = (selectedIndices && selectedIndices.length) ?
        selectedIndices.map(i => (data.averages || [])[i]).filter(v => typeof v === 'number') :
        (data.averages || []);
    const overall = avgs.length ? (avgs.reduce((a, b) => a + b, 0) / avgs.length).toFixed(1) : '-';
    const tuntasPct = persentaseTuntas.toFixed(1);
    const tidakTuntasPct = (100 - persentaseTuntas).toFixed(1);
    const renderStudentList = students => students.length
        ? `<ol>${students.map(student => `
            <li>
                <span>${arabicTextHTML(student.name)}</span>
                <strong>${student.average.toFixed(1)}</strong>
            </li>`).join('')}</ol>`
        : '<p class="rekap-student-empty">Data nama siswa tidak tersedia.</p>';
    const tuntasStudents = rankedStudents.filter(student => student.average >= KKM);
    const tidakTuntasStudents = rankedStudents.filter(student => student.average < KKM);
    const failedSubjects = showFailedStudents ? subjectIndices.map((subjectIndex, scoreIndex) => ({
        name: (data.subjects || [])[subjectIndex],
        passed: scores.map((student, index) => ({
            studentIndex: index,
            name: (data.studentNames || [])[index],
            score: student[scoreIndex]
        })).filter(student => typeof student.score === 'number' && student.score >= KKM),
        students: scores.map((student, index) => ({
            studentIndex: index,
            name: (data.studentNames || [])[index],
            score: student[scoreIndex]
        })).filter(student => typeof student.score === 'number' && student.score < KKM)
    })) : [];
    const failedStudentCount = totalStudents;
    const failedStudentsHTML = failedSubjects.length ? `
            <aside class="failed-students">
                <h3><span class="failed-students-title"><i class="fa-solid fa-users"></i> Daftar Siswa</span><span>${failedStudentCount} siswa</span></h3>
                <div class="failed-student-subject-controls">
                    <label class="failed-student-search">
                        <i class="fa-solid fa-magnifying-glass" aria-hidden="true"></i>
                        <input type="search" class="failed-student-search-input" placeholder="Cari nama siswa..." aria-label="Cari nama siswa">
                    </label>
                    <span class="failed-student-subject-label"><i class="fa-solid fa-book"></i> Mata Pelajaran</span>
                    <div class="failed-student-subject-filter" aria-label="Pilih mapel">
                        ${failedSubjects.map((subject, index) => `
                        <button type="button" class="mapel-chip-btn failed-subject-toggle active-chip" data-subject-index="${index}" aria-pressed="true">
                            ${arabicTextHTML(subject.name)} <span class="mapel-chip-x" aria-hidden="true">&times;</span>
                        </button>`).join('')}
                    </div>
                    <button type="button" class="btn btn-sm btn-primary" data-failed-subject-action="all">Pilih Semua</button>
                    <button type="button" class="btn btn-sm" data-failed-subject-action="none">Hapus Semua</button>
                </div>
                <div class="failed-student-subjects">
                    ${failedSubjects.map((subject, index) => `
                    <section class="failed-student-subject" data-subject-index="${index}">
                        <h4><span class="failed-student-subject-name">${arabicTextHTML(subject.name)}</span></h4>
                        <div class="student-status-columns">
                            <div class="student-status-list is-passed">
                                <h5><i class="fa-solid fa-circle-check"></i> Siswa Tuntas <span>${subject.passed.length}</span></h5>
                                <ol>${subject.passed.map(student =>
                                    `<li><span class="failed-student-name" title="${escapeHTML(student.name || 'Nama tidak terbaca')}">${arabicTextHTML(student.name || 'Nama tidak terbaca')}</span><strong>${student.score.toFixed(1)}</strong></li>`
                                ).join('') || '<li class="student-list-empty">Tidak ada</li>'}</ol>
                            </div>
                            <div class="student-status-list is-failed">
                                <h5><i class="fa-solid fa-circle-xmark"></i> Siswa Tidak Tuntas <span>${subject.students.length}</span></h5>
                                <ol>${subject.students.map(student =>
                                    `<li><span class="failed-student-name" title="${escapeHTML(student.name || 'Nama tidak terbaca')}">${arabicTextHTML(student.name || 'Nama tidak terbaca')}</span><strong>${student.score.toFixed(1)}</strong></li>`
                                ).join('') || '<li class="student-list-empty">Tidak ada</li>'}</ol>
                            </div>
                        </div>
                    </section>`).join('')}
                </div>
                <p class="failed-student-no-results" hidden>Nama siswa tidak ditemukan.</p>
            </aside>` : '';
    const ketuntasanTitleHTML = `<div class="section-title" style="font-size:1.1rem; margin-top:24px;">Ketuntasan${titleSuffix ? ' ' + titleSuffix : ''}</div>`;
    const summaryHTML = `
            ${ketuntasanTitleHTML}
            <div class="rekap-info-row" style="margin-bottom:12px;">
                <span class="rekap-info-chip"><i class="fa-solid fa-users"></i> Siswa: ${totalStudents}</span>
                <span class="rekap-info-chip"><i class="fa-solid fa-circle-check" style="color:var(--success);"></i> Tuntas: ${tuntas}</span>
                <span class="rekap-info-chip"><i class="fa-solid fa-circle-xmark" style="color:var(--danger);"></i> Tidak Tuntas: ${tidakTuntas}</span>
            </div>
            ${showChart ? `
            <div class="ketuntasan-chart-row">
                <div class="ketuntasan-chart-title"><i class="fa-solid fa-chart-simple"></i> Grafik Persentase Ketuntasan</div>
                <div class="ketuntasan-bar-track">
                    <div class="ketuntasan-bar-fill" data-target-width="${tuntasPct}%" style="width:${tuntasPct}%;">${tuntasPct}%</div>
                </div>
                <div class="ketuntasan-legend">
                    <div class="ketuntasan-legend-item"><span class="legend-dot" style="background:var(--success);"></span> Tuntas: <b>${tuntas}</b> siswa (${tuntasPct}%)</div>
                    <div class="ketuntasan-legend-item"><span class="legend-dot" style="background:var(--danger);"></span> Tidak Tuntas: <b>${tidakTuntas}</b> siswa (${tidakTuntasPct}%)</div>
                </div>
            </div>` : ''}
            <div class="rekap-minmax">
                <section class="rekap-student-group is-passed">
                    <div class="rekap-mm-card">
                        <button type="button" class="rekap-mm-card-button" data-rekap-student-toggle aria-expanded="false">
                            <span class="mm-label" style="color:var(--success);"><i class="fa-solid fa-check-circle"></i> Tuntas</span>
                            <span class="mm-value count-value" data-count-to="${persentaseTuntas}" data-decimals="1" data-suffix="%">0%</span>
                            <span class="mm-subject">${tuntas} siswa</span>
                            <span class="rekap-student-hint"><i class="fa-solid fa-chevron-down"></i> Lihat nama</span>
                        </button>
                    </div>
                    <div class="rekap-student-list" hidden>${renderStudentList(tuntasStudents)}</div>
                </section>
                <section class="rekap-student-group is-failed">
                    <div class="rekap-mm-card">
                        <button type="button" class="rekap-mm-card-button" data-rekap-student-toggle aria-expanded="false">
                            <span class="mm-label" style="color:var(--danger);"><i class="fa-solid fa-times-circle"></i> Tidak Tuntas</span>
                            <span class="mm-value count-value" data-count-to="${100 - persentaseTuntas}" data-decimals="1" data-suffix="%">0%</span>
                            <span class="mm-subject">${tidakTuntas} siswa</span>
                            <span class="rekap-student-hint"><i class="fa-solid fa-chevron-down"></i> Lihat nama</span>
                        </button>
                    </div>
                    <div class="rekap-student-list" hidden>${renderStudentList(tidakTuntasStudents)}</div>
                </section>
            </div>
            <div class="ketuntasan-average"><span>Rerata: ${overall}</span><span>KKM: ${KKM}</span></div>`;

    if (splitPanels) return {
        summary: `<div class="ketuntasan-summary">${summaryHTML}</div>`,
        students: failedStudentsHTML
    };

    return `
            <div class="ketuntasan-layout${failedSubjects.length ? ' has-failed-students' : ''}">
                <div class="ketuntasan-summary">${summaryHTML}</div>
                ${failedStudentsHTML}
            </div>`;
}

function filterBySheet(sheetName) {
    sheetSelect.value = sheetName;
    if (allSheetsData[sheetName]?.subjects?.length) {
        const d = allSheetsData[sheetName];
        currentSubjectNames = d.subjects;
        currentAverages = d.averages;
        currentStudentCount = d.studentCount || 0;
        currentSheetName = sheetName;
        renderKelasTable(currentSubjectNames, currentAverages, sheetName, currentStudentCount, d);
    }
    saveUIState();
    document.querySelector('.main-tab[data-pane="kelas"]').click();
}

// Hitung ringkasan (total siswa, rata2, ketuntasan, dst) dari sekumpulan sheet — dipakai buat
// kuartal aktif maupun buat bandingin antar kuartal (tren), makanya gak nempel ke variabel global.
function computeOverviewStats(sheetsData, sheetNames) {
    let totalSiswa = 0;
    let allStudentAverages = [];
    let belowKKMCount = 0;
    const classStats = [];

    (sheetNames || []).forEach(name => {
        const d = sheetsData ? sheetsData[name] : null;
        if (!d) return;
        totalSiswa += d.studentCount || 0;

        const studentAverages = computeStudentAverages(d.scores || []);
        allStudentAverages = allStudentAverages.concat(studentAverages);

        (d.averages || []).forEach(avg => { if (isBelowKKM(avg)) belowKKMCount++; });

        const total = studentAverages.length;
        if (total > 0) {
            const tuntas = getTuntasCount(studentAverages);
            const pct = (tuntas / total) * 100;
            const avgOverall = d.averages && d.averages.length ?
                (d.averages.reduce((a, b) => a + b, 0) / d.averages.length) : 0;
            classStats.push({ name, pct, avg: avgOverall });
        }
    });

    const totalAll = allStudentAverages.length;
    const totalTuntas = getTuntasCount(allStudentAverages);
    const pctSekolah = totalAll > 0 ? (totalTuntas / totalAll) * 100 : 0;
    const rataSekolah = totalAll > 0 ? (allStudentAverages.reduce((a, b) => a + b, 0) / totalAll) : 0;

    classStats.sort((a, b) => b.pct - a.pct);

    return { totalSiswa, pctSekolah, rataSekolah, belowKKMCount, classStats, hasData: totalAll > 0 };
}

// Satukan alias MID/MDTA dan MTSD/MDTW untuk ringkasan seluruh sekolah.
function splitSheetsByJenjang(sheetNames) {
    const mid = [];
    const mdtw = [];
    (sheetNames || []).forEach(name => {
        if (getClassType(name) === 'MDTW') mdtw.push(name); else mid.push(name);
    });
    return { mid, mdtw };
}

const JENJANG_META = {
    MID:  { label: 'MDTA', subtitle: 'Madrasah Ibtidaiyah Diniyah (MID)',  icon: 'fa-seedling',       accentClass: 'accent-mid' },
    MDTW: { label: 'MDTW', subtitle: 'Madrasah Tsanawiyah Diniyah (MTsD)', icon: 'fa-graduation-cap', accentClass: 'accent-mtsd' }
};

// Bangun satu section penuh (kartu ringkasan + highlight kelas + tren antar kuartal)
// buat satu jenjang tertentu. Semua angka di dalamnya cuma dihitung dari sheet-sheet
// yang emang masuk jenjang itu, gak ketimpuk sama jenjang satunya.
function buildJenjangSectionHTML(jenjangKey, sheetNames) {
    const meta = JENJANG_META[jenjangKey];
    const stats = computeOverviewStats(allSheetsData, sheetNames);

    let bodyHTML;
    if (!stats.hasData) {
        bodyHTML = `
            <div style="text-align:center; padding:24px 12px; color:var(--muted);">
                <div style="font-size:1.4rem; margin-bottom:6px; opacity:0.5;"><i class="fa-solid ${meta.icon}"></i></div>
                <div style="font-size:0.85rem;">Belum ada data kelas untuk jenjang ${meta.label}.</div>
            </div>`;
    } else {
        const statsGridHTML = `
            <div class="overview-stats-grid">
                <div class="overview-stat-card">
                    <div class="mini-label"><i class="fa-solid fa-users"></i> Total Siswa</div>
                    <div class="mini-value count-value" data-count-to="${stats.totalSiswa}" data-decimals="0">0</div>
                </div>
                <div class="overview-stat-card">
                    <div class="mini-label"><i class="fa-solid fa-bullseye"></i> Rata-rata</div>
                    <div class="mini-value count-value" data-count-to="${stats.rataSekolah.toFixed(1)}" data-decimals="1">0.0</div>
                </div>
                <div class="overview-stat-card">
                    <div class="mini-label"><i class="fa-solid fa-circle-check"></i> Ketuntasan</div>
                    <div class="mini-value count-value" data-count-to="${stats.pctSekolah.toFixed(1)}" data-decimals="1" data-suffix="%">0.0%</div>
                </div>
                <div class="overview-stat-card">
                    <div class="mini-label"><i class="fa-solid fa-triangle-exclamation"></i> Mapel di Bawah KKM</div>
                    <div class="mini-value count-value" data-count-to="${stats.belowKKMCount}" data-decimals="0">0</div>
                </div>
            </div>`;

        // Kelas paling unggul / perlu perhatian, dihitung dalam lingkup jenjang ini aja
        const best = stats.classStats[0];
        const worst = stats.classStats.length > 1 ? stats.classStats[stats.classStats.length - 1] : null;
        let highlightHTML = '';
        if (best) {
            highlightHTML += `
                <div class="overview-highlight-card good">
                    <div class="mini-label"><i class="fa-solid fa-trophy"></i> Kelas Paling Unggul</div>
                    <div class="overview-highlight-name">${arabicTextHTML(best.name)}</div>
                    <div class="overview-highlight-sub">Ketuntasan ${best.pct.toFixed(1)}% &middot; Rata-rata ${best.avg.toFixed(1)}</div>
                </div>`;
        }
        if (worst) {
            highlightHTML += `
                <div class="overview-highlight-card bad">
                    <div class="mini-label"><i class="fa-solid fa-arrow-trend-down"></i> Perlu Perhatian</div>
                    <div class="overview-highlight-name">${arabicTextHTML(worst.name)}</div>
                    <div class="overview-highlight-sub">Ketuntasan ${worst.pct.toFixed(1)}% &middot; Rata-rata ${worst.avg.toFixed(1)}</div>
                </div>`;
        }

        // Tren rata-rata antar Kuartal I/II/III buat jenjang ini, ditarik dari quartalCache
        const qKeys = ['Kuartal_I', 'Kuartal_II', 'Kuartal_III'];
        const trendBars = qKeys.map(q => {
            const cache = quartalCache[q];
            const names = cache ? splitSheetsByJenjang(cache.originalSheetNames || [])[jenjangKey === 'MDTW' ? 'mdtw' : 'mid'] : [];
            const s = cache ? computeOverviewStats(cache.allSheetsData || {}, names) : null;
            return { label: QUARTAL_LABELS[q] || q, rata: s && s.hasData ? s.rataSekolah : 0, hasData: !!(s && s.hasData) };
        });
        let trendHTML = '';
        if (trendBars.some(b => b.hasData)) {
            let trendBarsHTML = '';
            trendBars.forEach(b => {
                const targetH = b.hasData ? Math.min(100, b.rata).toFixed(1) : 0;
                trendBarsHTML += `
                    <div class="jc-bar-col" title="${b.label}: ${b.hasData ? b.rata.toFixed(1) : 'belum ada data'}">
                        <div class="jc-bar-value count-value" data-count-to="${b.hasData ? b.rata.toFixed(1) : 0}" data-decimals="1">0.0</div>
                        <div class="jc-bar-track">
                            <div class="jc-bar-fill" data-target-height="${targetH}%" style="height:0%"></div>
                        </div>
                        <div class="jc-bar-label">${b.label.replace('_', ' ')}</div>
                    </div>`;
            });
            trendHTML = `
                <div class="section-title" style="font-size:0.95rem; margin-top:20px;"><i class="fa-solid fa-chart-line"></i> Tren Rata-rata Antar Kuartal</div>
                <div class="jenjang-compare-chart-wrap">
                    <div class="jenjang-compare-chart">${trendBarsHTML}</div>
                </div>`;
        }

        bodyHTML = `
            ${statsGridHTML}
            ${highlightHTML ? `<div class="overview-highlight-grid" style="margin-top:14px;">${highlightHTML}</div>` : ''}
            ${trendHTML}`;
    }

    return `
        <div class="overview-jenjang-section ${meta.accentClass}">
            <div class="overview-jenjang-heading">
                <span class="overview-jenjang-icon"><i class="fa-solid ${meta.icon}"></i></span>
                <span class="overview-jenjang-titles">
                    <span class="overview-jenjang-title">${meta.label}</span>
                    <span class="overview-jenjang-subtitle">${meta.subtitle}</span>
                </span>
            </div>
            ${bodyHTML}
        </div>`;
}

// Ringkasan sekolah dipisah per jenjang MDTA/MID dan MDTW/MTsD.
function renderOverview() {
    const overviewContent = document.getElementById('overviewContent');
    if (!overviewContent) return;

    if (!originalSheetNames.length) {
        overviewContent.innerHTML = `
            <div style="text-align:center; padding:48px 20px; color:var(--muted);">
                <div style="font-size:2rem; margin-bottom:10px; opacity:0.6;"><i class="fa-solid fa-cloud-arrow-up"></i></div>
                <div style="font-weight:700; color:var(--text-soft); margin-bottom:4px;">Belum ada data nilai</div>
                <div style="font-size:0.85rem;">Tekan tombol <b>Upload Data</b> di atas untuk memuat file Excel nilai ujian.</div>
            </div>`;
        return;
    }

    const { mid, mdtw } = splitSheetsByJenjang(originalSheetNames);

    overviewContent.innerHTML = `
            ${buildJenjangSectionHTML('MID', mid)}
            ${buildJenjangSectionHTML('MDTW', mdtw)}`;

    runCountUps(overviewContent);
}

function refreshAllViews() {
    if (!currentSubjectNames.length) return;
    const src = allSheetsData[currentSheetName] || groupedData[currentSheetName] || null;
    renderKelasTable(currentSubjectNames, currentAverages, currentSheetName, currentStudentCount, src);
    buildJenjangDropdown();
    buildGrafikJenjangDropdown();
    renderOverview();
}

// Draw one line of (correctly-shaped) text to a PNG using the browser's own canvas text engine.
// A single fillText() call shapes Arabic ligatures correctly — html2canvas's own text drawing
// does NOT, since it draws glyph-by-glyph, which is what breaks Arabic letter joining.
function textToImageDataURL(text, { fontSize, color, weight, fontFamily }) {
    const scale = 4; // supersample for crisp output at the 2x export scale
    const probe = document.createElement('canvas').getContext('2d');
    probe.font = `${weight} ${fontSize * scale}px ${fontFamily}`;
    const textWidth = probe.measureText(text).width;
    const w = Math.ceil(textWidth) + Math.ceil(fontSize * scale * 0.3);
    const h = Math.ceil(fontSize * scale * 1.5);
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.font = `${weight} ${fontSize * scale}px ${fontFamily}`;
    ctx.fillStyle = color;
    ctx.direction = 'rtl';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, w - Math.ceil(fontSize * scale * 0.15), h / 2);
    return { dataUrl: canvas.toDataURL('image/png'), widthPx: w / scale, heightPx: h / scale };
}

// Replace every Arabic (Amiri-font) text leaf inside `container` with a rasterized
// so html2canvas just copies pixels instead of re-drawing (and breaking) the Arabic script.
function rasterizeArabicText(container) {
    container.querySelectorAll('*').forEach(el => {
        if (el.children.length > 0) return; // only text leaves
        const text = el.textContent;
        if (!text || !text.trim()) return;
        const cs = getComputedStyle(el);
        if (!cs.fontFamily || !cs.fontFamily.toLowerCase().includes('amiri')) return;

        const fontSize = parseFloat(cs.fontSize) || 16;
        const { dataUrl, heightPx } = textToImageDataURL(text.trim(), {
            fontSize,
            color: cs.color || '#000000',
            weight: cs.fontWeight || '700',
            fontFamily: cs.fontFamily
        });

        const img = document.createElement('img');
        img.src = dataUrl;
        img.style.height = heightPx + 'px';
        img.style.width = 'auto';
        img.style.display = 'inline-block';
        img.style.verticalAlign = 'middle';
        el.textContent = '';
        el.appendChild(img);
    });
}

// Export Functions
// sourceEl: the live .section-container element to snapshot (so the PDF looks exactly like the app card)
async function exportPDF(subjects, averages, title, sourceEl, reportLabel, extraEls = []) {
    if (!subjects.length) { alert("Tidak ada data untuk diekspor."); return; }
    const btns = document.querySelectorAll('.btn-pdf');
    btns.forEach(b => b.disabled = true);
    try {
        await ensurePDFLibs();
    } catch (err) {
        alert("Gagal memuat modul export PDF. Cek koneksi internet.");
        btns.forEach(b => b.disabled = false);
        return;
    }

    try {
        const cardEl = sourceEl || document.querySelector('#pane-kelas .section-container');
        if (!cardEl) { alert("Tidak ada data untuk diekspor."); return; }

        // Clone the live card so we can strip buttons / lock-in final numbers without touching the real UI
        const clone = cardEl.cloneNode(true);
        clone.classList.add('pdf-export-clone');
        extraEls.forEach(extraEl => {
            const extraClone = extraEl.cloneNode(true);
            while (extraClone.firstChild) clone.appendChild(extraClone.firstChild);
        });
        clone.querySelectorAll('.export-row').forEach(row => row.remove());
        clone.querySelectorAll('.failed-students').forEach(list => list.remove());
        clone.style.margin = '0';
        clone.style.width = '760px';
        clone.style.maxWidth = '760px';

        // Freeze any count-up animation at its final value + full progress-bar width
        clone.querySelectorAll('.count-value[data-count-to]').forEach(el => {
            const target = parseFloat(el.dataset.countTo);
            const decimals = el.dataset.decimals !== undefined ? parseInt(el.dataset.decimals, 10) : 2;
            const suffix = el.dataset.suffix || '';
            if (!isNaN(target)) el.textContent = target.toFixed(decimals) + suffix;
        });
        clone.querySelectorAll('.progress-bar-fill[data-target-width]').forEach(bar => {
            bar.style.transition = 'none';
            bar.style.width = bar.dataset.targetWidth;
        });

        // Header (report title + date) styled to match the app's own type/colors
        const pageBg = getComputedStyle(document.body).backgroundColor || '#14110C';
        const textColor = getComputedStyle(document.body).color || '#F3EEE2';
        const header = document.createElement('div');
        header.style.cssText =
            `font-family:'Space Grotesk',sans-serif;color:${textColor};margin-bottom:16px;`;
        header.innerHTML = `
                <div style="font-weight:800;font-size:20px;">${reportLabel || 'Laporan Rata-rata'} — ${title}</div>
                <div style="opacity:0.6;font-size:12px;margin-top:4px;">Madrasah Diniyah "Miftahul Ulum" · ${new Date().toLocaleDateString('id-ID')}</div>`;

        const wrapper = document.createElement('div');
        wrapper.style.cssText = `display:inline-block;padding:24px;background:${pageBg};`;
        wrapper.appendChild(header);
        wrapper.appendChild(clone);

        pdfContainer.innerHTML = '';
        pdfContainer.appendChild(wrapper);

        // Wait for webfonts (Amiri etc.) to finish loading, THEN rasterize the Arabic text —
        // this is what fixes the broken letter-joining, since it no longer needs html2canvas
        // to draw the Arabic script itself. (Must happen after the clone is attached to the
        // document, otherwise getComputedStyle can't resolve the CSS font-family rules.)
        if (document.fonts && document.fonts.ready) { await document.fonts.ready; }
        rasterizeArabicText(clone);

        // Let images/layout settle before rasterizing the whole card
        await new Promise(r => setTimeout(r, 80));

        const canvas = await html2canvas(wrapper, {
            backgroundColor: pageBg,
            scale: 2,
            useCORS: true,
            allowTaint: true
        });

        // Size the PDF page exactly to the captured card — no extra blank/black canvas around it
        const pxToMm = px => px * 0.264583;
        const contentWmm = pxToMm(canvas.width / 2);
        const contentHmm = pxToMm(canvas.height / 2);

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: contentWmm >= contentHmm ? 'landscape' : 'portrait',
            unit: 'mm',
            format: [contentWmm, contentHmm]
        });
        doc.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, contentWmm, contentHmm);

        const safe = title.replace(/[^a-zA-Z0-9]/g, '_');
        doc.save(`Laporan_${safe}.pdf`);
    } finally {
        btns.forEach(b => b.disabled = false);
    }
}

function exportCSV(subjects, averages, title) {
    if (!subjects.length) { alert("Tidak ada data."); return; }
    const sorted = sortSubjectsByAverageDesc(subjects, averages);
    let csv = '\uFEFF#;Mata Pelajaran;Rata-rata;Status\n';
    sorted.subjects.forEach((sub, idx) => {
        const avg = sorted.averages[idx];
        csv += `${idx+1};"${sub}";${avg.toFixed(1)};${getStatusText(avg)}\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${title.replace(/[^a-zA-Z0-9]/g,'_')}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

function exportOverviewCSV() {
    if (!originalSheetNames.length) { alert("Tidak ada data."); return; }
    const { mid, mdtw } = splitSheetsByJenjang(originalSheetNames);
    let csv = '\uFEFF';
    [{ key: 'MID', names: mid }, { key: 'MDTW', names: mdtw }].forEach(({ key, names }) => {
        const meta = JENJANG_META[key];
        const stats = computeOverviewStats(allSheetsData, names);
        csv += `Ringkasan ${meta.label} (${meta.subtitle})\n`;
        if (!stats.hasData) {
            csv += 'Belum ada data untuk jenjang ini\n\n';
            return;
        }
        csv += `Total Siswa;${stats.totalSiswa}\n`;
        csv += `Rata-rata;${stats.rataSekolah.toFixed(1)}\n`;
        csv += `Ketuntasan;${stats.pctSekolah.toFixed(1)}%\n`;
        csv += `Mapel di Bawah KKM;${stats.belowKKMCount}\n`;
        csv += 'Kelas;Ketuntasan (%);Rata-rata\n';
        stats.classStats.forEach(c => {
            csv += `"${c.name}";${c.pct.toFixed(1)};${c.avg.toFixed(1)}\n`;
        });
        csv += '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Ringkasan_Sekolah_${QUARTAL_LABELS[currentQuartal].replace(/\s+/g,'_')}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

// Data Management
function processAllSheets() {
    if (!workbookData) return;
    const waliSheet = workbookData.SheetNames.find(name => normalizeCellText(name).toUpperCase() === 'DATA WALI KELAS');
    const waliMapping = waliSheet ? readWaliKelasSheet(workbookData.Sheets[waliSheet]) : {};
    const sheetNames = workbookData.SheetNames.filter(name => name !== waliSheet);
    if (!sheetNames.length) return;

    waliKelas = {};
    originalSheetNames = [...sheetNames];
    groupedData = {};
    allSheetsData = {};

    for (const name of sheetNames) {
        const ws = workbookData.Sheets[name];
        if (!ws) { allSheetsData[name] = { subjects: [], averages: [], studentCount: 0, scores: [], studentNames: [], studentMeta: [] }; continue; }
        const { subjects, scores, studentNames, studentMeta, studentCount, asuhan } = detectHeaderAndData(ws);

        // Pisah per kelas jika ada kolom KELAS dengan lebih dari satu nilai unik di dalam satu sheet
        const kelasValues = (studentMeta || []).map(m => m.kelas).filter(v => v && String(v).trim());
        const uniqueKelas = [...new Set(kelasValues)];

        if (uniqueKelas.length > 1 && kelasValues.length === (studentNames || []).length) {
            uniqueKelas.forEach(kls => {
                const indices = [];
                studentMeta.forEach((m, i) => { if (m.kelas === kls) indices.push(i); });
                const klsScores = indices.map(i => scores[i]);
                const klsName = `${name} — ${kls}`;
                originalSheetNames.push(klsName);
                allSheetsData[klsName] = {
                    subjects,
                    averages: computeAverages(subjects, klsScores),
                    studentCount: klsScores.length,
                    scores: klsScores,
                    studentNames: indices.map(i => studentNames[i]),
                    studentMeta: indices.map(i => studentMeta[i]),
                    asuhan
                };
                waliKelas[klsName] = findWaliKelas(klsName, { studentMeta: indices.map(i => studentMeta[i]) }, waliMapping) || asuhan;
            });
        } else {
            const avgs = computeAverages(subjects, scores);
            allSheetsData[name] = { subjects, averages: avgs, studentCount, scores: scores || [], studentNames, studentMeta, asuhan };
            waliKelas[name] = findWaliKelas(name, { studentMeta }, waliMapping) || asuhan;
        }
    }

    const groupMap = {};
    for (const name of sheetNames) {
        const key = getClassGroupKey(name);
        if (!groupMap[key]) groupMap[key] = [];
        groupMap[key].push(name);
    }
    for (const [key, sheetList] of Object.entries(groupMap)) {
        groupedData[key] = mergeSheetData(sheetList);
    }

    const uiState = getUIState();
    let targetSheet = uiState.sheet;

    sheetSelect.innerHTML = '';
    originalSheetNames.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        sheetSelect.appendChild(opt);
    });
    filterBar.style.display = 'flex';
    currentSheetName = originalSheetNames.includes(targetSheet) ? targetSheet : originalSheetNames[0];
    sheetSelect.value = currentSheetName;

    let found = false;
    if (targetSheet) {
        if (allSheetsData[targetSheet] && allSheetsData[targetSheet].subjects && allSheetsData[targetSheet]
            .subjects.length) {
            currentSubjectNames = allSheetsData[targetSheet].subjects;
            currentAverages = allSheetsData[targetSheet].averages;
            currentStudentCount = allSheetsData[targetSheet].studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, targetSheet, currentStudentCount, allSheetsData[targetSheet]);
            found = true;
        } else if (groupedData[targetSheet] && groupedData[targetSheet].subjects && groupedData[targetSheet]
            .subjects.length) {
            currentSubjectNames = groupedData[targetSheet].subjects;
            currentAverages = groupedData[targetSheet].averages;
            currentStudentCount = groupedData[targetSheet].studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, targetSheet.replace('_', ' '),
            currentStudentCount, groupedData[targetSheet]);
            found = true;
        }
    }

    if (!found) {
        const first = originalSheetNames[0];
        const fd = allSheetsData[first];
        if (fd?.subjects.length) {
            currentSubjectNames = fd.subjects;
            currentAverages = fd.averages;
            currentStudentCount = fd.studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, first, currentStudentCount, fd);
        } else { renderKelasTable([], [], ''); }
    }

    buildJenjangDropdown();
    buildGrafikJenjangDropdown();
    buildLegerDropdown();
    saveToLocalStorage();
    saveUIState();
    updateQuartalTabsUI();
}

async function loadExcel(file) {
    uploadBtnText.innerText = "Memuat modul...";
    try {
        await ensureXLSX();
    } catch (err) {
        alert("Gagal memuat modul pembaca Excel. Cek koneksi internet.");
        uploadBtnText.innerText = "Upload Data";
        return;
    }
    uploadBtnText.innerText = "Membaca file...";
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            workbookData = XLSX.read(data, { type: 'array', cellDates: false });
            processAllSheets();
            let shortName = file.name;
            if (shortName.length > 15) shortName = shortName.substring(0, 15) + '...';
            uploadBtnText.innerText = shortName;
        } catch (err) { alert("Gagal membaca file: " + file.name);
            uploadBtnText.innerText = "Upload Data"; }
    };
    reader.readAsArrayBuffer(file);
}

// Quartal Management
function renderQuartalView(quartal) {
    const cache = quartalCache[quartal];
    if (!cache || (!Object.keys(cache.groupedData).length && !Object.keys(cache.allSheetsData || {}).length)) {
        filterBar.style.display = 'none';
        jenjangDropdownBar.style.display = 'none';
        sheetSelect.innerHTML = '';
        jenjangSelect.innerHTML = '';
        mapelSelect.innerHTML = '';
        uploadBtnText.innerText = "Upload Data";
        groupedData = {};
        allSheetsData = {};
        originalSheetNames = [];
        renderKelasTable([], [], '');
        buildLegerDropdown();
        jenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Menunggu data komparatif...</p>';
        saveUIState();
        return;
    }

    groupedData = cache.groupedData;
    allSheetsData = cache.allSheetsData || {};
    waliKelas = cache.waliKelas || waliKelas;
    originalSheetNames = cache.originalSheetNames || Object.keys(allSheetsData);
    if (!originalSheetNames.length && Object.keys(allSheetsData).length) originalSheetNames = Object.keys(
    allSheetsData);
    if (originalSheetNames.length && originalSheetNames.every(name => allSheetsData[name])) {
        const groups = {};
        originalSheetNames.forEach(name => {
            const key = getClassGroupKey(name);
            if (!groups[key]) groups[key] = [];
            groups[key].push(name);
        });
        groupedData = Object.fromEntries(Object.entries(groups).map(([key, names]) => [
            key, mergeCachedSheetData(names, allSheetsData)
        ]));
        cache.groupedData = groupedData;
    }

    const uiState = getUIState();
    let targetSheet = uiState.sheet;

    sheetSelect.innerHTML = '';
    if (originalSheetNames.length) {
        originalSheetNames.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            sheetSelect.appendChild(opt);
        });
        filterBar.style.display = 'flex';
        currentSheetName = originalSheetNames.includes(targetSheet) ? targetSheet : originalSheetNames[0];
        sheetSelect.value = currentSheetName;
    }

    let found = false;
    if (targetSheet) {
        if (groupedData[targetSheet] && groupedData[targetSheet].subjects && groupedData[targetSheet].subjects
            .length) {
            currentSubjectNames = groupedData[targetSheet].subjects;
            currentAverages = groupedData[targetSheet].averages;
            currentStudentCount = groupedData[targetSheet].studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, targetSheet.replace('_', ' '),
            currentStudentCount, groupedData[targetSheet]);
            found = true;
        } else if (allSheetsData[targetSheet] && allSheetsData[targetSheet].subjects && allSheetsData[
            targetSheet].subjects.length) {
            currentSubjectNames = allSheetsData[targetSheet].subjects;
            currentAverages = allSheetsData[targetSheet].averages;
            currentStudentCount = allSheetsData[targetSheet].studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, targetSheet, currentStudentCount, allSheetsData[targetSheet]);
            found = true;
        }
    }

    if (!found && originalSheetNames.length) {
        const first = originalSheetNames[0];
        const fd = allSheetsData[first];
        if (fd && fd.subjects && fd.subjects.length) {
            currentSubjectNames = fd.subjects;
            currentAverages = fd.averages;
            currentStudentCount = fd.studentCount || 0;
            renderKelasTable(currentSubjectNames, currentAverages, first, currentStudentCount, fd);
            found = true;
        }
    }

    if (!found) renderKelasTable([], [], '');
    uploadBtnText.innerText = "Upload Data";
    buildJenjangDropdown();
    buildGrafikJenjangDropdown();
    buildLegerDropdown();
    saveUIState();
}

function updateQuartalTabsUI() {
    quartalTabs.querySelectorAll('.quartal-btn').forEach(btn => {
        const q = btn.dataset.q;
        const cache = quartalCache[q];
        const hasData = cache && (Object.keys(cache.groupedData || {}).length || Object.keys(cache
            .allSheetsData || {}).length);
        btn.innerHTML = QUARTAL_LABELS[q] + (hasData ? ' <span class="qdot"></span>' : '');
        btn.classList.toggle('active-quartal', q === currentQuartal);
    });
}

function loadAllFromLocalStorage() {
    quartalCache = {};
    ['Kuartal_I', 'Kuartal_II', 'Kuartal_III'].forEach(q => {
        const data = loadFromLocalStorage(q);
        if (data) quartalCache[q] = data;
        else quartalCache[q] = { groupedData: {}, allSheetsData: {}, originalSheetNames: [] };
    });

    const uiState = getUIState();
    if (uiState.quartal && ['Kuartal_I', 'Kuartal_II', 'Kuartal_III'].includes(uiState.quartal)) {
        currentQuartal = uiState.quartal;
    } else {
        const firstWithData = ['Kuartal_I', 'Kuartal_II', 'Kuartal_III'].find(q => {
            const c = quartalCache[q];
            return c && (Object.keys(c.groupedData).length || Object.keys(c.allSheetsData || {}).length);
        });
        currentQuartal = firstWithData || 'Kuartal_I';
    }

    if (uiState.kkm !== undefined && Number.isFinite(Number(uiState.kkm))) {
        const oldScale = uiState.kkmScale === 100 ? 100 : 10;
        if (uiState.kkm >= 0 && uiState.kkm <= oldScale) {
            KKM = Number(uiState.kkm) * (100 / oldScale);
        }
        kkmInput.value = KKM;
    }

    updateQuartalTabsUI();
    renderQuartalView(currentQuartal);
    saveUIState();
}

function clearStoredData() {
    if (!confirm(`Hapus semua data ${QUARTAL_LABELS[currentQuartal]}? Aksi ini tidak bisa dibatalkan.`)) return;
    clearLocalStorage(currentQuartal);
    quartalCache[currentQuartal] = { groupedData: {}, allSheetsData: {}, originalSheetNames: [] };
    workbookData = null;
    currentSubjectNames = [];
    currentAverages = [];
    currentStudentCount = 0;
    currentSheetName = '';
    originalSheetNames = [];
    saveUIState();
    updateQuartalTabsUI();
    renderQuartalView(currentQuartal);
}

// Event Listeners
document.addEventListener('input', e => {
    if (!e.target.matches('.failed-student-search-input')) return;
    const query = e.target.value.trim().toLocaleLowerCase();
    const panel = e.target.closest('.failed-students');
    let matchingStudents = 0;
    panel.querySelectorAll('.failed-student-subject').forEach(subject => {
        const students = subject.querySelectorAll('li');
        let subjectHasMatch = false;
        students.forEach(student => {
            const matches = student.querySelector('.failed-student-name').textContent.toLocaleLowerCase().includes(query);
            student.hidden = !matches;
            subjectHasMatch ||= matches;
        });
        subject.hidden = !subjectHasMatch;
        if (subjectHasMatch) matchingStudents++;
    });
    panel.querySelector('.failed-student-no-results').hidden = matchingStudents > 0;
});

document.addEventListener('click', e => {
   const rekapToggle = e.target.closest('[data-rekap-student-toggle]');
   if (rekapToggle) {
       const list = rekapToggle.closest('.rekap-student-group').querySelector('.rekap-student-list');
       const isExpanded = rekapToggle.getAttribute('aria-expanded') === 'true';
       rekapToggle.setAttribute('aria-expanded', String(!isExpanded));
       list.hidden = isExpanded;
       return;
   }
   const button = e.target.closest('.failed-subject-toggle');
   if (button) {
       const panel = button.closest('.failed-students');
       const isActive = button.classList.toggle('active-chip');
       button.setAttribute('aria-pressed', String(isActive));
       panel.querySelector(`.failed-student-subject[data-subject-index="${button.dataset.subjectIndex}"]`)
           .classList.toggle('is-hidden', !isActive);
       return;
   }
   const action = e.target.closest('[data-failed-subject-action]');
   if (!action) return;
   const isActive = action.dataset.failedSubjectAction === 'all';
   const studentsPanel = action.closest('.failed-students');
   const toggles = studentsPanel.querySelectorAll('.failed-subject-toggle');
   const subjects = studentsPanel.querySelectorAll('.failed-student-subject');
   toggles.forEach((toggle, index) => {
       toggle.classList.toggle('active-chip', isActive);
       toggle.setAttribute('aria-pressed', String(isActive));
       subjects[index].classList.toggle('is-hidden', !isActive);
   });
});

quartalTabs.querySelectorAll('.quartal-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        currentQuartal = btn.dataset.q;
        saveUIState();
        updateQuartalTabsUI();
        renderQuartalView(currentQuartal);
    });
});

function handleKKMChange(newValue) {
    const val = parseFloat(newValue);
    if (isNaN(val) || val < 0 || val > 100) { kkmInput.value = KKM; return; }
    KKM = Number.parseFloat(val.toFixed(2));
    const displayValue = Number.isInteger(KKM) ? String(KKM) : KKM.toString();
    kkmInput.value = displayValue;
    saveUIState();
    refreshAllViews();
}

kkmInput.addEventListener('change', function() { handleKKMChange(this.value); });
kkmInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') { e.preventDefault();
        handleKKMChange(this.value);
        this.blur(); }
});
resetKkmBtn.addEventListener('click', function() { KKM = DEFAULT_KKM;
    kkmInput.value = KKM;
    saveUIState();
    refreshAllViews(); });

fileInput.addEventListener('change', (e) => { if (e.target.files.length) loadExcel(e.target.files[0]);
    fileInput.value = ''; });
sheetSelect.addEventListener('change', (e) => { if (!e.target.value) return;
    filterBySheet(e.target.value); });

jenjangSelect.addEventListener('change', (e) => {
    if (!e.target.value) return;
    populateMapelSelect(e.target.value);
    renderSingleJenjang(e.target.value, getSelectedMapels());
    renderJenjangStudentSearch();
});

jenjangStudentSearchInput.addEventListener('input', renderJenjangStudentSearch);
jenjangStudentSearchInput.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
        jenjangStudentSearchInput.value = '';
        renderJenjangStudentSearch();
    }
});

document.addEventListener('click', e => {
    if (!e.target.closest('#jenjangStudentSearch')) jenjangStudentSearchResults.hidden = true;
});

grafikJenjangSelect.addEventListener('change', (e) => {
    if (!e.target.value) return;
    populateGrafikMapelSelect(e.target.value);
    renderJenjangCompareChart(e.target.value, getSelectedGrafikMapels());
});

selectAllMapelBtn.addEventListener('click', () => {
    mapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.add('active-chip');
        btn.innerHTML = `${arabicTextHTML(btn.dataset.value)} <span class="mapel-chip-x">&times;</span>`;
    });
    const key = jenjangSelect.value;
    if (key) renderSingleJenjang(key, getSelectedMapels());
});
deselectAllMapelBtn.addEventListener('click', () => {
    mapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.remove('active-chip');
        btn.innerHTML = arabicTextHTML(btn.dataset.value);
    });
    const key = jenjangSelect.value;
    if (key) renderSingleJenjang(key, getSelectedMapels());
});
grafikSelectAllMapelBtn.addEventListener('click', () => {
    grafikMapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.add('active-chip');
        btn.innerHTML = `${arabicTextHTML(btn.dataset.value)} <span class="mapel-chip-x">&times;</span>`;
    });
    const key = grafikJenjangSelect.value;
    if (key) renderJenjangCompareChart(key, getSelectedGrafikMapels());
});
grafikDeselectAllMapelBtn.addEventListener('click', () => {
    grafikMapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.remove('active-chip');
        btn.innerHTML = arabicTextHTML(btn.dataset.value);
    });
    const key = grafikJenjangSelect.value;
    if (key) renderJenjangCompareChart(key, getSelectedGrafikMapels());
});

let paneSwitching = false;
function switchPane(paneId) {
    if (paneSwitching) return;
    const nextPane = document.getElementById('pane-' + paneId);
    const currentPane = document.querySelector('.main-pane.active');
    if (!nextPane || currentPane === nextPane) return;

    document.querySelectorAll('.main-tab').forEach(t => t.classList.remove('active'));
    const matchingTab = document.querySelector(`.main-tab[data-pane="${paneId}"]`);
    if (matchingTab) matchingTab.classList.add('active');
    document.getElementById('mainTabs').classList.toggle('kelas-active', paneId === 'kelas');
    document.getElementById('mainTabs').classList.toggle('jenjang-active', paneId === 'jenjang');
    if (paneId !== 'jenjang') jenjangStudentSearchResults.hidden = true;

    paneSwitching = true;

    const showNextPane = () => {
        nextPane.classList.add('active');
        // force reflow so the transition to pane-show plays
        void nextPane.offsetWidth;
        nextPane.classList.add('pane-show');

        if (paneId === 'overview') renderOverview();

        runCountUps(nextPane);
        paneSwitching = false;
    };

    if (currentPane) {
        currentPane.classList.remove('pane-show');
        setTimeout(() => {
            currentPane.classList.remove('active');
            showNextPane();
        }, 220);
    } else {
        showNextPane();
    }
}

document.querySelectorAll('.main-tab').forEach(tab => {
    tab.addEventListener('click', function() {
        switchPane(this.dataset.pane);
    });
});

legerSheetSelect.addEventListener('change', () => renderLeger());

document.getElementById('exportPDFLeger')?.addEventListener('click', async () => {
    const sheetName = legerSheetSelect.value;
    const data = allSheetsData[sheetName];
    const sourceTable = legerContainer.querySelector('.leger-table');
    if (!data?.scores?.length || !sourceTable) {
        alert('Tidak ada data leger untuk diekspor.');
        return;
    }

    const button = document.getElementById('exportPDFLeger');
    button.disabled = true;
    try {
        await ensurePDFLibs();
        if (document.fonts?.ready) await document.fonts.ready;

        const sourceRows = [...sourceTable.tBodies[0].rows];
        const rowsPerPage = 30;
        const pageCount = Math.ceil(sourceRows.length / rowsPerPage);
        const pdfBackground = '#FFFFFF';
        const doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

        for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            const page = document.createElement('div');
            page.className = 'leger-pdf-page';
            const teacher = waliKelas[sheetName] || data.asuhan || '';
            page.innerHTML = `
                <header>
                    <h1>LEGER NILAI UJIAN</h1>
                    <p>Madrasah Diniyah "Miftahul Ulum" · ${escapeHTML(formatClassLabel(sheetName))}${teacher ? ` · Wali Kelas: ${arabicTextHTML(teacher)}` : ''}</p>
                    <small>Halaman ${pageIndex + 1} dari ${pageCount} · ${new Date().toLocaleDateString('id-ID')}</small>
                </header>`;
            const table = sourceTable.cloneNode(false);
            table.appendChild(sourceTable.tHead.cloneNode(true));
            const tbody = document.createElement('tbody');
            sourceRows.slice(pageIndex * rowsPerPage, (pageIndex + 1) * rowsPerPage)
                .forEach(row => tbody.appendChild(row.cloneNode(true)));
            table.appendChild(tbody);
            page.appendChild(table);
            pdfContainer.innerHTML = '';
            pdfContainer.appendChild(page);
            rasterizeArabicText(page);
            await new Promise(resolve => setTimeout(resolve, 60));

            const canvas = await html2canvas(page, {
                backgroundColor: pdfBackground,
                scale: 3,
                useCORS: true
            });
            if (pageIndex > 0) doc.addPage('a4', 'landscape');
            const imageW = canvas.width / 3 * 0.264583;
            const imageH = canvas.height / 3 * 0.264583;
            const pdfMargin = 1.5;
            const usableWmm = 297 - (pdfMargin * 2);
            const usableHmm = 210 - (pdfMargin * 2);
            const scale = Math.min(usableWmm / imageW, usableHmm / imageH);
            const width = imageW * scale;
            const height = imageH * scale;
            doc.addImage(canvas.toDataURL('image/png'), 'PNG',
                (297 - width) / 2, pdfMargin, width, height);
        }

        const safe = sheetName.replace(/[^a-zA-Z0-9]/g, '_');
        doc.save(`Leger_${safe}.pdf`);
    } catch (error) {
        console.error('Gagal membuat PDF leger:', error);
        alert('Gagal membuat PDF leger. Periksa koneksi internet dan coba lagi.');
    } finally {
        pdfContainer.innerHTML = '';
        button.disabled = false;
    }
});

document.getElementById('exportExcelLeger').addEventListener('click', exportAllLegerToExcel);

document.getElementById('exportPDFKelas').addEventListener('click', () => {
    if (!currentSubjectNames.length) { alert("Tidak ada data"); return; }
    const src = document.querySelector('#pane-kelas .section-container');
    exportPDF(currentSubjectNames, currentAverages, currentSheetName || 'Kelas', src);
});

document.getElementById('exportCSVKelas').addEventListener('click', () => {
    if (!currentSubjectNames.length) { alert("Tidak ada data"); return; }
    exportCSV(currentSubjectNames, currentAverages, currentSheetName || 'Kelas');
});

document.getElementById('exportPDFOverview').addEventListener('click', () => {
    if (!originalSheetNames.length) { alert("Tidak ada data"); return; }
    const src = document.querySelector('#pane-overview .section-container');
    exportPDF(['Ringkasan Sekolah'], [0], 'Ringkasan Sekolah', src, 'Laporan Ringkasan');
});

document.getElementById('exportCSVOverview').addEventListener('click', exportOverviewCSV);

document.getElementById('clearDataBtn').addEventListener('click', clearStoredData);

const THEME_BODY_CLASSES = ['theme-green'];

function applyTheme(theme) {
    theme = theme === 'green' ? 'green' : 'purple';
    document.body.classList.remove(...THEME_BODY_CLASSES);
    if (theme === 'green') document.body.classList.add('theme-green');
    localStorage.setItem('theme', theme);
    updateThemeUI(theme);
}

function updateThemeUI(theme) {
    themeBtns.forEach(btn => btn.classList.toggle('active-theme', btn.dataset.theme === theme));
}

themeBtns.forEach(btn => btn.addEventListener('click', () => applyTheme(btn.dataset.theme)));

const savedTheme = localStorage.getItem('theme') || 'purple';
applyTheme(savedTheme);

const dashboardEl = document.querySelector('.dashboard');
dashboardEl.addEventListener('dragover', (e) => { e.preventDefault();
    e.stopPropagation(); });
dashboardEl.addEventListener('drop', (e) => { e.preventDefault();
    e.stopPropagation(); if (e.dataTransfer.files.length) loadExcel(e.dataTransfer.files[0]); });

updateQuartalTabsUI();
loadAllFromLocalStorage();

// ========== SIDEBAR NAV ==========
document.querySelectorAll('.sidebar-item').forEach(item => {
    item.addEventListener('click', function() {
        document.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active-nav'));
        this.classList.add('active-nav');
        const target = this.dataset.side;
        if (target === 'overview') {
            switchPane('overview');
        } else if (target === 'kelas') {
            switchPane('kelas');
        } else if (target === 'jenjang') {
            switchPane('jenjang');
        }
    });
});
