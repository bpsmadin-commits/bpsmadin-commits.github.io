// Lazy-load library berat: baru diambil pas benar-benar dipakai
const CDN = {
    xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js',
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
    let studentCount = 0;
    for (const name of sheetList) {
        const data = sheetsData[name];
        if (!data || !data.subjects.length || !data.scores.length) continue;
        if (!subjects.length) subjects = data.subjects.slice();
        if (data.subjects.length === subjects.length) {
            scores = scores.concat(data.scores);
            studentNames = studentNames.concat(data.studentNames || data.scores.map(() => ''));
            studentCount += data.studentCount || 0;
        }
    }
    return { subjects, averages: computeAverages(subjects, scores), studentCount, scores, studentNames };
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
const kelasKetuntasanBox = document.getElementById('kelasKetuntasanBox');
const kelasMiniStats = document.getElementById('kelasMiniStats');
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
    const data = { groupedData, allSheetsData, originalSheetNames };
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
    return String(value).trim();
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
        }
    }

    const asuhanRaw = (asuhanColIdx !== -1 && headerRowIndex !== -1 && rows[headerRowIndex]) ?
        rows[headerRowIndex][asuhanColIdx] : '';

    let asuhanVal = detectAsuhanFromRows(rows);
    if (!asuhanVal && asuhanRaw !== '') {
        asuhanVal = parseAsuhanText(asuhanRaw);
    }

    return { subjects: subjectNames, scores: studentScores, studentNames, studentCount: studentScores.length, asuhan: asuhanVal };
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
        totalStudents = 0;
    for (const sheetName of sheetList) {
        const ws = workbookData.Sheets[sheetName];
        if (!ws) continue;
        const { subjects, scores, studentNames, studentCount } = detectHeaderAndData(ws);
        if (subjects.length === 0 || scores.length === 0) continue;
        if (allSubjects.length === 0) allSubjects = subjects.slice();
        for (let i = 0; i < scores.length; i++) {
            if (scores[i].length === allSubjects.length) {
                allScores.push(scores[i].slice());
                allStudentNames.push(studentNames[i] || '');
            }
        }
        totalStudents += studentCount;
    }
    return { subjects: allSubjects, averages: computeAverages(allSubjects, allScores), studentCount: totalStudents,
        scores: allScores, studentNames: allStudentNames };
}

function computeStudentAverages(scores) {
    if (!scores || !scores.length) return [];
    return scores.map(student => {
        let sum = 0,
            count = 0;
        for (let val of student) { if (typeof val === 'number' && !isNaN(val)) { sum += val;
                count++; } }
        return count > 0 ? sum / count : 0;
    });
}

function getClassNumber(sheetName) {
    const digits = sheetName.match(/\d+/);
    if (digits) return parseInt(digits[0], 10);

    const roman = sheetName.match(/\b(XII|XI|IX|VIII|VII|VI|X|IV|V|III|II|I)\b/i);
    if (!roman) return 0;
    return { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10, XI: 11, XII: 12 }[roman[0].toUpperCase()];
}

function getClassType(sheetName) {
    const specificType = sheetName.match(/MDTA|MDTW|MTSD/i);
    if (specificType) return specificType[0].toUpperCase();
    return /MTS/i.test(sheetName) ? 'MTSD' : 'MID';
}

function getClassGroupKey(sheetName) {
    return `${getClassNumber(sheetName)}_${getClassType(sheetName)}`;
}

function formatClassLabel(sheetName) {
    return sheetName.replace(/^KLS\b/i, 'Kelas').replace(/[-_]+/g, ' ').replace(/\s+/g, ' ').trim();
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
    if (avg >= 75) return 'good';
    if (avg >= KKM) return 'ok';
    return 'bad';
}

function getProgressBarHTML(avg) {
    const pct = Math.min(100, avg);
    const cls = getProgressBarClass(avg);
    return `<div class="progress-cell"><div class="progress-bar-wrap"><div class="progress-bar-fill ${cls}" style="width:0%" data-target-width="${pct}%"></div></div><strong class="count-value" data-count-to="${avg}" data-decimals="1">0.0</strong></div>`;
}

// Animate a numeric text node counting up from 0 to its target value
let countAnimSeq = 0;
function animateCountValue(el, target, decimals, suffix, duration = 1000) {
    const seq = ++countAnimSeq;
    el.dataset.animSeq = seq;
    const startTime = performance.now();
    function step(now) {
        if (parseInt(el.dataset.animSeq, 10) !== seq) return; // a newer animation took over
        const t = Math.min((now - startTime) / duration, 1);
        const eased = 1 - Math.pow(1 - t, 4);
        el.textContent = (target * eased).toFixed(decimals) + suffix;
        if (t < 1) requestAnimationFrame(step);
        else el.textContent = target.toFixed(decimals) + suffix;
    }
    requestAnimationFrame(step);
}

// Find every [data-count-to] / progress bar inside `scope` and (re)animate them from 0.
// Safe to call repeatedly — e.g. every time a tab/pane is switched into view.
function runCountUps(scope) {
    const root = scope || document;

    // Snap progress bars back to 0 instantly so the fill can replay every single time
    const bars = root.querySelectorAll('.progress-bar-fill[data-target-width]');
    bars.forEach(bar => {
        bar.style.transition = 'none';
        bar.style.width = '0%';
    });

    // Sama juga buat batang grafik perjenjang (height, bukan width)
    const jcBars = root.querySelectorAll('.jc-bar-fill[data-target-height]');
    jcBars.forEach(bar => {
        bar.style.transition = 'none';
        bar.style.height = '0%';
    });

    void root.offsetHeight; // force reflow before re-enabling the transition
    bars.forEach(bar => { bar.style.transition = ''; });
    jcBars.forEach(bar => { bar.style.transition = ''; });

    root.querySelectorAll('.count-value[data-count-to]').forEach(el => {
        const target = parseFloat(el.dataset.countTo);
        if (isNaN(target)) return;
        const decimals = el.dataset.decimals !== undefined ? parseInt(el.dataset.decimals, 10) : 2;
        const suffix = el.dataset.suffix || '';
        animateCountValue(el, target, decimals, suffix);
    });

    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            bars.forEach(bar => { bar.style.width = bar.dataset.targetWidth; });
            jcBars.forEach(bar => { bar.style.height = bar.dataset.targetHeight; });
        });
    });
}

function getStatusBadge(avg) {
    if (avg >= 75) return '<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Baik</span>';
    if (avg >= KKM) return '<span class="badge badge-warning"><i class="fa-solid fa-exclamation-triangle"></i> Cukup</span>';
    return '<span class="badge badge-danger"><i class="fa-solid fa-circle-exclamation"></i> Perlu Bimbingan</span>';
}

function getStatusText(avg) {
    if (avg >= 75) return "Baik";
    if (avg >= KKM) return "Cukup";
    return "Perlu Bimbingan";
}

function isBelowKKM(avg) { return avg < KKM && avg > 0; }

function buildMiniStatsHTML(subjects, averages, wali = '') {
    if (!subjects.length || !averages.length) return '';
    const maxVal = Math.max(...averages),
        minVal = Math.min(...averages);
    const maxIdx = averages.indexOf(maxVal),
        minIdx = averages.indexOf(minVal);
    const overall = (averages.reduce((a, b) => a + b, 0) / averages.length).toFixed(1);
    const waliBadge = wali ? `
            <div class="mini-wali" style="margin:0 0 12px; padding:8px 12px; border-radius:999px; background:rgba(255,255,255,0.06); border:1px solid rgba(255,255,255,0.12); color:var(--muted); font-size:0.8rem; display:inline-flex; align-items:center; gap:8px;">
                <i class="fa-solid fa-user-tie"></i><span>Wali Kelas: <strong style="color:var(--text);">${wali}</strong></span>
            </div>` : '';
    return `
            ${waliBadge}
            <div class="mini-stats">
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-trophy"></i> Tertinggi</div><div class="mini-value count-value" data-count-to="${maxVal}" data-decimals="1">0.0</div><div class="mini-sub" style="font-family:'Amiri','Traditional Arabic',serif;">${subjects[maxIdx]}</div></div>
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-arrow-trend-down"></i> Terendah</div><div class="mini-value count-value" data-count-to="${minVal}" data-decimals="1">0.0</div><div class="mini-sub" style="font-family:'Amiri','Traditional Arabic',serif;">${subjects[minIdx]}</div></div>
                <div class="mini-stat"><div class="mini-label"><i class="fa-solid fa-bullseye"></i> Rata-rata</div><div class="mini-value count-value" data-count-to="${overall}" data-decimals="1">0.0</div><div class="mini-sub">Keseluruhan</div></div>
            </div>`;
}

// Render Per Kelas
function renderKelasTable(subjects, averages, sheetName, studentCount = 0, sourceData = null) {
    const title = sheetName || currentSheetName || 'Kelas';
    const wali = waliKelas[sheetName || currentSheetName] || (sourceData?.asuhan) || '';
    kelasTableTitle.textContent = `Tabel Analisis — ${title}` + (studentCount ? ` (${studentCount} siswa)` : '') + (wali ? ` · Wali Kelas: ${wali}` : '');

    if (!subjects.length || !averages.length) {
        tableBodyKelas.innerHTML =
            `<tr><td colspan="4" style="padding:0;"><div class="empty-dropzone"><div class="empty-dropzone-icon"><i class="fa-solid fa-file-excel"></i></div><div class="empty-dropzone-title">Belum ada data</div><div class="empty-dropzone-text">Silakan upload file Excel nilai siswa untuk mulai melihat analisis performa akademik.</div><label for="file-input" class="empty-dropzone-btn"><i class="fa-solid fa-cloud-arrow-up"></i> Upload Data Sekarang</label></div></td></tr>`;
        kelasKetuntasanBox.innerHTML = '';
        kelasMiniStats.innerHTML = '';
        return;
    }

    const sorted = sortSubjectsByAverageDesc(subjects, averages);
    currentSubjectNames = sorted.subjects;
    currentAverages = sorted.averages;
    currentStudentCount = studentCount;
    // Status & dataset di hero sudah dihapus, jadi tidak perlu update elemen itu lagi

    kelasMiniStats.innerHTML = buildMiniStatsHTML(currentSubjectNames, currentAverages, wali);

    tableBodyKelas.innerHTML = '';
    currentSubjectNames.forEach((sub, idx) => {
        const avg = currentAverages[idx];
        const kkmClass = isBelowKKM(avg) ? 'below-kkm' : '';
        const row = tableBodyKelas.insertRow();
        row.className = kkmClass;
        row.insertCell(0).innerText = idx + 1;
        row.insertCell(1).innerHTML = `<span class="mapel-name">${sub}</span>`;
        row.insertCell(2).innerHTML = getProgressBarHTML(avg);
        row.insertCell(3).innerHTML = getStatusBadge(avg);
    });

    kelasKetuntasanBox.innerHTML = sourceData ? buildKetuntasanHTML(sourceData, null, false) : '';

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
            `<tr class="${kkmClass}"><td>${i+1}</td><td><span class="mapel-name">${sub}</span></td><td>${getProgressBarHTML(avg)}</td><td>${getStatusBadge(avg)}</td></tr>`;
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
        btn.innerHTML = `${sub} <span class="mapel-chip-x">&times;</span>`;
        btn.addEventListener('click', () => {
            btn.classList.toggle('active-chip');
            btn.innerHTML = btn.classList.contains('active-chip')
                ? `${sub} <span class="mapel-chip-x">&times;</span>`
                : sub;
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
        btn.innerHTML = `${sub} <span class="mapel-chip-x">&times;</span>`;
        btn.addEventListener('click', () => {
            btn.classList.toggle('active-chip');
            btn.innerHTML = btn.classList.contains('active-chip')
                ? `${sub} <span class="mapel-chip-x">&times;</span>`
                : sub;
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
                <div class="jc-bar-label">${formatClassLabel(b.name)}</div>
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

function buildKetuntasanHTML(data, selectedIndices, showChart, titleSuffix, showFailedStudents = true) {
    const allScores = data.scores || [];
    // Jika ada mapel yang dipilih (selectedIndices), hitung ketuntasan hanya dari kolom mapel itu.
    // Jika tidak (null/kosong), pakai semua mapel seperti semula.
    const scores = (selectedIndices && selectedIndices.length && allScores.length) ?
        allScores.map(student => selectedIndices.map(i => student[i])) :
        allScores;

    let totalStudents = data.studentCount || 0;
    let tuntas = 0,
        tidakTuntas = 0,
        persentaseTuntas = 0;

    if (scores.length) {
        const studentAverages = computeStudentAverages(scores);
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
    const failedStudents = showFailedStudents ? computeStudentAverages(scores)
        .map((average, index) => ({ name: (data.studentNames || [])[index], average }))
        .filter(student => student.average < KKM) : [];
    const failedStudentsHTML = failedStudents.length ? `
            <details class="failed-students">
                <summary>Siswa Belum Tuntas (${failedStudents.length})</summary>
                ${failedStudents.some(student => student.name) ? `
                <ol>${failedStudents.map(student => `<li><span>${student.name ? escapeHTML(student.name) : 'Nama tidak terbaca'}</span><strong>${student.average.toFixed(1)}</strong></li>`).join('')}</ol>` :
                '<p>Nama siswa tidak tersedia. Upload ulang file Excel untuk memuat nama.</p>'}
            </details>` : '';

    return `
            <div class="section-title" style="font-size:1.1rem; margin-top:24px;"><i class="fa-solid fa-ranking-star"></i> Ketuntasan${titleSuffix ? ' ' + titleSuffix : ''}</div>
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
                <div class="rekap-mm-card"><div class="mm-label" style="color:var(--success);"><i class="fa-solid fa-check-circle"></i> Tuntas</div><div class="mm-value count-value" data-count-to="${persentaseTuntas}" data-decimals="1" data-suffix="%">0%</div><div class="mm-subject">${tuntas} siswa</div></div>
                <div class="rekap-mm-card"><div class="mm-label" style="color:var(--danger);"><i class="fa-solid fa-times-circle"></i> Tidak Tuntas</div><div class="mm-value count-value" data-count-to="${100 - persentaseTuntas}" data-decimals="1" data-suffix="%">0%</div><div class="mm-subject">${tidakTuntas} siswa</div></div>
            </div>
            ${failedStudentsHTML}
            <div style="margin-top: 12px; background: var(--surface-soft); border-radius: var(--radius-lg); padding: 8px 12px; display: flex; justify-content: space-between; font-size: 0.8rem; color: var(--muted);"><span>Rerata: ${overall}</span><span>KKM: ${KKM}</span></div>`;
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

// Pisahin daftar nama sheet jadi 2 kelompok jenjang: MID sama MTsD, pake aturan
// yang sama kayak getSheetsInJenjang (nama sheet yang ngandung "MTS" masuk MTsD).
function splitSheetsByJenjang(sheetNames) {
    const mid = [];
    const mtsd = [];
    (sheetNames || []).forEach(name => {
        if (/MTS/i.test(name)) mtsd.push(name); else mid.push(name);
    });
    return { mid, mtsd };
}

const JENJANG_META = {
    MID:  { label: 'MID',  subtitle: 'Madrasah Ibtidaiyah Diniyah',  icon: 'fa-seedling',        accentClass: 'accent-mid' },
    MTSD: { label: 'MTsD', subtitle: 'Madrasah Tsanawiyah Diniyah', icon: 'fa-graduation-cap',  accentClass: 'accent-mtsd' }
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
                    <div class="overview-highlight-name">${best.name}</div>
                    <div class="overview-highlight-sub">Ketuntasan ${best.pct.toFixed(1)}% &middot; Rata-rata ${best.avg.toFixed(1)}</div>
                </div>`;
        }
        if (worst) {
            highlightHTML += `
                <div class="overview-highlight-card bad">
                    <div class="mini-label"><i class="fa-solid fa-arrow-trend-down"></i> Perlu Perhatian</div>
                    <div class="overview-highlight-name">${worst.name}</div>
                    <div class="overview-highlight-sub">Ketuntasan ${worst.pct.toFixed(1)}% &middot; Rata-rata ${worst.avg.toFixed(1)}</div>
                </div>`;
        }

        // Tren rata-rata antar Kuartal I/II/III buat jenjang ini, ditarik dari quartalCache
        const qKeys = ['Kuartal_I', 'Kuartal_II', 'Kuartal_III'];
        const trendBars = qKeys.map(q => {
            const cache = quartalCache[q];
            const names = cache ? splitSheetsByJenjang(cache.originalSheetNames || [])[jenjangKey === 'MTSD' ? 'mtsd' : 'mid'] : [];
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

// Ringkasan seluruh sekolah, dipisah per jenjang (MID & MTsD), dipakai buat pane Overview
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

    const { mid, mtsd } = splitSheetsByJenjang(originalSheetNames);

    overviewContent.innerHTML = `
            ${buildJenjangSectionHTML('MID', mid)}
            ${buildJenjangSectionHTML('MTSD', mtsd)}`;

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
async function exportPDF(subjects, averages, title, sourceEl, reportLabel) {
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
        clone.querySelectorAll('.export-row').forEach(row => row.remove());
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
    const { mid, mtsd } = splitSheetsByJenjang(originalSheetNames);
    let csv = '\uFEFF';
    [{ key: 'MID', names: mid }, { key: 'MTSD', names: mtsd }].forEach(({ key, names }) => {
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
    const sheetNames = workbookData.SheetNames;
    if (!sheetNames.length) return;

    originalSheetNames = [...sheetNames];
    groupedData = {};
    allSheetsData = {};

    for (const name of sheetNames) {
        const ws = workbookData.Sheets[name];
        if (!ws) { allSheetsData[name] = { subjects: [], averages: [], studentCount: 0, scores: [], studentNames: [] }; continue; }
        const { subjects, scores, studentNames, studentCount, asuhan } = detectHeaderAndData(ws);
        const avgs = computeAverages(subjects, scores);
        allSheetsData[name] = { subjects, averages: avgs, studentCount, scores: scores || [], studentNames, asuhan };
        if (asuhan) waliKelas[name] = asuhan;
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
        jenjangContainer.innerHTML =
            '<p style="color:var(--muted); text-align:center; padding:40px; font-size:1rem;">Menunggu data komparatif...</p>';
        saveUIState();
        return;
    }

    groupedData = cache.groupedData;
    allSheetsData = cache.allSheetsData || {};
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
    KKM = Math.round(val);
    kkmInput.value = KKM;
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
});

grafikJenjangSelect.addEventListener('change', (e) => {
    if (!e.target.value) return;
    populateGrafikMapelSelect(e.target.value);
    renderJenjangCompareChart(e.target.value, getSelectedGrafikMapels());
});

selectAllMapelBtn.addEventListener('click', () => {
    mapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.add('active-chip');
        btn.innerHTML = `${btn.dataset.value} <span class="mapel-chip-x">&times;</span>`;
    });
    const key = jenjangSelect.value;
    if (key) renderSingleJenjang(key, getSelectedMapels());
});
deselectAllMapelBtn.addEventListener('click', () => {
    mapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.remove('active-chip');
        btn.innerHTML = btn.dataset.value;
    });
    const key = jenjangSelect.value;
    if (key) renderSingleJenjang(key, getSelectedMapels());
});
grafikSelectAllMapelBtn.addEventListener('click', () => {
    grafikMapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.add('active-chip');
        btn.innerHTML = `${btn.dataset.value} <span class="mapel-chip-x">&times;</span>`;
    });
    const key = grafikJenjangSelect.value;
    if (key) renderJenjangCompareChart(key, getSelectedGrafikMapels());
});
grafikDeselectAllMapelBtn.addEventListener('click', () => {
    grafikMapelSelect.querySelectorAll('.mapel-chip-btn').forEach(btn => {
        btn.classList.remove('active-chip');
        btn.innerHTML = btn.dataset.value;
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

const THEME_BODY_CLASSES = ['dark', 'theme-purple', 'theme-orange'];

function applyTheme(theme) {
    document.body.classList.remove(...THEME_BODY_CLASSES);
    if (theme === 'dark') document.body.classList.add('dark');
    else if (theme === 'purple') document.body.classList.add('theme-purple');
    else if (theme === 'orange') document.body.classList.add('theme-orange');
    // 'light' => no extra class, uses the :root defaults
    localStorage.setItem('theme', theme);
    updateThemeUI(theme);
}

function updateThemeUI(theme) {
    themeBtns.forEach(btn => btn.classList.toggle('active-theme', btn.dataset.theme === theme));
}

themeBtns.forEach(btn => btn.addEventListener('click', () => applyTheme(btn.dataset.theme)));

const savedTheme = localStorage.getItem('theme') || 'dark';
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
