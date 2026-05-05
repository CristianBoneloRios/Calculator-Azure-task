/* ═══════════════════════════════════════════════════════════════
   AZURE CSV NORMALIZER — normalizer.js
   Converts Azure DevOps-exported CSV/Excel to re-importable format
   
   Transformations applied:
   1. Clear numeric IDs (Azure auto-assigns on import)
   2. Strip email from "Assigned To" (e.g., "Name <email>" → "Name")
   3. Set State field for Test Case rows to a user-defined value
   ═══════════════════════════════════════════════════════════════ */

'use strict';

(function () {

  // ─── State ──────────────────────────────────────────────────────
  const NormState = {
    rawRows: null,     // parsed rows (array of arrays)
    headers: null,     // header row (array of strings)
    fileName: null,
    normalized: null   // normalized rows ready for export
  };

  // ─── DOM references ─────────────────────────────────────────────
  const $  = id => document.getElementById(id);
  const normUploadZone    = $('normUploadZone');
  const normFileInput     = $('normFileInput');
  const normSelectFileBtn = $('normSelectFileBtn');
  const normOptions       = $('normOptions');
  const normEmptyState    = $('normEmptyState');
  const normStatsBar      = $('normStatsBar');
  const normPreviewWrap   = $('normPreviewWrap');
  const normPreviewHead   = $('normPreviewHead');
  const normPreviewBody   = $('normPreviewBody');
  const normStateInput    = $('normStateInput');
  const normCleanAssignee = $('normCleanAssignee');
  const normClearId       = $('normClearId');
  const normDownloadBtn   = $('normDownloadBtn');
  const normClearBtn      = $('normClearBtn');
  const normStatFile      = $('normStatFile');
  const normStatCases     = $('normStatCases');
  const normStatSteps     = $('normStatSteps');

  // ─── Init ────────────────────────────────────────────────────────
  function init() {
    normSelectFileBtn.addEventListener('click', () => normFileInput.click());
    normFileInput.addEventListener('change', e => {
      if (e.target.files.length) handleFile(e.target.files[0]);
    });

    // Drag & drop
    normUploadZone.addEventListener('dragover', e => {
      e.preventDefault();
      normUploadZone.classList.add('dragging');
    });
    normUploadZone.addEventListener('dragleave', () => {
      normUploadZone.classList.remove('dragging');
    });
    normUploadZone.addEventListener('drop', e => {
      e.preventDefault();
      normUploadZone.classList.remove('dragging');
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    });

    // Option changes trigger re-normalization
    normStateInput.addEventListener('input', reNormalize);
    normCleanAssignee.addEventListener('change', reNormalize);
    normClearId.addEventListener('change', reNormalize);

    // Download
    normDownloadBtn.addEventListener('click', downloadNormalized);

    // Clear
    normClearBtn.addEventListener('click', resetNormalizer);
  }

  // ─── File handling ───────────────────────────────────────────────
  function handleFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    NormState.fileName = file.name;

    if (ext === 'csv') {
      parseCsv(file);
    } else if (ext === 'xlsx' || ext === 'xls') {
      parseExcel(file);
    } else {
      showNormToast('Formato no soportado. Usa CSV, XLSX o XLS.', 'error');
    }
  }

  function parseCsv(file) {
    Papa.parse(file, {
      skipEmptyLines: false,
      complete(result) {
        const data = result.data;
        if (!data || data.length < 2) {
          showNormToast('El archivo CSV está vacío o no tiene filas.', 'error');
          return;
        }
        NormState.headers = data[0];
        NormState.rawRows = data.slice(1);
        onDataLoaded();
      },
      error(err) {
        showNormToast('Error al leer el CSV: ' + err.message, 'error');
      }
    });
  }

  function parseExcel(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (!data || data.length < 2) {
          showNormToast('El archivo Excel está vacío o no tiene filas.', 'error');
          return;
        }
        NormState.headers = data[0].map(String);
        NormState.rawRows = data.slice(1);
        onDataLoaded();
      } catch (err) {
        showNormToast('Error al leer el Excel: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function onDataLoaded() {
    reNormalize();
    normOptions.style.display = '';
    normEmptyState.style.display = 'none';
    normStatsBar.style.display = '';
    normPreviewWrap.style.display = '';
  }

  // ─── Normalization ───────────────────────────────────────────────
  /**
   * Azure DevOps exported CSV columns (0-indexed):
   *   0  ID
   *   1  Work Item Type
   *   2  Title
   *   3  Test Step
   *   4  Step Action
   *   5  Step Expected
   *   6  Area Path
   *   7  Assigned To
   *   8  State
   */
  function colIndex(name) {
    if (!NormState.headers) return -1;
    const lower = name.toLowerCase();
    return NormState.headers.findIndex(h => String(h).toLowerCase().trim() === lower);
  }

  function reNormalize() {
    if (!NormState.rawRows) return;

    const idxID       = colIndex('id');
    const idxType     = colIndex('work item type');
    const idxAssignee = colIndex('assigned to');
    const idxState    = colIndex('state');

    const targetState   = normStateInput.value.trim();
    const cleanAssignee = normCleanAssignee.checked;
    const clearId       = normClearId.checked;

    const normalized = NormState.rawRows.map(row => {
      const r = row.map(cell => (cell === null || cell === undefined) ? '' : String(cell));

      // Pad row to at least header length
      while (r.length < NormState.headers.length) r.push('');

      const isTestCaseRow = idxType >= 0 &&
        r[idxType].trim().toLowerCase() === 'test case';

      // 1. Clear ID
      if (clearId && idxID >= 0 && isTestCaseRow) {
        r[idxID] = '';
      }

      // 2. Clean Assigned To — strip "<email>" portion
      if (cleanAssignee && idxAssignee >= 0 && r[idxAssignee]) {
        r[idxAssignee] = r[idxAssignee].replace(/\s*<[^>]+>/, '').trim();
      }

      // 3. Set State for Test Case rows
      if (targetState && idxState >= 0 && isTestCaseRow) {
        r[idxState] = targetState;
      }

      return r;
    });

    NormState.normalized = normalized;

    // Update stats
    const testCases = normalized.filter((r, i) => {
      const type = idxType >= 0 ? r[idxType] : '';
      return type.trim().toLowerCase() === 'test case';
    }).length;

    const stepRows = normalized.filter(r => {
      const step = idxType >= 0 ? r[idxType] : '';
      const testStep = colIndex('test step') >= 0 ? r[colIndex('test step')] : '';
      return step.trim() === '' && testStep !== '';
    }).length;

    normStatFile.textContent  = NormState.fileName;
    normStatCases.textContent = `${testCases} Test Case${testCases !== 1 ? 's' : ''}`;
    normStatSteps.textContent = `${stepRows} paso${stepRows !== 1 ? 's' : ''}`;

    renderPreview();
  }

  // ─── Preview ─────────────────────────────────────────────────────
  function renderPreview() {
    const headers = NormState.headers;
    const rows    = NormState.normalized;
    const preview = rows.slice(0, 30);

    // Head
    normPreviewHead.innerHTML = '<tr>' +
      headers.map(h => `<th>${escHtml(h)}</th>`).join('') +
      '</tr>';

    // Body
    const idxType = colIndex('work item type');
    normPreviewBody.innerHTML = preview.map(row => {
      const isTestCase = idxType >= 0 &&
        row[idxType].trim().toLowerCase() === 'test case';
      const cls = isTestCase ? ' class="norm-row-testcase"' : ' class="norm-row-step"';
      return `<tr${cls}>` +
        row.map(cell => `<td>${escHtml(cell)}</td>`).join('') +
        '</tr>';
    }).join('');
  }

  // ─── Download ─────────────────────────────────────────────────────
  function downloadNormalized() {
    if (!NormState.normalized) return;

    const allRows = [NormState.headers, ...NormState.normalized];

    const csv = Papa.unparse(allRows, {
      quotes: false,        // only quote when necessary
      quoteChar: '"',
      escapeChar: '"',
      delimiter: ',',
      newline: '\r\n',
      skipEmptyLines: false
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const base = NormState.fileName.replace(/\.[^.]+$/, '');
    a.href     = url;
    a.download = `${base}_normalizado.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showNormToast('¡Archivo descargado! Listo para importar en Azure DevOps.', 'success');
  }

  // ─── Reset ───────────────────────────────────────────────────────
  function resetNormalizer() {
    NormState.rawRows   = null;
    NormState.headers   = null;
    NormState.fileName  = null;
    NormState.normalized = null;

    normFileInput.value = '';
    normOptions.style.display       = 'none';
    normStatsBar.style.display      = 'none';
    normPreviewWrap.style.display   = 'none';
    normEmptyState.style.display    = '';
    normPreviewHead.innerHTML       = '';
    normPreviewBody.innerHTML       = '';
    normStateInput.value            = 'Design';
    normCleanAssignee.checked       = true;
    normClearId.checked             = true;
  }

  // ─── Helpers ─────────────────────────────────────────────────────
  function escHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function showNormToast(message, type) {
    // Reuse the app's toast system if available, otherwise fallback
    if (typeof showToast === 'function') {
      showToast(message, type);
    } else {
      const container = document.getElementById('toastContainer');
      if (!container) return;
      const toast = document.createElement('div');
      toast.className = `toast toast-${type}`;
      toast.innerHTML = `<i class="fas fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i> ${escHtml(message)}`;
      container.appendChild(toast);
      setTimeout(() => toast.classList.add('show'), 10);
      setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 350);
      }, 3500);
    }
  }

  // ─── Boot ────────────────────────────────────────────────────────
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
