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

  const DEFAULT_COLUMN_WIDTHS = [90, 160, 340, 90, 360, 360, 240, 240, 150];

  // ─── State ──────────────────────────────────────────────────────
  const NormState = {
    rawRows: null,     // parsed rows (array of arrays)
    headers: null,     // header row (array of strings)
    fileName: null,
    normalized: null,  // normalized rows ready for export
    columnWidths: [...DEFAULT_COLUMN_WIDTHS]
  };

  const TARGET_HEADERS = [
    'ID',
    'Work Item Type',
    'Title',
    'Test Step',
    'Step Action',
    'Step Expected',
    'Area Path',
    'Assigned To',
    'State'
  ];

  // ─── DOM references ─────────────────────────────────────────────
  const $  = id => document.getElementById(id);
  const normUploadZone    = $('normUploadZone');
  const normFileInput     = $('normFileInput');
  const normSelectFileBtn = $('normSelectFileBtn');
  const normOptions       = $('normOptions');
  const normEmptyState    = $('normEmptyState');
  const normStatsBar      = $('normStatsBar');
  const normPreviewWrap   = $('normPreviewWrap');
  const normPreviewTable  = $('normPreviewTable');
  const normPreviewHead   = $('normPreviewHead');
  const normPreviewBody   = $('normPreviewBody');
  const normAreaPathInput = $('normAreaPathInput');
  const normAssigneeEmailInput = $('normAssigneeEmailInput');
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
    normAreaPathInput.addEventListener('input', reNormalize);
    normAssigneeEmailInput.addEventListener('input', reNormalize);
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
    autoFillAreaPathFromData();
    autoFillAssigneeEmailFromData();
    reNormalize();
    normOptions.style.display = '';
    normEmptyState.style.display = 'none';
    normStatsBar.style.display = '';
    normPreviewWrap.style.display = '';
  }

  function autoFillAreaPathFromData() {
    if (!normAreaPathInput) return;

    const idxArea = colIndex('area path');
    if (idxArea < 0 || !NormState.rawRows || !NormState.rawRows.length) {
      normAreaPathInput.value = '';
      return;
    }

    normAreaPathInput.value = '';

    for (const row of NormState.rawRows) {
      const candidate = getValue(row, idxArea);
      if (candidate) {
        normAreaPathInput.value = candidate;
        break;
      }
    }
  }

  function autoFillAssigneeEmailFromData() {
    if (!normAssigneeEmailInput || !NormState.rawRows || !NormState.rawRows.length) return;

    const idxAssignee = colIndex('assigned to');
    if (idxAssignee < 0) {
      normAssigneeEmailInput.value = '';
      return;
    }

    normAssigneeEmailInput.value = '';
    for (const row of NormState.rawRows) {
      const assignee = getValue(row, idxAssignee);
      const match = assignee.match(/<\s*([^>\s]+@[^>\s]+)\s*>/);
      if (match && match[1]) {
        normAssigneeEmailInput.value = match[1].trim();
        break;
      }
    }
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
    const lower = String(name || '').toLowerCase().trim();
    return NormState.headers.findIndex(h => normalizeHeader(h) === lower);
  }

  function normalizeHeader(value) {
    return String(value || '')
      .replace(/^\uFEFF/, '')
      .toLowerCase()
      .trim();
  }

  function getValue(row, index) {
    if (index < 0 || index >= row.length) return '';
    const value = row[index];
    return value === null || value === undefined ? '' : String(value).trim();
  }

  function isCompletelyEmptyRow(row) {
    return row.every(cell => String(cell || '').trim() === '');
  }

  function reNormalize() {
    if (!NormState.rawRows) return;

    const idxID       = colIndex('id');
    const idxType     = colIndex('work item type');
    const idxTitle    = colIndex('title');
    const idxTestStep = colIndex('test step');
    const idxAction   = colIndex('step action');
    const idxExpected = colIndex('step expected');
    const idxArea     = colIndex('area path');
    const idxAssignee = colIndex('assigned to');
    const idxState    = colIndex('state');

    const overrideAreaPath = normAreaPathInput.value.trim();
    const overrideAssigneeEmail = (normAssigneeEmailInput.value || '').trim().replace(/^<|>$/g, '');
    const targetState   = normStateInput.value.trim();
    const cleanAssignee = normCleanAssignee.checked;
    const clearId       = normClearId.checked;

    const normalized = NormState.rawRows.map(row => {
      const r = row.map(cell => (cell === null || cell === undefined) ? '' : String(cell));

      if (isCompletelyEmptyRow(r)) return null;

      const workItemType = getValue(r, idxType);
      const isTestCaseRow = workItemType.toLowerCase() === 'test case';
      const hasStepNumber = getValue(r, idxTestStep) !== '';
      const isStepRow = !isTestCaseRow && hasStepNumber;

      const idValue = clearId ? '' : getValue(r, idxID);
      const titleValue = getValue(r, idxTitle);
      const testStepValue = getValue(r, idxTestStep);
      const actionValue = getValue(r, idxAction);
      const expectedValue = getValue(r, idxExpected);
      const areaValue = getValue(r, idxArea);
      const finalAreaPath = overrideAreaPath || areaValue;

      let assigneeValue = getValue(r, idxAssignee);
      if (cleanAssignee && assigneeValue) {
        assigneeValue = assigneeValue.replace(/\s*<[^>]+>/, '').trim();
      } else if (!cleanAssignee && overrideAssigneeEmail && assigneeValue) {
        const assigneeName = assigneeValue.replace(/\s*<[^>]+>/, '').trim();
        assigneeValue = assigneeName ? `${assigneeName} <${overrideAssigneeEmail}>` : '';
      }
      if (assigneeValue.toLowerCase() === 'design') {
        assigneeValue = '';
      }

      let stateValue = getValue(r, idxState);
      if (targetState && isTestCaseRow) {
        stateValue = targetState;
      }

      const out = Array(TARGET_HEADERS.length).fill('');

      if (isTestCaseRow) {
        out[0] = idValue;
        out[1] = 'Test Case';
        out[2] = titleValue;
        out[6] = finalAreaPath;
        out[7] = assigneeValue;
        out[8] = stateValue;
        return out;
      }

      if (isStepRow) {
        out[3] = testStepValue;
        out[4] = actionValue;
        out[5] = expectedValue;
        return out;
      }

      // Fallback: preserve known fields in strict Azure schema.
      out[0] = idValue;
      out[1] = workItemType;
      out[2] = titleValue;
      out[3] = testStepValue;
      out[4] = actionValue;
      out[5] = expectedValue;
      out[6] = finalAreaPath;
      out[7] = assigneeValue;
      out[8] = stateValue;
      return isCompletelyEmptyRow(out) ? null : out;
    }).filter(Boolean);

    NormState.normalized = normalized;

    // Update stats
    const testCases = normalized.filter(r => {
      const type = r[1] || '';
      return type.trim().toLowerCase() === 'test case';
    }).length;

    const stepRows = normalized.filter(r => {
      const step = r[1] || '';
      const testStep = r[3] || '';
      return step.trim() === '' && testStep !== '';
    }).length;

    normStatFile.textContent  = NormState.fileName;
    normStatCases.textContent = `${testCases} Test Case${testCases !== 1 ? 's' : ''}`;
    normStatSteps.textContent = `${stepRows} paso${stepRows !== 1 ? 's' : ''}`;

    renderPreview();
  }

  // ─── Preview ─────────────────────────────────────────────────────
  function renderPreview() {
    const headers = TARGET_HEADERS;
    const rows    = NormState.normalized;

    ensureColumnWidths();
    applyColumnWidths();

    // Head
    normPreviewHead.innerHTML = '<tr>' +
      headers.map((h, index) => `<th data-col-index="${index}"><span class="norm-th-label">${escHtml(h)}</span><span class="norm-col-resizer" data-col-index="${index}" title="Arrastra para ajustar ancho"></span></th>`).join('') +
      '</tr>';

    // Body
    const idxType = 1;
    normPreviewBody.innerHTML = rows.map((row, rowIndex) => {
      const isTestCase = (row[idxType] || '').trim().toLowerCase() === 'test case';
      const cls = isTestCase ? ' class="norm-row-testcase"' : ' class="norm-row-step"';
      return `<tr${cls}>` +
        row.map((cell, colIndex) => `<td contenteditable="true" spellcheck="false" data-row-index="${rowIndex}" data-col-index="${colIndex}">${escHtml(cell)}</td>`).join('') +
        '</tr>';
    }).join('');

    bindPreviewEditing();
    bindColumnResizers();
  }

  function bindPreviewEditing() {
    normPreviewBody.querySelectorAll('td[contenteditable="true"]').forEach(cell => {
      const syncCell = () => {
        const rowIndex = Number(cell.dataset.rowIndex);
        const colIndex = Number(cell.dataset.colIndex);
        if (!Number.isInteger(rowIndex) || !Number.isInteger(colIndex)) return;
        if (!NormState.normalized || !NormState.normalized[rowIndex]) return;

        let value = String(cell.textContent || '').trim();

        // Assigned To (col 7) cannot be 'Design'. Azure expects empty or valid user.
        if (colIndex === 7 && value.toLowerCase() === 'design') {
          value = '';
          cell.textContent = '';
          showNormToast('En "Assigned To", el valor "Design" no es válido. Se dejó vacío.', 'info');
        }

        if (colIndex === 7 && normCleanAssignee.checked && value) {
          value = value.replace(/\s*<[^>]+>/, '').trim();
          cell.textContent = value;
        }

        NormState.normalized[rowIndex][colIndex] = value;
      };

      cell.addEventListener('input', syncCell);
      cell.addEventListener('blur', syncCell);
    });
  }

  function bindColumnResizers() {
    normPreviewHead.querySelectorAll('.norm-col-resizer').forEach(handle => {
      handle.addEventListener('mousedown', event => {
        event.preventDefault();
        event.stopPropagation();

        const colIndex = Number(handle.dataset.colIndex);
        if (!Number.isInteger(colIndex)) return;

        const startX = event.clientX;
        const startWidth = Number(NormState.columnWidths[colIndex]) || DEFAULT_COLUMN_WIDTHS[colIndex] || 180;

        const onMouseMove = moveEvent => {
          const delta = moveEvent.clientX - startX;
          const newWidth = Math.max(70, Math.min(1000, startWidth + delta));
          NormState.columnWidths[colIndex] = newWidth;
          applyColumnWidths();
        };

        const onMouseUp = () => {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  function ensureColumnWidths() {
    if (!Array.isArray(NormState.columnWidths) || NormState.columnWidths.length !== TARGET_HEADERS.length) {
      NormState.columnWidths = [...DEFAULT_COLUMN_WIDTHS];
    }
  }

  function applyColumnWidths() {
    if (!normPreviewTable) return;

    const existingColgroup = normPreviewTable.querySelector('colgroup');
    if (existingColgroup) existingColgroup.remove();

    const colgroup = document.createElement('colgroup');
    NormState.columnWidths.forEach(width => {
      const col = document.createElement('col');
      col.style.width = `${width}px`;
      colgroup.appendChild(col);
    });

    normPreviewTable.insertBefore(colgroup, normPreviewHead);
  }

  function validateNormalizedRows() {
    const issues = [];
    if (!Array.isArray(NormState.normalized)) return issues;

    NormState.normalized.forEach((row, index) => {
      const dataRowNumber = index + 1;
      const workItemType = String(row[1] || '').trim().toLowerCase();
      const title = String(row[2] || '').trim();
      const testStep = String(row[3] || '').trim();
      const stepAction = String(row[4] || '').trim();
      const stepExpected = String(row[5] || '').trim();
      const areaPath = String(row[6] || '').trim();
      const assignedTo = String(row[7] || '').trim().toLowerCase();

      const isTestCase = workItemType === 'test case';
      const isStepRow = workItemType === '' && testStep !== '';

      if (assignedTo === 'design') {
        issues.push({ row: dataRowNumber, message: 'Assigned To no puede ser "Design".' });
      }
      if (isTestCase && !title) {
        issues.push({ row: dataRowNumber, message: 'Test Case sin título (Title).' });
      }
      if (isTestCase && !areaPath) {
        issues.push({ row: dataRowNumber, message: 'Test Case sin Area Path.' });
      }
      if (isStepRow && !stepAction && !stepExpected) {
        issues.push({ row: dataRowNumber, message: 'Paso sin Step Action y sin Step Expected.' });
      }
    });

    return issues;
  }

  // ─── Download ─────────────────────────────────────────────────────
  function downloadNormalized() {
    if (!NormState.normalized) return;

    const issues = validateNormalizedRows();
    if (issues.length) {
      const maxRowsToShow = 25;
      const detailLines = issues
        .slice(0, maxRowsToShow)
        .map(issue => `- Fila ${issue.row}: ${issue.message}`)
        .join('\n');
      const remaining = issues.length - maxRowsToShow;
      const suffix = remaining > 0 ? `\n... y ${remaining} problema(s) más.` : '';

      showNormToast(`Se detectaron ${issues.length} fila(s) con valores inválidos.`, 'error');
      window.alert(`No se puede descargar todavía. Corrige estos datos antes de exportar:\n\n${detailLines}${suffix}`);
      return;
    }

    const allRows = [TARGET_HEADERS, ...NormState.normalized];

    const csv = Papa.unparse(allRows, {
      quotes: true,
      quoteChar: '"',
      escapeChar: '"',
      delimiter: ',',
      newline: '\r\n',
      skipEmptyLines: false
    });

    const blob = new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8;' });
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
    NormState.columnWidths = [...DEFAULT_COLUMN_WIDTHS];

    normFileInput.value = '';
    normOptions.style.display       = 'none';
    normStatsBar.style.display      = 'none';
    normPreviewWrap.style.display   = 'none';
    normEmptyState.style.display    = '';
    normPreviewHead.innerHTML       = '';
    normPreviewBody.innerHTML       = '';
    normAreaPathInput.value         = '';
    normAssigneeEmailInput.value    = '';
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
