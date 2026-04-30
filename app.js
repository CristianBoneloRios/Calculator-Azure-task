/* ═══════════════════════════════════════════════════════════════
   AZURE TASK HOURS CALCULATOR — app.js
   Supports: CSV (PapaParse) · Excel (SheetJS)
   Detects Azure DevOps hour columns automatically
   ═══════════════════════════════════════════════════════════════ */

'use strict';

// ─────────────────────────────────────────
// STATE
// ─────────────────────────────────────────
const AppState = {
  files: []   // { id, name, ext, rows, headers, colMap, summary }
};

const ABOUT_PHOTO_STORAGE_KEY = 'about_profile_photo_dataurl';
const ABOUT_PHOTO_LOCK_KEY = 'about_profile_photo_locked';
const ABOUT_PHOTO_CHANGE_PASSWORD = '580622';
const SHARED_PROFILE_PHOTO_CANDIDATES = [
  'assets/profile-photo.jpg',
  'assets/profile-photo.jpeg',
  'assets/profile-photo.png',
  'assets/profile-photo.webp',
  'assets/profile-photo.svg'
];

const GITHUB_REPO_OWNER = 'CristianBoneloRios';
const GITHUB_REPO_NAME = 'Calculator-Azure-task';
const GITHUB_REPO_BRANCH = 'main';
const GITHUB_PROFILE_PHOTO_PATH = 'assets/profile-photo.png';

// ─────────────────────────────────────────
// AZURE DEVOPS COLUMN MAPPING (EN + ES)
// ─────────────────────────────────────────
const COL_CANDIDATES = {
  id:               ['id', 'work item id', 'id de elemento de trabajo', 'workitemid'],
  title:            ['title', 'título', 'titulo', 'name', 'nombre'],
  type:             ['work item type', 'tipo de elemento de trabajo', 'type', 'tipo', 'item type'],
  assignedTo:       ['assigned to', 'asignado a', 'assignee', 'asignatario'],
  state:            ['state', 'estado', 'status', 'estatus'],
  completedWork:    ['completed work', 'trabajo completado', 'horas completadas', 'actual work',
                     'horas reales', 'completed hours', 'trabajo real'],
  originalEstimate: ['original estimate', 'estimación original', 'estimacion original',
                     'estimated work', 'estimate', 'estimado', 'horas estimadas',
                     'story points', 'story point', 'puntos de historia'],
  remainingWork:    ['remaining work', 'trabajo restante', 'remaining', 'restante',
                     'horas restantes', 'work remaining'],
  tags:             ['tags', 'etiquetas', 'tag'],
  priority:         ['priority', 'prioridad'],
  iterationPath:    ['iteration path', 'ruta de iteración', 'ruta de iteracion', 'sprint', 'iteration'],
  areaPath:         ['area path', 'ruta de área', 'ruta de area', 'area'],
};

// ─────────────────────────────────────────
// DOM REFS
// ─────────────────────────────────────────
const DOM = {};

document.addEventListener('DOMContentLoaded', () => {
  DOM.sidebarToggleBtn   = document.getElementById('sidebarToggleBtn');
  DOM.sidebar            = document.getElementById('sidebar');
  DOM.sidebarOverlay     = document.getElementById('sidebarOverlay');
  DOM.mainContent        = document.getElementById('mainContent');
  DOM.uploadZone         = document.getElementById('uploadZone');
  DOM.fileInput          = document.getElementById('fileInput');
  DOM.selectFilesBtn     = document.getElementById('selectFilesBtn');
  DOM.fileHistory        = document.getElementById('fileHistory');
  DOM.fileResultsContainer = document.getElementById('fileResultsContainer');
  DOM.clearAllBtn        = document.getElementById('clearAllBtn');
  DOM.statFiles          = document.getElementById('statFiles');
  DOM.statTasks          = document.getElementById('statTasks');
  DOM.statCompleted      = document.getElementById('statCompleted');
  DOM.statEstimate       = document.getElementById('statEstimate');
  DOM.statRemaining      = document.getElementById('statRemaining');
  DOM.statProgress       = document.getElementById('statProgress');
  DOM.summarySection     = document.getElementById('summary-section');
  DOM.resultsSection     = document.getElementById('results-section');
  DOM.resultsEmptyState  = document.getElementById('resultsEmptyState');
  DOM.toastContainer     = document.getElementById('toastContainer');
  DOM.aboutMeBtn           = document.getElementById('aboutMeBtn');
  DOM.aboutModal           = document.getElementById('aboutModal');
  DOM.aboutModalBackdrop   = document.getElementById('aboutModalBackdrop');
  DOM.aboutModalClose      = document.getElementById('aboutModalClose');
  DOM.aboutPhotoInput      = document.getElementById('aboutPhotoInput');
  DOM.aboutModalAvatarImg  = document.getElementById('aboutModalAvatarImg');
  DOM.aboutModalAvatarIcon = document.getElementById('aboutModalAvatarIcon');
  DOM.aboutMeAvatarThumb   = document.getElementById('aboutMeAvatarThumb');
  DOM.aboutMeAvatarIcon    = document.getElementById('aboutMeAvatarIcon');
  DOM.photoGuardBackdrop   = document.getElementById('photoGuardBackdrop');
  DOM.photoGuardModal      = document.getElementById('photoGuardModal');
  DOM.photoGuardInput      = document.getElementById('photoGuardInput');
  DOM.photoGuardError      = document.getElementById('photoGuardError');
  DOM.photoGuardCancelBtn  = document.getElementById('photoGuardCancelBtn');
  DOM.photoGuardConfirmBtn = document.getElementById('photoGuardConfirmBtn');
  DOM.publishPhotoBtn      = document.getElementById('publishPhotoBtn');
  DOM.githubPublishBackdrop = document.getElementById('githubPublishBackdrop');
  DOM.githubPublishModal    = document.getElementById('githubPublishModal');
  DOM.githubTokenInput      = document.getElementById('githubTokenInput');
  DOM.githubPublishError    = document.getElementById('githubPublishError');
  DOM.githubPublishCancelBtn = document.getElementById('githubPublishCancelBtn');
  DOM.githubPublishConfirmBtn = document.getElementById('githubPublishConfirmBtn');

  const hasLocalPhoto = restoreAboutPhotoFromStorage();
  if (!hasLocalPhoto) {
    loadSharedAboutPhoto();
  }
  initEvents();
  updateGlobalStats();
  updateFileHistory();
  updateResultsEmptyState();
});

// ─────────────────────────────────────────
// EVENT INIT
// ─────────────────────────────────────────
function initEvents() {
  // Sidebar toggle
  DOM.sidebarToggleBtn.addEventListener('click', toggleSidebar);
  DOM.sidebarOverlay.addEventListener('click', closeSidebar);

  // About Me modal
  DOM.aboutMeBtn.addEventListener('click', openAboutModal);
  DOM.aboutModalClose.addEventListener('click', closeAboutModal);
  DOM.aboutModalBackdrop.addEventListener('click', closeAboutModal);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeAboutModal(); });

  // Photo upload
  DOM.aboutPhotoInput.addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;

    // After the first photo is stored, changing it requires password.
    if (isAboutPhotoLocked()) {
      const authorized = await requestPhotoChangeAuthorization();
      if (!authorized) {
        e.target.value = '';
        return;
      }
    }

    const reader = new FileReader();
    reader.onload = ev => {
      const src = ev.target.result;

      applyAboutPhoto(src);
      persistAboutPhoto(src);
      showToast('Foto guardada localmente en este navegador.', 'success');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  });

  // Publish profile photo globally (GitHub repo)
  DOM.publishPhotoBtn.addEventListener('click', handlePublishPhotoToGithub);

  // File input
  DOM.selectFilesBtn.addEventListener('click', () => DOM.fileInput.click());
  DOM.fileInput.addEventListener('change', e => handleFiles(e.target.files));

  // Drag & drop
  DOM.uploadZone.addEventListener('dragover',  e => { e.preventDefault(); DOM.uploadZone.classList.add('drag-over'); });
  DOM.uploadZone.addEventListener('dragleave', e => { if (!DOM.uploadZone.contains(e.relatedTarget)) DOM.uploadZone.classList.remove('drag-over'); });
  DOM.uploadZone.addEventListener('drop',      e => { e.preventDefault(); DOM.uploadZone.classList.remove('drag-over'); handleFiles(e.dataTransfer.files); });
  DOM.uploadZone.addEventListener('click',     e => { if (e.target === DOM.uploadZone || e.target.closest('.upload-zone-content')) {} });

  // Nav sidebar items
  const navItems = document.querySelectorAll('.nav-item[data-section]');
  navItems.forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      const target = document.getElementById(item.dataset.section);
      if (target) { target.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      setActiveNavItem(item.dataset.section);
    });
  });

  initSectionObserver();

  // Clear all
  DOM.clearAllBtn.addEventListener('click', clearAll);
}

function setActiveNavItem(sectionId) {
  document.querySelectorAll('.nav-item[data-section]').forEach(item => {
    item.classList.toggle('active', item.dataset.section === sectionId);
  });
}

function initSectionObserver() {
  const sectionIds = ['upload-section', 'summary-section', 'results-section'];
  const sections = sectionIds
    .map(id => document.getElementById(id))
    .filter(Boolean);

  if (!sections.length || typeof IntersectionObserver === 'undefined') return;

  const observer = new IntersectionObserver(entries => {
    const visible = entries
      .filter(entry => entry.isIntersecting)
      .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];

    if (visible && visible.target && visible.target.id) {
      setActiveNavItem(visible.target.id);
    }
  }, {
    root: null,
    threshold: [0.25, 0.45, 0.65],
    rootMargin: '-20% 0px -55% 0px'
  });

  sections.forEach(section => observer.observe(section));
}

// ─────────────────────────────────────────
// SIDEBAR TOGGLE

// ─────────────────────────────────────────
// ABOUT ME MODAL
// ─────────────────────────────────────────
function openAboutModal() {
  DOM.aboutModalBackdrop.classList.add('active');
  DOM.aboutModal.classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeAboutModal() {
  DOM.aboutModalBackdrop.classList.remove('active');
  DOM.aboutModal.classList.remove('active');
  document.body.style.overflow = '';
}

function requestPhotoChangeAuthorization() {
  return new Promise(resolve => {
    const onConfirm = () => {
      const password = DOM.photoGuardInput.value.trim();
      if (password === ABOUT_PHOTO_CHANGE_PASSWORD) {
        cleanup();
        resolve(true);
        return;
      }

      DOM.photoGuardError.textContent = 'Clave incorrecta. Intenta de nuevo.';
      DOM.photoGuardModal.classList.remove('shake');
      void DOM.photoGuardModal.offsetWidth;
      DOM.photoGuardModal.classList.add('shake');
      DOM.photoGuardInput.focus();
      DOM.photoGuardInput.select();
    };

    const onCancel = () => {
      cleanup();
      showToast('Cambio de foto cancelado.', 'info');
      resolve(false);
    };

    const onKeyDown = e => {
      if (e.key === 'Enter') onConfirm();
      if (e.key === 'Escape') onCancel();
    };

    const cleanup = () => {
      DOM.photoGuardConfirmBtn.removeEventListener('click', onConfirm);
      DOM.photoGuardCancelBtn.removeEventListener('click', onCancel);
      DOM.photoGuardBackdrop.removeEventListener('click', onCancel);
      document.removeEventListener('keydown', onKeyDown);

      DOM.photoGuardBackdrop.classList.remove('active');
      DOM.photoGuardModal.classList.remove('active');
      DOM.photoGuardModal.classList.remove('shake');
      DOM.photoGuardError.textContent = '';
      DOM.photoGuardInput.value = '';
      document.body.style.overflow = '';
    };

    DOM.photoGuardConfirmBtn.addEventListener('click', onConfirm);
    DOM.photoGuardCancelBtn.addEventListener('click', onCancel);
    DOM.photoGuardBackdrop.addEventListener('click', onCancel);
    document.addEventListener('keydown', onKeyDown);

    DOM.photoGuardBackdrop.classList.add('active');
    DOM.photoGuardModal.classList.add('active');
    document.body.style.overflow = 'hidden';
    DOM.photoGuardInput.focus();
  });
}

function applyAboutPhoto(src) {
  DOM.aboutModalAvatarImg.src = src;
  DOM.aboutModalAvatarImg.style.display = 'block';
  DOM.aboutModalAvatarIcon.style.display = 'none';

  DOM.aboutMeAvatarThumb.src = src;
  DOM.aboutMeAvatarThumb.style.display = 'block';
  DOM.aboutMeAvatarIcon.style.display = 'none';
}

function persistAboutPhoto(src) {
  localStorage.setItem(ABOUT_PHOTO_STORAGE_KEY, src);
  localStorage.setItem(ABOUT_PHOTO_LOCK_KEY, 'true');
}

function restoreAboutPhotoFromStorage() {
  const storedPhoto = localStorage.getItem(ABOUT_PHOTO_STORAGE_KEY);
  if (!storedPhoto) return false;
  applyAboutPhoto(storedPhoto);

  // Keep compatibility if photo existed before lock flag was created.
  if (!localStorage.getItem(ABOUT_PHOTO_LOCK_KEY)) {
    localStorage.setItem(ABOUT_PHOTO_LOCK_KEY, 'true');
  }
  return true;
}

function isAboutPhotoLocked() {
  return localStorage.getItem(ABOUT_PHOTO_LOCK_KEY) === 'true';
}

function loadSharedAboutPhoto() {
  const tryLoad = index => {
    if (index >= SHARED_PROFILE_PHOTO_CANDIDATES.length) return;
    const candidate = SHARED_PROFILE_PHOTO_CANDIDATES[index];
    const img = new Image();
    img.onload = () => applyAboutPhoto(candidate);
    img.onerror = () => tryLoad(index + 1);
    img.src = candidate;
  };

  tryLoad(0);
}

async function handlePublishPhotoToGithub() {
  const photoDataUrl = localStorage.getItem(ABOUT_PHOTO_STORAGE_KEY);
  if (!photoDataUrl || !photoDataUrl.startsWith('data:image/')) {
    showToast('Primero sube una foto desde este dispositivo para poder publicarla.', 'error');
    return;
  }

  const token = await requestGithubToken();
  if (!token) return;

  try {
    DOM.publishPhotoBtn.disabled = true;
    DOM.publishPhotoBtn.classList.add('pulse');

    // Normalize to PNG so the repo path stays fixed regardless of original filename/type.
    const pngDataUrl = await convertImageDataUrlToPng(photoDataUrl);
    const contentBase64 = pngDataUrl.split(',')[1];

    const sha = await getRepoFileSha(token, GITHUB_PROFILE_PHOTO_PATH);
    const body = {
      message: 'Update shared profile photo from web app',
      content: contentBase64,
      branch: GITHUB_REPO_BRANCH
    };
    if (sha) body.sha = sha;

    const response = await fetch(`https://api.github.com/repos/${GITHUB_REPO_OWNER}/${GITHUB_REPO_NAME}/contents/${GITHUB_PROFILE_PHOTO_PATH}`, {
      method: 'PUT',
      headers: {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      const msg = err && err.message ? err.message : `Error HTTP ${response.status}`;
      throw new Error(msg);
    }

    showToast('Foto publicada en GitHub. Quedara visible para todos tras el deploy.', 'success');
  } catch (error) {
    showToast(`No se pudo publicar la foto: ${error.message}`, 'error');
  } finally {
    DOM.publishPhotoBtn.disabled = false;
    DOM.publishPhotoBtn.classList.remove('pulse');
  }
}

async function getRepoFileSha(token, path) {
  const response = await fetch(`https://api.github.com/repos/${GITHUB_REPO_OWNER}/${GITHUB_REPO_NAME}/contents/${path}?ref=${GITHUB_REPO_BRANCH}`, {
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${token}`
    }
  });

  if (response.status === 404) return null;
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    const msg = err && err.message ? err.message : `Error HTTP ${response.status}`;
    throw new Error(msg);
  }

  const data = await response.json();
  return data.sha || null;
}

function requestGithubToken() {
  return new Promise(resolve => {
    const onConfirm = () => {
      const token = DOM.githubTokenInput.value.trim();
      if (!token) {
        DOM.githubPublishError.textContent = 'Debes ingresar un token de GitHub.';
        DOM.githubTokenInput.focus();
        return;
      }
      cleanup();
      resolve(token);
    };

    const onCancel = () => {
      cleanup();
      resolve(null);
    };

    const onKeyDown = e => {
      if (e.key === 'Enter') onConfirm();
      if (e.key === 'Escape') onCancel();
    };

    const cleanup = () => {
      DOM.githubPublishConfirmBtn.removeEventListener('click', onConfirm);
      DOM.githubPublishCancelBtn.removeEventListener('click', onCancel);
      DOM.githubPublishBackdrop.removeEventListener('click', onCancel);
      document.removeEventListener('keydown', onKeyDown);

      DOM.githubPublishBackdrop.classList.remove('active');
      DOM.githubPublishModal.classList.remove('active');
      DOM.githubPublishError.textContent = '';
      DOM.githubTokenInput.value = '';
      document.body.style.overflow = '';
    };

    DOM.githubPublishConfirmBtn.addEventListener('click', onConfirm);
    DOM.githubPublishCancelBtn.addEventListener('click', onCancel);
    DOM.githubPublishBackdrop.addEventListener('click', onCancel);
    document.addEventListener('keydown', onKeyDown);

    DOM.githubPublishBackdrop.classList.add('active');
    DOM.githubPublishModal.classList.add('active');
    document.body.style.overflow = 'hidden';
    DOM.githubTokenInput.focus();
  });
}

function convertImageDataUrlToPng(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        reject(new Error('No se pudo convertir la imagen.'));
        return;
      }

      ctx.drawImage(img, 0, 0);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => reject(new Error('Formato de imagen invalido.'));
    img.src = dataUrl;
  });
}

// ─────────────────────────────────────────
function toggleSidebar() {
  const isMobile = window.innerWidth <= 768;
  if (isMobile) {
    DOM.sidebar.classList.toggle('mobile-open');
    DOM.sidebarOverlay.classList.toggle('visible');
  } else {
    DOM.sidebar.classList.toggle('collapsed');
  }
}

function closeSidebar() {
  DOM.sidebar.classList.remove('mobile-open');
  DOM.sidebarOverlay.classList.remove('visible');
}

// ─────────────────────────────────────────
// FILE HANDLING
// ─────────────────────────────────────────
function handleFiles(fileList) {
  if (!fileList || !fileList.length) return;

  Array.from(fileList).forEach(file => {
    const name = file.name;
    const ext  = name.split('.').pop().toLowerCase();

    if (!['csv', 'xlsx', 'xls'].includes(ext)) {
      showToast(`Formato no soportado: ${name}`, 'error');
      return;
    }

    // Avoid duplicate filenames
    if (AppState.files.some(f => f.name === name)) {
      showToast(`El archivo "${name}" ya fue cargado.`, 'info');
      return;
    }

    if (ext === 'csv') {
      parseCSV(file);
    } else {
      parseExcel(file);
    }
  });

  // Reset input so the same file can be uploaded again after clearing
  DOM.fileInput.value = '';
}

// ─────────────────────────────────────────
// CSV PARSING (PapaParse)
// ─────────────────────────────────────────
function parseCSV(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: 'greedy',
    dynamicTyping: false,
    complete(results) {
      if (!results.data || !results.data.length) {
        showToast(`El archivo "${file.name}" está vacío o tiene un formato inválido.`, 'error');
        return;
      }
      processData(file.name, 'csv', results.meta.fields || [], results.data);
      showToast(`"${file.name}" cargado exitosamente.`, 'success');
    },
    error(err) {
      showToast(`Error al leer "${file.name}": ${err.message}`, 'error');
    }
  });
}

// ─────────────────────────────────────────
// EXCEL PARSING (SheetJS)
// ─────────────────────────────────────────
function parseExcel(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      if (!raw.length) {
        showToast(`El archivo "${file.name}" está vacío.`, 'error');
        return;
      }

      // First row = headers
      const headers = raw[0].map(h => String(h).trim());
      const rows = raw.slice(1)
        .filter(r => r.some(c => String(c).trim() !== ''))
        .map(r => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? String(r[i]) : ''; });
          return obj;
        });

      processData(file.name, file.name.endsWith('.xls') ? 'xls' : 'xlsx', headers, rows);
      showToast(`"${file.name}" cargado exitosamente.`, 'success');
    } catch (err) {
      showToast(`Error al leer "${file.name}": ${err.message}`, 'error');
    }
  };
  reader.onerror = () => showToast(`Error al leer el archivo "${file.name}".`, 'error');
  reader.readAsArrayBuffer(file);
}

// ─────────────────────────────────────────
// PROCESS DATA
// ─────────────────────────────────────────
function processData(name, ext, headers, rows) {
  const colMap   = identifyColumns(headers);
  const summary  = calculateSummary(rows, colMap);

  const fileEntry = {
    id:      Date.now() + Math.random(),
    name,
    ext,
    headers,
    rows,
    colMap,
    summary
  };

  AppState.files.push(fileEntry);
  renderFileCard(fileEntry);
  updateGlobalStats();
  updateFileHistory();
  updateResultsEmptyState();
}

// ─────────────────────────────────────────
// COLUMN IDENTIFICATION
// ─────────────────────────────────────────
function identifyColumns(headers) {
  const norm = headers.map(h => String(h).toLowerCase().trim().replace(/\s+/g, ' '));
  const result = {};

  for (const [key, candidates] of Object.entries(COL_CANDIDATES)) {
    for (const candidate of candidates) {
      const idx = norm.indexOf(candidate);
      if (idx !== -1) {
        result[key] = headers[idx]; // store original header name
        break;
      }
    }
  }
  return result;
}

// ─────────────────────────────────────────
// HOURS PARSING
// ─────────────────────────────────────────
function parseHours(value) {
  if (value === null || value === undefined || value === '') return 0;
  const str = String(value).trim()
    .replace(/,/g, '.')
    .replace(/[hH]oras?/g, '')
    .trim();
  const n = parseFloat(str);
  return isNaN(n) ? 0 : Math.max(0, n);
}

// ─────────────────────────────────────────
// CALCULATE SUMMARY
// ─────────────────────────────────────────
function calculateSummary(rows, colMap) {
  let completedWork    = 0;
  let originalEstimate = 0;
  let remainingWork    = 0;

  rows.forEach(row => {
    if (colMap.completedWork)    completedWork    += parseHours(row[colMap.completedWork]);
    if (colMap.originalEstimate) originalEstimate += parseHours(row[colMap.originalEstimate]);
    if (colMap.remainingWork)    remainingWork    += parseHours(row[colMap.remainingWork]);
  });

  const progress = originalEstimate > 0
    ? Math.min(100, Math.round((completedWork / originalEstimate) * 100))
    : 0;

  return {
    totalRows:        rows.length,
    completedWork:    round2(completedWork),
    originalEstimate: round2(originalEstimate),
    remainingWork:    round2(remainingWork),
    progress,
    hasHourCols: !!(colMap.completedWork || colMap.originalEstimate || colMap.remainingWork)
  };
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function fmtHours(n) {
  if (n === 0) return '0h';
  const h = Math.floor(n);
  const m = Math.round((n - h) * 60);
  if (m === 0) return `${h}h`;
  return `${h}h ${m}m`;
}

// ─────────────────────────────────────────
// RENDER FILE CARD
// ─────────────────────────────────────────
function renderFileCard(fileEntry) {
  const { id, name, ext, rows, headers, colMap, summary } = fileEntry;
  const isExcel = ext === 'xlsx' || ext === 'xls';
  const iconClass = isExcel ? 'xlsx' : 'csv';
  const iconSvg   = isExcel
    ? '<i class="fas fa-file-excel"></i>'
    : '<i class="fas fa-file-csv"></i>';

  const card = document.createElement('div');
  card.className = 'file-result-card';
  card.dataset.fileId = id;

  // ── Header
  const headerDiv = document.createElement('div');
  headerDiv.className = 'file-card-header open';

  const pillsHTML = buildPillsHTML(summary, colMap);

  headerDiv.innerHTML = `
    <div class="file-card-left">
      <div class="file-card-icon ${iconClass}">${iconSvg}</div>
      <div class="file-card-info">
        <div class="file-card-name" title="${escHtml(name)}">${escHtml(name)}</div>
        <div class="file-card-meta">${summary.totalRows} filas · ${headers.length} columnas</div>
      </div>
    </div>
    <div class="file-card-pills">${pillsHTML}</div>
    <div class="file-card-toggle open"><i class="fas fa-chevron-down"></i></div>
  `;

  // ── Body
  const bodyDiv = document.createElement('div');
  bodyDiv.className = 'file-card-body open';

  // Hours summary bar
  const hoursBar = buildHoursSummaryBar(summary, colMap);

  // Progress bar
  const progressBar = buildProgressBar(summary);

  // Table
  const table = buildTable(rows, headers, colMap);

  // Warning if no hour cols detected
  const warning = !summary.hasHourCols
    ? `<div class="no-hours-warning"><i class="fas fa-exclamation-triangle"></i>
        No se detectaron columnas de horas estándar de Azure DevOps en este archivo.
        Se muestran todos los datos disponibles.</div>`
    : '';

  bodyDiv.appendChild(hoursBar);
  bodyDiv.appendChild(progressBar);
  if (warning) bodyDiv.insertAdjacentHTML('beforeend', warning);
  bodyDiv.appendChild(table);

  card.appendChild(headerDiv);
  card.appendChild(bodyDiv);

  // Toggle collapse on header click
  headerDiv.addEventListener('click', () => {
    const isOpen = bodyDiv.classList.contains('open');
    bodyDiv.classList.toggle('open', !isOpen);
    headerDiv.classList.toggle('open', !isOpen);
    headerDiv.querySelector('.file-card-toggle').classList.toggle('open', !isOpen);
  });

  DOM.fileResultsContainer.prepend(card);
}

function buildPillsHTML(summary, colMap) {
  const pills = [];
  pills.push(`<span class="pill pill-blue"><i class="fas fa-tasks"></i>${summary.totalRows} tareas</span>`);
  if (colMap.completedWork)
    pills.push(`<span class="pill pill-green"><i class="fas fa-check-circle"></i>${fmtHours(summary.completedWork)} completadas</span>`);
  if (colMap.originalEstimate)
    pills.push(`<span class="pill pill-orange"><i class="fas fa-clock"></i>${fmtHours(summary.originalEstimate)} estimadas</span>`);
  if (colMap.remainingWork)
    pills.push(`<span class="pill pill-purple"><i class="fas fa-hourglass-half"></i>${fmtHours(summary.remainingWork)} restantes</span>`);
  return pills.join('');
}

function buildHoursSummaryBar(summary, colMap) {
  const bar = document.createElement('div');
  bar.className = 'hours-summary-bar';

  const cards = [
    { label: 'Total Tareas',         val: summary.totalRows,                         icon: 'fas fa-tasks',          color: 'rgba(59,130,246,.2)',   text: '#60a5fa' },
    { label: 'Horas Completadas',    val: fmtHours(summary.completedWork),            icon: 'fas fa-check-circle',   color: 'rgba(16,185,129,.2)',   text: '#34d399', show: !!colMap.completedWork },
    { label: 'Estimación Original',  val: fmtHours(summary.originalEstimate),         icon: 'fas fa-clock',          color: 'rgba(245,158,11,.2)',   text: '#fcd34d', show: !!colMap.originalEstimate },
    { label: 'Trabajo Restante',     val: fmtHours(summary.remainingWork),            icon: 'fas fa-hourglass-half', color: 'rgba(139,92,246,.2)',   text: '#c4b5fd', show: !!colMap.remainingWork },
    { label: 'Progreso',             val: summary.progress + '%',                     icon: 'fas fa-percentage',     color: 'rgba(236,72,153,.2)',   text: '#f9a8d4', show: summary.originalEstimate > 0 },
  ];

  cards.forEach(c => {
    if (c.show === false) return;
    const el = document.createElement('div');
    el.className = 'hour-stat';
    el.innerHTML = `
      <div class="hour-stat-icon" style="background:${c.color}; color:${c.text}">
        <i class="${c.icon}"></i>
      </div>
      <div class="hour-stat-text">
        <span class="hour-stat-val" style="color:${c.text}">${c.val}</span>
        <span class="hour-stat-lbl">${c.label}</span>
      </div>`;
    bar.appendChild(el);
  });

  return bar;
}

function buildProgressBar(summary) {
  const div = document.createElement('div');
  if (!summary.hasHourCols || summary.originalEstimate === 0) {
    div.style.display = 'none';
    return div;
  }

  div.className = 'progress-bar-wrap';
  const pct = summary.progress;
  const color = pct >= 100 ? '#34d399' : pct >= 75 ? '#60a5fa' : pct >= 50 ? '#fcd34d' : '#f87171';

  div.innerHTML = `
    <div class="progress-info">
      <span>Progreso de completado</span>
      <span style="color:${color}; font-weight:700">${pct}%</span>
    </div>
    <div class="progress-track">
      <div class="progress-fill" style="width:0%; background: linear-gradient(90deg, ${color}, ${color}cc)"></div>
    </div>`;

  // Animate after render
  setTimeout(() => {
    const fill = div.querySelector('.progress-fill');
    if (fill) fill.style.width = `${pct}%`;
  }, 80);

  return div;
}

function buildTable(rows, headers, colMap) {
  const hourCols = new Set([colMap.completedWork, colMap.originalEstimate, colMap.remainingWork].filter(Boolean));

  const container = document.createElement('div');
  container.className = 'table-container';

  const table = document.createElement('table');
  table.className = 'data-table';

  // THEAD
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    if (hourCols.has(h)) th.className = 'col-hour';
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);

  // TBODY
  const tbody = document.createElement('tbody');
  rows.forEach(row => {
    const tr = document.createElement('tr');
    headers.forEach(h => {
      const td = document.createElement('td');
      const rawVal = row[h] !== undefined ? String(row[h]) : '';

      if (hourCols.has(h)) {
        td.className = 'col-hour';
        const n = parseHours(rawVal);
        td.textContent = n > 0 ? fmtHours(n) : '—';
      } else if (h === colMap.state) {
        td.innerHTML = renderStateBadge(rawVal);
      } else {
        td.textContent = rawVal;
        td.title = rawVal.length > 40 ? rawVal : '';
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  // TFOOT — totals for hour columns
  if (hourCols.size > 0) {
    const tfoot = document.createElement('tfoot');
    const fr = document.createElement('tr');

    let first = true;
    headers.forEach(h => {
      const td = document.createElement('td');
      if (first) {
        td.textContent = 'TOTAL';
        td.style.fontWeight = '700';
        td.style.color = 'var(--text-1)';
        first = false;
      } else if (hourCols.has(h)) {
        const total = rows.reduce((acc, row) => acc + parseHours(row[h]), 0);
        td.textContent = total > 0 ? fmtHours(round2(total)) : '—';
        td.className = 'col-hour';
      } else {
        td.textContent = '';
      }
      fr.appendChild(td);
    });

    tfoot.appendChild(fr);
    table.appendChild(tfoot);
  }

  container.appendChild(table);
  return container;
}

function renderStateBadge(state) {
  const s = String(state).toLowerCase().trim();
  let cls = '';
  if (['done', 'closed', 'completed', 'completado', 'cerrado', 'terminado'].includes(s))
    cls = 'state-done';
  else if (['active', 'in progress', 'en curso', 'en progreso', 'activo'].includes(s))
    cls = 'state-active';
  else if (['new', 'to do', 'nuevo', 'a hacer', 'pendiente'].includes(s))
    cls = 'state-new';
  else if (['resolved', 'resuelto'].includes(s))
    cls = 'state-resolved';
  else
    cls = '';

  if (!state) return '';
  return cls
    ? `<span class="state-badge ${cls}">${escHtml(state)}</span>`
    : escHtml(state);
}

// ─────────────────────────────────────────
// GLOBAL STATS UPDATE
// ─────────────────────────────────────────
function updateGlobalStats() {
  const files = AppState.files;
  let totalTasks     = 0;
  let totalCompleted = 0;
  let totalEstimate  = 0;
  let totalRemaining = 0;

  files.forEach(f => {
    totalTasks     += f.summary.totalRows;
    totalCompleted += f.summary.completedWork;
    totalEstimate  += f.summary.originalEstimate;
    totalRemaining += f.summary.remainingWork;
  });

  const progress = totalEstimate > 0
    ? Math.min(100, Math.round((totalCompleted / totalEstimate) * 100))
    : 0;

  DOM.statFiles.textContent     = files.length;
  DOM.statTasks.textContent     = totalTasks;
  DOM.statCompleted.textContent = fmtHours(round2(totalCompleted));
  DOM.statEstimate.textContent  = fmtHours(round2(totalEstimate));
  DOM.statRemaining.textContent = fmtHours(round2(totalRemaining));
  DOM.statProgress.textContent  = progress + '%';
}

// ─────────────────────────────────────────
// FILE HISTORY IN SIDEBAR
// ─────────────────────────────────────────
function updateFileHistory() {
  DOM.fileHistory.innerHTML = '';

  if (!AppState.files.length) {
    DOM.fileHistory.innerHTML = `
      <div class="empty-history">
        <i class="fas fa-inbox"></i>
        <p>No hay archivos</p>
      </div>`;
    return;
  }

  AppState.files.forEach(f => {
    const item = document.createElement('div');
    item.className = 'history-item';
    const isExcel = f.ext === 'xlsx' || f.ext === 'xls';
    item.innerHTML = `
      <i class="${isExcel ? 'fas fa-file-excel' : 'fas fa-file-csv'}"></i>
      <span class="history-name" title="${escHtml(f.name)}">${escHtml(f.name)}</span>
      <span class="history-hours">${fmtHours(f.summary.completedWork)}</span>
    `;
    item.addEventListener('click', () => {
      const card = document.querySelector(`[data-file-id="${f.id}"]`);
      if (card) {
        card.scrollIntoView({ behavior: 'smooth', block: 'start' });
        card.style.boxShadow = '0 0 0 2px var(--accent-blue)';
        setTimeout(() => { card.style.boxShadow = ''; }, 2000);
      }
      if (window.innerWidth <= 768) closeSidebar();
    });
    DOM.fileHistory.appendChild(item);
  });
}

// ─────────────────────────────────────────
// SHOW / HIDE SECTIONS
// ─────────────────────────────────────────
function showResultsAndStats() {
  DOM.summarySection.style.display = '';
  DOM.resultsSection.style.display = '';
}

function updateResultsEmptyState() {
  if (!DOM.resultsEmptyState) return;
  DOM.resultsEmptyState.style.display = AppState.files.length ? 'none' : '';
}

function clearAll() {
  if (!AppState.files.length) return;
  AppState.files = [];
  DOM.fileResultsContainer.innerHTML = '';
  updateGlobalStats();
  updateFileHistory();
  updateResultsEmptyState();
  showToast('Todos los archivos han sido eliminados.', 'info');
}

// ─────────────────────────────────────────
// TOAST NOTIFICATIONS
// ─────────────────────────────────────────
function showToast(message, type = 'info') {
  const ICONS = { success: 'fas fa-check-circle', error: 'fas fa-times-circle', info: 'fas fa-info-circle' };
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<i class="${ICONS[type] || ICONS.info}"></i><span>${escHtml(message)}</span>`;
  DOM.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.transition = 'opacity .3s ease, transform .3s ease';
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(20px)';
    setTimeout(() => toast.remove(), 320);
  }, 4000);
}

// ─────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
