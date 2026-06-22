/* ═══════════════════════════════════════════════════════════════
  AZURE TASK SUITE — app.js
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

// Azure DevOps Configuration
const AzureConfig = {
  orgUrl: null,
  project: null,
  isConnected: false
};

// Azure runtime state (in-memory + sessionStorage)
const AzureState = {
  rows: []  // raw rows from last successful fetch
};
const AZURE_ROWS_SESSION_KEY = 'azure_rows_cache';

const ABOUT_PHOTO_STORAGE_KEY = 'about_profile_photo_dataurl';
const ABOUT_PHOTO_LOCK_KEY = 'about_profile_photo_locked';
const ABOUT_PHOTO_CHANGE_PASSWORD = '580622';
const SHARED_PROFILE_PHOTO_CANDIDATES = [
  'https://cristiandevbonelo.github.io/porfoliocristian/assets/risa.png',
  'assets/profile-photo.svg',
  'assets/profile-photo.jpg',
  'assets/profile-photo.jpeg',
  'assets/profile-photo.png',
  'assets/profile-photo.webp'
];

const AZURE_DEFAULT_MAX_ITEMS = 5000;

function getAzureBackendCandidates() {
  const candidates = ['api/azure.php', './api/azure.php', '/api/azure.php'];
  const normalized = [];

  for (const candidate of candidates) {
    try {
      normalized.push(new URL(candidate, window.location.href).toString());
    } catch (_) {
      // Ignore malformed URL candidates and continue with valid ones.
    }
  }

  return [...new Set(normalized)];
}

const GITHUB_REPO_OWNER = 'CristianBoneloRios';
const GITHUB_REPO_NAME = 'Calculator-Azure-task';
const GITHUB_REPO_BRANCH = 'main';
const GITHUB_PROFILE_PHOTO_PATH = 'assets/profile-photo.png';

// ─────────────────────────────────────────
// SAFE LOCALSTORAGE HELPERS
// ─────────────────────────────────────────
function safeGetItem(key) {
  try {
    return localStorage.getItem(key);
  } catch (e) {
    console.warn('localStorage unavailable:', e.message);
    return null;
  }
}

function safeSetItem(key, value) {
  try {
    localStorage.setItem(key, value);
    return true;
  } catch (e) {
    console.warn('localStorage unavailable:', e.message);
    return false;
  }
}

// ─────────────────────────────────────────
// AZURE DEVOPS COLUMN MAPPING (EN + ES)
// ─────────────────────────────────────────
const COL_CANDIDATES = {
  id:               ['id', 'work item id', 'id de elemento de trabajo', 'workitemid'],
  title:            ['title', 'título', 'titulo', 'name', 'nombre', 'summary', 'descripción', 'descripcion',
                     'task name', 'task title', 'asunto', 'subject', 'tarea'],
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
  workDate:         ['date', 'fecha', 'work date', 'task date', 'fecha de trabajo',
                     'completed date', 'completed on', 'changed date', 'closed date',
                     'resolved date', 'finish date', 'end date', 'start date',
                     'startdate', 'día', 'dia'],
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
  DOM.diagnosticoBtn       = document.getElementById('diagnosticoBtn');
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

  // Azure DevOps DOM elements
  DOM.connectAzureBtn       = document.getElementById('connectAzureBtn');
  DOM.azureConnectBackdrop  = document.getElementById('azureConnectBackdrop');
  DOM.azureConnectModal     = document.getElementById('azureConnectModal');
  DOM.azureOrgInput         = document.getElementById('azureOrgInput');
  DOM.azureProjectInput     = document.getElementById('azureProjectInput');
  DOM.azurePatInput         = document.getElementById('azurePatInput');
  DOM.azurePatServerStatus  = document.getElementById('azurePatServerStatus');
  DOM.azureConnectError     = document.getElementById('azureConnectError');
  DOM.azureConnectCancelBtn = document.getElementById('azureConnectCancelBtn');
  DOM.azureDeletePatBtn     = document.getElementById('azureDeletePatBtn');
  DOM.azureConnectConfirmBtn= document.getElementById('azureConnectConfirmBtn');
  DOM.azureLoadingBackdrop  = document.getElementById('azureLoadingBackdrop');
  DOM.azureLoadingModal     = document.getElementById('azureLoadingModal');
  DOM.azureLoadingStatus    = document.getElementById('azureLoadingStatus');
  DOM.azureLoadingText      = document.getElementById('azureLoadingText');
  DOM.azureLoadingProgressFill = document.getElementById('azureLoadingProgressFill');
  DOM.azureLoadingPercent   = document.getElementById('azureLoadingPercent');

  // Azure tasks section
  DOM.azureConnectedStrip   = document.getElementById('azureConnectedStrip');
  DOM.azureConnectedProject = document.getElementById('azureConnectedProject');
  DOM.azureShowTasksBtn     = document.getElementById('azureShowTasksBtn');
  DOM.azureShowTasksBtnLabel= document.getElementById('azureShowTasksBtnLabel');
  DOM.azureTasksSection     = document.getElementById('azureTasksSection');
  DOM.azureTasksSummaryText = document.getElementById('azureTasksSummaryText');
  DOM.azureSummaryProgress  = document.getElementById('azureSummaryProgress');
  DOM.azureSummaryProgressFill = document.getElementById('azureSummaryProgressFill');
  DOM.azureSummaryProgressPercent = document.getElementById('azureSummaryProgressPercent');
  DOM.azureReloadBtn        = document.getElementById('azureReloadBtn');
  DOM.azureFilterUser       = document.getElementById('azureFilterUser');
  DOM.azureFilterState      = document.getElementById('azureFilterState');
  DOM.azureFilterType       = document.getElementById('azureFilterType');
  DOM.azureFilterDateFrom   = document.getElementById('azureFilterDateFrom');
  DOM.azureFilterDateTo     = document.getElementById('azureFilterDateTo');
  DOM.azureApplyFilterBtn   = document.getElementById('azureApplyFilterBtn');
  DOM.azureResetFilterBtn   = document.getElementById('azureResetFilterBtn');
  DOM.azureSkeleton         = document.getElementById('azureSkeleton');
  DOM.azureTasksResults     = document.getElementById('azureTasksResults');
  DOM.azureTasksEmpty       = document.getElementById('azureTasksEmpty');
  DOM.azureResultsMeta      = document.getElementById('azureResultsMeta');
  DOM.azureResultsCount     = document.getElementById('azureResultsCount');
  DOM.footerPhotoWrap       = document.getElementById('footerPhotoWrap');
  DOM.footerPhotoPreviewBackdrop = document.getElementById('footerPhotoPreviewBackdrop');
  DOM.footerPhotoPreviewImg = document.getElementById('footerPhotoPreviewImg');
  DOM.workspaceLoginBtn     = document.getElementById('workspaceLoginBtn');
  DOM.workspaceLoginBtnName = document.getElementById('workspaceLoginBtnName');
  DOM.workspaceLoginBtnRole = document.getElementById('workspaceLoginBtnRole');

  const hasLocalPhoto = restoreAboutPhotoFromStorage();
  if (!hasLocalPhoto) {
    loadPublicProfilePhoto();
  }
  restoreAzureConnection();
  initUploadEvents();
  initEvents();
  syncWorkspaceLoginButton();
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
  DOM.diagnosticoBtn.addEventListener('click', openDiagnostico);
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

  // Azure DevOps connection
  DOM.connectAzureBtn.addEventListener('click', openAzureConnectModal);

  // Azure tasks panel
  DOM.azureShowTasksBtn.addEventListener('click', fetchAzureWorkItems);
  DOM.azureReloadBtn.addEventListener('click', fetchAzureWorkItems);
  DOM.azureApplyFilterBtn.addEventListener('click', applyAzureFilters);
  DOM.azureResetFilterBtn.addEventListener('click', resetAzureFilters);
  DOM.azureFilterUser.addEventListener('keydown', e => { if (e.key === 'Enter') applyAzureFilters(); });
  DOM.azureFilterState.addEventListener('change', applyAzureFilters);
  DOM.azureFilterType.addEventListener('change', applyAzureFilters);
  DOM.azureFilterDateFrom.addEventListener('change', applyAzureFilters);
  DOM.azureFilterDateTo.addEventListener('change', applyAzureFilters);

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

  // Footer photo hover preview
  initFooterPhotoHoverPreview();
}

function initUploadEvents() {
  if (DOM.selectFilesBtn && DOM.fileInput) {
    DOM.selectFilesBtn.addEventListener('click', () => DOM.fileInput.click());
  }

  if (DOM.fileInput) {
    DOM.fileInput.addEventListener('change', e => handleFiles(e.target.files));
  }

  if (!DOM.uploadZone) return;

  const preventDefaults = e => {
    e.preventDefault();
    e.stopPropagation();
  };

  ['dragenter', 'dragover'].forEach(eventName => {
    DOM.uploadZone.addEventListener(eventName, e => {
      preventDefaults(e);
      DOM.uploadZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    DOM.uploadZone.addEventListener(eventName, e => {
      preventDefaults(e);
      if (eventName === 'dragleave') {
        if (!DOM.uploadZone.contains(e.relatedTarget)) {
          DOM.uploadZone.classList.remove('drag-over');
        }
        return;
      }

      DOM.uploadZone.classList.remove('drag-over');
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        handleFiles(e.dataTransfer.files);
      }
    });
  });
}

function initFooterPhotoHoverPreview() {
  if (!DOM.footerPhotoWrap || !DOM.footerPhotoPreviewBackdrop || !DOM.footerPhotoPreviewImg) return;
  if (!window.matchMedia || !window.matchMedia('(hover: hover)').matches) return;

  const thumb = DOM.footerPhotoWrap.querySelector('img');
  if (thumb && thumb.src) {
    DOM.footerPhotoPreviewImg.src = thumb.src;
  }

  let hideTimer = null;

  const showPreview = () => {
    if (hideTimer) {
      clearTimeout(hideTimer);
      hideTimer = null;
    }
    DOM.footerPhotoPreviewBackdrop.classList.add('active');
  };

  const hidePreview = () => {
    hideTimer = window.setTimeout(() => {
      DOM.footerPhotoPreviewBackdrop.classList.remove('active');
    }, 40);
  };

  DOM.footerPhotoWrap.addEventListener('mouseenter', showPreview);
  DOM.footerPhotoWrap.addEventListener('mouseleave', hidePreview);
}

async function syncWorkspaceLoginButton() {
  if (!DOM.workspaceLoginBtn || !DOM.workspaceLoginBtnName || !DOM.workspaceLoginBtnRole) return;

  DOM.workspaceLoginBtnName.textContent = 'Iniciar sesión';
  DOM.workspaceLoginBtnName.classList.remove('logged-user');
  DOM.workspaceLoginBtnRole.textContent = 'Acceder al Workspace';

  try {
    const response = await fetch('api/auth.php?action=session', {
      method: 'GET',
      credentials: 'same-origin'
    });

    if (!response.ok) return;

    const data = await response.json().catch(() => null);
    if (!data || !data.authenticated || !data.user) return;

    const fullName = String(data.user.full_name || data.user.email || '').trim();
    if (!fullName) return;

    DOM.workspaceLoginBtnName.innerHTML = `Sesión iniciada: <span class="workspace-user-highlight">${escapeHtml(fullName)}</span>`;
    DOM.workspaceLoginBtnName.classList.add('logged-user');
    DOM.workspaceLoginBtnRole.textContent = 'Entrar al Workspace';
    DOM.workspaceLoginBtn.title = `Sesión iniciada: ${fullName}`;
  } catch (_) {
    // Silent fallback to default state if auth service is unavailable.
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setActiveNavItem(sectionId) {
  document.querySelectorAll('.nav-item[data-section]').forEach(item => {
    item.classList.toggle('active', item.dataset.section === sectionId);
  });
}

function initSectionObserver() {
  const sectionIds = ['upload-section', 'summary-section', 'results-section', 'normalizer-section'];
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

// ─────────────────────────────────────────
// DIAGNOSTICO
// ─────────────────────────────────────────
function openDiagnostico() {
  const baseUrl = window.location.origin;
  const diagUrl = baseUrl + '/api/diagnose.php';
  window.open(diagUrl, '_blank');
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
  safeSetItem(ABOUT_PHOTO_STORAGE_KEY, src);
  safeSetItem(ABOUT_PHOTO_LOCK_KEY, 'true');
}

function restoreAboutPhotoFromStorage() {
  const storedPhoto = safeGetItem(ABOUT_PHOTO_STORAGE_KEY);
  if (!storedPhoto) return false;
  applyAboutPhoto(storedPhoto);

  // Keep compatibility if photo existed before lock flag was created.
  if (!safeGetItem(ABOUT_PHOTO_LOCK_KEY)) {
    safeSetItem(ABOUT_PHOTO_LOCK_KEY, 'true');
  }
  return true;
}

function isAboutPhotoLocked() {
  return safeGetItem(ABOUT_PHOTO_LOCK_KEY) === 'true';
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

async function loadPublicProfilePhoto() {
  try {
    const response = await fetch('api/public_profile.php', {
      method: 'GET',
      credentials: 'same-origin'
    });

    if (!response.ok) {
      loadSharedAboutPhoto();
      return;
    }

    const data = await response.json().catch(() => null);
    const photoUrl = String(data?.profile?.photo_url || '').trim();

    if (photoUrl !== '') {
      applyAboutPhoto(photoUrl);
      return;
    }

    loadSharedAboutPhoto();
  } catch (_) {
    loadSharedAboutPhoto();
  }
}

async function handlePublishPhotoToGithub() {
  const photoDataUrl = safeGetItem(ABOUT_PHOTO_STORAGE_KEY);
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
      let msg = `Error HTTP ${response.status}`;
      if (response.status === 401) {
        msg = 'Token inválido o expirado. Revisa que lo copiaste correctamente.';
      } else if (response.status === 403) {
        msg = 'El token no tiene permisos suficientes. Asegúrate de seleccionar "repo" al crear el token.';
      } else if (response.status === 422) {
        msg = 'Error de validación. Verifica que el repositorio y la rama existen.';
      } else if (err && err.message) {
        msg = err.message;
      }
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
    let msg = `Error HTTP ${response.status}`;
    if (response.status === 401) {
      msg = 'Token inválido o expirado. Revisa que lo copiaste correctamente.';
    } else if (response.status === 403) {
      msg = 'El token no tiene permisos suficientes. Asegúrate de seleccionar "repo" al crear el token.';
    } else if (err && err.message) {
      msg = err.message;
    }
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
      if (token.includes(' ') || token.includes('\n')) {
        DOM.githubPublishError.textContent = '⚠️ El token tiene espacios. Revisa que lo copiaste correctamente.';
        DOM.githubTokenInput.focus();
        return;
      }
      if (!token.startsWith('github_pat_') && !token.startsWith('ghp_')) {
        DOM.githubPublishError.textContent = '⚠️ El token no parece válido. Debe empezar con github_pat_ o ghp_';
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

  const str = String(value)
    .trim()
    .toLowerCase()
    .replace(/,/g, '.')
    .replace(/[—–-]/g, '')
    .replace(/horas?|hrs?/g, 'h')
    .replace(/minutos?|mins?/g, 'm')
    .replace(/\s+/g, ' ')
    .trim();

  if (!str) return 0;

  const hoursMatch = str.match(/(\d+(?:\.\d+)?)\s*h\b/);
  const minsMatch = str.match(/(\d+(?:\.\d+)?)\s*m\b/);

  if (hoursMatch || minsMatch) {
    const h = hoursMatch ? parseFloat(hoursMatch[1]) : 0;
    const m = minsMatch ? parseFloat(minsMatch[1]) : 0;
    const total = h + (m / 60);
    return Number.isNaN(total) ? 0 : Math.max(0, total);
  }

  const n = parseFloat(str);
  return Number.isNaN(n) ? 0 : Math.max(0, n);
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

  // Daily hours distribution panel
  const dailyHoursPanel = buildDailyHoursPanel(rows, headers, colMap);

  // Table & Kanban container
  const table = buildTable(rows, headers, colMap);
  const kanban = buildKanban(rows, headers, colMap);
  
  const viewContainer = document.createElement('div');
  viewContainer.className = 'view-container';
  
  const viewToggle = buildViewToggle();
  table.style.display = 'block';
  kanban.style.display = 'none';
  
  viewContainer.appendChild(viewToggle);
  viewContainer.appendChild(table);
  viewContainer.appendChild(kanban);
  
  // Attach toggle listeners
  viewToggle.querySelectorAll('.view-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const view = btn.dataset.view;
      const isTable = view === 'table';
      
      viewToggle.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      
      table.style.display = isTable ? 'block' : 'none';
      kanban.style.display = isTable ? 'none' : 'block';
    });
  });

  // Warning if no hour cols detected
  const warning = !summary.hasHourCols
    ? `<div class="no-hours-warning"><i class="fas fa-exclamation-triangle"></i>
        No se detectaron columnas de horas estándar de Azure DevOps en este archivo.
        Se muestran todos los datos disponibles.</div>`
    : '';

  bodyDiv.appendChild(hoursBar);
  bodyDiv.appendChild(progressBar);
  bodyDiv.appendChild(dailyHoursPanel);
  if (warning) bodyDiv.insertAdjacentHTML('beforeend', warning);
  bodyDiv.appendChild(viewContainer);

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

function buildDailyHoursPanel(rows, headers, colMap) {
  const panel = document.createElement('div');
  panel.className = 'daily-hours-panel';

  const distribution = calculateDailyHoursDistribution(rows, headers, colMap);

  const header = document.createElement('div');
  header.className = 'daily-hours-header';

  const titleBlock = document.createElement('div');
  titleBlock.className = 'daily-hours-title-block';
  titleBlock.innerHTML = `
    <h4><i class="fas fa-calendar-day"></i> Distribucion de horas por dia</h4>
    <p>${distribution.subtitle}</p>
  `;

  const totalBadge = document.createElement('div');
  totalBadge.className = 'daily-hours-total-badge';
  totalBadge.innerHTML = `<span>Total</span><strong>${fmtHoursDecimal(distribution.totalHours)}</strong>`;
  const totalBadgeValue = totalBadge.querySelector('strong');

  header.appendChild(titleBlock);
  header.appendChild(totalBadge);
  panel.appendChild(header);

  if (!distribution.items.length) {
    const empty = document.createElement('div');
    empty.className = 'daily-hours-empty';
    empty.innerHTML = '<i class="fas fa-circle-info"></i>No se encontraron filas con fecha y horas validas para construir la distribucion diaria.';
    panel.appendChild(empty);
    return panel;
  }

  const firstKey = distribution.items[0].key;
  const lastKey = distribution.items[distribution.items.length - 1].key;

  const toolbar = document.createElement('div');
  toolbar.className = 'daily-hours-toolbar';
  toolbar.innerHTML = `
    <div class="daily-filter-group">
      <div class="daily-quick-ranges" data-role="quick-ranges">
        <button type="button" class="daily-quick-btn" data-period="7d">7 dias</button>
        <button type="button" class="daily-quick-btn" data-period="15d">15 dias</button>
        <button type="button" class="daily-quick-btn" data-period="30d">30 dias</button>
        <button type="button" class="daily-quick-btn" data-period="month">Mes actual</button>
      </div>
      <label>
        Desde
        <input type="date" class="daily-filter-input" data-role="from" min="${firstKey}" max="${lastKey}" value="${firstKey}">
      </label>
      <label>
        Hasta
        <input type="date" class="daily-filter-input" data-role="to" min="${firstKey}" max="${lastKey}" value="${lastKey}">
      </label>
      <button type="button" class="daily-filter-reset" data-role="reset">
        <i class="fas fa-rotate-left"></i> Reset
      </button>
    </div>
    <div class="daily-trend-badge daily-trend-flat" data-role="trend"></div>
  `;

  const sparklineCard = document.createElement('div');
  sparklineCard.className = 'daily-hours-sparkline-card';
  sparklineCard.innerHTML = `
    <div class="daily-sparkline-head">
      <span><i class="fas fa-wave-square"></i> Tendencia semanal</span>
      <small data-role="spark-meta">0 semanas</small>
    </div>
    <div class="daily-sparkline-canvas" data-role="spark-canvas"></div>
  `;

  const list = document.createElement('div');
  list.className = 'daily-hours-list';

  const countFooter = document.createElement('div');
  countFooter.className = 'daily-hours-count';
  countFooter.textContent = 'Dias registrados: 0';

  const fromInput = toolbar.querySelector('[data-role="from"]');
  const toInput = toolbar.querySelector('[data-role="to"]');
  const resetBtn = toolbar.querySelector('[data-role="reset"]');
  const quickButtons = [...toolbar.querySelectorAll('.daily-quick-btn')];
  const trendBadge = toolbar.querySelector('[data-role="trend"]');
  const sparkMeta = sparklineCard.querySelector('[data-role="spark-meta"]');
  const sparkCanvas = sparklineCard.querySelector('[data-role="spark-canvas"]');

  const setQuickActive = period => {
    quickButtons.forEach(btn => {
      btn.classList.toggle('active', !!period && btn.dataset.period === period);
    });
  };

  const renderRange = () => {
    let startKey = fromInput.value || firstKey;
    let endKey = toInput.value || lastKey;

    if (startKey > endKey) {
      const tmp = startKey;
      startKey = endKey;
      endKey = tmp;
      fromInput.value = startKey;
      toInput.value = endKey;
    }

    const filteredItems = filterDistributionByRange(distribution.items, startKey, endKey);
    const total = round2(filteredItems.reduce((acc, item) => acc + item.hours, 0));
    totalBadgeValue.textContent = fmtHoursDecimal(total);

    renderTrendBadge(trendBadge, filteredItems);
    renderWeeklySparkline(sparkCanvas, sparkMeta, filteredItems);
    renderDailyList(list, filteredItems, { rows, headers, colMap });
    countFooter.textContent = `Dias registrados: ${filteredItems.length}`;
  };

  fromInput.addEventListener('change', () => {
    setQuickActive('');
    renderRange();
  });
  toInput.addEventListener('change', () => {
    setQuickActive('');
    renderRange();
  });

  quickButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const period = btn.dataset.period;
      const baseEnd = toInput.value || lastKey;

      let startKey = firstKey;
      let endKey = lastKey;

      if (period === '7d') {
        endKey = clampIsoDateKey(baseEnd, firstKey, lastKey);
        startKey = clampIsoDateKey(shiftIsoDate(endKey, -6), firstKey, lastKey);
      } else if (period === '15d') {
        endKey = clampIsoDateKey(baseEnd, firstKey, lastKey);
        startKey = clampIsoDateKey(shiftIsoDate(endKey, -14), firstKey, lastKey);
      } else if (period === '30d') {
        endKey = clampIsoDateKey(baseEnd, firstKey, lastKey);
        startKey = clampIsoDateKey(shiftIsoDate(endKey, -29), firstKey, lastKey);
      } else if (period === 'month') {
        endKey = clampIsoDateKey(baseEnd, firstKey, lastKey);
        const monthBounds = getMonthBoundsFromIso(endKey);
        startKey = clampIsoDateKey(monthBounds.startKey, firstKey, lastKey);
        endKey = clampIsoDateKey(monthBounds.endKey, firstKey, lastKey);
      }

      if (startKey > endKey) {
        startKey = endKey;
      }

      fromInput.value = startKey;
      toInput.value = endKey;
      setQuickActive(period);
      renderRange();
    });
  });

  resetBtn.addEventListener('click', () => {
    fromInput.value = firstKey;
    toInput.value = lastKey;
    setQuickActive('');
    renderRange();
  });

  panel.appendChild(toolbar);
  panel.appendChild(sparklineCard);
  panel.appendChild(list);
  panel.appendChild(countFooter);
  renderRange();
  return panel;
}

function filterDistributionByRange(items, startKey, endKey) {
  return items.filter(item => item.key >= startKey && item.key <= endKey);
}

function renderDailyList(container, items, fileData) {
  container.innerHTML = '';

  if (!items.length) {
    const empty = document.createElement('div');
    empty.className = 'daily-hours-empty';
    empty.innerHTML = '<i class="fas fa-filter-circle-xmark"></i>No hay datos en el rango seleccionado.';
    container.appendChild(empty);
    return;
  }

  const maxHours = Math.max(...items.map(item => item.hours));
  items.forEach(item => {
    const row = document.createElement('div');
    row.className = 'daily-hours-row';
    row.style.cursor = 'pointer';
    row.dataset.dateKey = item.key;
    row.dataset.dateLabel = item.label;
    row.dataset.totalHours = item.hours;

    const pct = maxHours > 0 ? Math.max(6, Math.round((item.hours / maxHours) * 100)) : 0;
    row.innerHTML = `
      <div class="daily-hours-date">${escHtml(item.label)}</div>
      <div class="daily-hours-bar-wrap">
        <div class="daily-hours-bar-track">
          <div class="daily-hours-bar-fill" style="width:${pct}%"></div>
        </div>
      </div>
      <div class="daily-hours-value">${fmtHoursDecimal(item.hours)}</div>
    `;

    row.addEventListener('click', () => {
      showDailyTasksModal(item.key, item.label, item.hours, fileData);
    });

    container.appendChild(row);
  });
}

function renderTrendBadge(container, items) {
  if (!container) return;

  if (items.length < 2) {
    container.className = 'daily-trend-badge daily-trend-flat';
    container.innerHTML = '<i class="fas fa-minus"></i> Sin comparacion suficiente';
    return;
  }

  const prev = items[items.length - 2];
  const last = items[items.length - 1];
  const diff = round2(last.hours - prev.hours);

  if (diff > 0) {
    const pct = prev.hours > 0 ? Math.round((diff / prev.hours) * 100) : null;
    container.className = 'daily-trend-badge daily-trend-up';
    container.innerHTML = `<i class="fas fa-arrow-trend-up"></i> Subio ${fmtHoursDecimal(diff)}${pct !== null ? ` (${pct}%)` : ''}`;
    return;
  }

  if (diff < 0) {
    const absDiff = Math.abs(diff);
    const pct = prev.hours > 0 ? Math.round((absDiff / prev.hours) * 100) : null;
    container.className = 'daily-trend-badge daily-trend-down';
    container.innerHTML = `<i class="fas fa-arrow-trend-down"></i> Bajo ${fmtHoursDecimal(absDiff)}${pct !== null ? ` (${pct}%)` : ''}`;
    return;
  }

  container.className = 'daily-trend-badge daily-trend-flat';
  container.innerHTML = '<i class="fas fa-equals"></i> Sin cambio vs dia anterior';
}

function renderWeeklySparkline(container, metaNode, items) {
  if (!container || !metaNode) return;

  const weekly = aggregateByWeek(items);
  metaNode.textContent = `${weekly.length} semana${weekly.length !== 1 ? 's' : ''}`;

  if (!weekly.length) {
    container.innerHTML = '<div class="daily-sparkline-empty">Sin datos para graficar.</div>';
    return;
  }

  const svgMarkup = buildSparklineSVG(weekly.map(w => w.hours));
  const labelsMarkup = weekly
    .slice(-4)
    .map(w => `<span>${escHtml(w.label)}</span>`)
    .join('');

  container.innerHTML = `
    ${svgMarkup}
    <div class="daily-sparkline-labels">${labelsMarkup}</div>
  `;
}

function aggregateByWeek(items) {
  const byWeek = new Map();

  items.forEach(item => {
    const weekStart = getWeekStartKey(item.key);
    if (!weekStart) return;
    byWeek.set(weekStart, round2((byWeek.get(weekStart) || 0) + item.hours));
  });

  return [...byWeek.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, hours]) => ({
      key,
      hours,
      label: formatWeekLabel(key)
    }));
}

function getWeekStartKey(isoKey) {
  const parts = parseIsoDateParts(isoKey);
  if (!parts) return null;

  const date = new Date(parts.year, parts.month - 1, parts.day);
  const weekday = (date.getDay() + 6) % 7;
  date.setDate(date.getDate() - weekday);

  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function formatWeekLabel(isoKey) {
  const parts = parseIsoDateParts(isoKey);
  if (!parts) return isoKey;
  return `Sem ${String(parts.day).padStart(2, '0')}/${String(parts.month).padStart(2, '0')}`;
}

function buildSparklineSVG(values) {
  const width = 300;
  const height = 76;
  const padX = 8;
  const padY = 10;
  const innerW = width - (padX * 2);
  const innerH = height - (padY * 2);

  const points = values.length > 1 ? values : [values[0], values[0]];
  const minVal = Math.min(...points);
  const maxVal = Math.max(...points);
  const range = Math.max(0.0001, maxVal - minVal);

  const coords = points.map((value, idx) => {
    const x = padX + (idx * innerW) / (points.length - 1);
    const y = padY + innerH - (((value - minVal) / range) * innerH);
    return { x, y };
  });

  const polyline = coords.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
  const areaPath = `M ${coords[0].x.toFixed(1)} ${height - padY} L ${polyline.replace(/,/g, ' ')} L ${coords[coords.length - 1].x.toFixed(1)} ${height - padY} Z`;
  const dots = coords.map(p => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2.8"></circle>`).join('');

  return `
    <svg class="daily-sparkline-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" role="img" aria-label="Tendencia semanal de horas">
      <path class="daily-sparkline-area" d="${areaPath}"></path>
      <polyline class="daily-sparkline-line" points="${polyline}"></polyline>
      <g class="daily-sparkline-dots">${dots}</g>
    </svg>
  `;
}

function parseIsoDateParts(isoKey) {
  const m = String(isoKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return {
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3])
  };
}

function isoKeyFromDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function shiftIsoDate(isoKey, daysDelta) {
  const parts = parseIsoDateParts(isoKey);
  if (!parts) return isoKey;

  const date = new Date(parts.year, parts.month - 1, parts.day);
  date.setDate(date.getDate() + daysDelta);
  return isoKeyFromDate(date);
}

function clampIsoDateKey(key, minKey, maxKey) {
  if (key < minKey) return minKey;
  if (key > maxKey) return maxKey;
  return key;
}

function getMonthBoundsFromIso(isoKey) {
  const parts = parseIsoDateParts(isoKey);
  if (!parts) {
    return { startKey: isoKey, endKey: isoKey };
  }

  const firstDay = new Date(parts.year, parts.month - 1, 1);
  const lastDay = new Date(parts.year, parts.month, 0);

  return {
    startKey: isoKeyFromDate(firstDay),
    endKey: isoKeyFromDate(lastDay)
  };
}

function calculateDailyHoursDistribution(rows, headers, colMap) {
  const hoursCol = pickHoursColumn(colMap);
  const dateCol = pickDateColumn(headers, colMap);

  const subtitles = [];
  if (hoursCol) subtitles.push(`Horas: ${hoursCol}`);
  if (dateCol) subtitles.push(`Fecha: ${dateCol}`);

  if (!hoursCol || !dateCol) {
    return {
      items: [],
      totalHours: 0,
      subtitle: subtitles.length
        ? subtitles.join(' · ')
        : 'No se pudo detectar automaticamente una columna de horas y fecha.'
    };
  }

  const byDate = new Map();

  rows.forEach(row => {
    const hours = parseHours(row[hoursCol]);
    if (hours <= 0) return;

    const parsed = parseDateLike(row[dateCol]);
    if (!parsed) return;

    byDate.set(parsed.key, round2((byDate.get(parsed.key) || 0) + hours));
  });

  const items = [...byDate.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, hours]) => ({
      key,
      label: formatDateLabel(key),
      hours
    }));

  const totalHours = round2(items.reduce((acc, item) => acc + item.hours, 0));

  return {
    items,
    totalHours,
    subtitle: subtitles.join(' · ')
  };
}

function pickHoursColumn(colMap) {
  return colMap.completedWork || colMap.originalEstimate || colMap.remainingWork || null;
}

function pickDateColumn(headers, colMap) {
  if (colMap.workDate) return colMap.workDate;

  const likelyDateCols = headers.filter(h => {
    const normalized = normalizeHeaderName(h);
    return normalized.includes('fecha')
      || normalized.includes('date')
      || normalized.includes('dia')
      || normalized.includes('day');
  });

  return likelyDateCols[0] || null;
}

function normalizeHeaderName(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function parseDateLike(value) {
  if (value === null || value === undefined) return null;

  const raw = String(value).trim();
  if (!raw) return null;

  const excelNumeric = parseFloat(raw);
  if (/^\d+(\.\d+)?$/.test(raw) && !Number.isNaN(excelNumeric) && excelNumeric > 59) {
    const parsedExcelDate = parseExcelSerialDate(excelNumeric);
    if (parsedExcelDate) return parsedExcelDate;
  }

  const dmy = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s+\d{1,2}(?::\d{2})?(?::\d{2})?\s*(?:[ap]\.?m\.?)?)?$/i);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    let year = Number(dmy[3]);
    if (year < 100) year += 2000;
    return buildDateKey(year, month, day);
  }

  const iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    return buildDateKey(year, month, day);
  }

  const normalizedRaw = raw
    .replace(/a\.\s*m\.?/i, 'AM')
    .replace(/p\.\s*m\.?/i, 'PM');

  const jsDate = new Date(normalizedRaw);
  if (!Number.isNaN(jsDate.getTime())) {
    return buildDateKey(jsDate.getFullYear(), jsDate.getMonth() + 1, jsDate.getDate());
  }

  return null;
}

function parseExcelSerialDate(serial) {
  const baseDate = new Date(Date.UTC(1899, 11, 30));
  baseDate.setUTCDate(baseDate.getUTCDate() + Math.floor(serial));
  return buildDateKey(baseDate.getUTCFullYear(), baseDate.getUTCMonth() + 1, baseDate.getUTCDate());
}

function buildDateKey(year, month, day) {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const date = new Date(year, month - 1, day);
  if (
    Number.isNaN(date.getTime())
    || date.getFullYear() !== year
    || date.getMonth() !== month - 1
    || date.getDate() !== day
  ) {
    return null;
  }

  const key = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  return { key };
}

function formatDateLabel(isoKey) {
  const [year, month, day] = isoKey.split('-');
  return `${day}/${month}/${year}`;
}

function fmtHoursDecimal(n) {
  const val = round2(Number(n) || 0);
  if (Number.isInteger(val)) return `${val} h`;
  return `${val.toFixed(2).replace(/\.00$/, '').replace(/0$/, '')} h`;
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
    tr.classList.add('azure-row-clickable');
    tr.addEventListener('click', () => openAzureTaskDetailModal(row));
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

function openAzureTaskDetailModal(row) {
  let backdrop = document.getElementById('azureTaskDetailBackdrop');
  let modal = document.getElementById('azureTaskDetailModal');

  if (!backdrop || !modal) {
    createAzureTaskDetailModal();
    backdrop = document.getElementById('azureTaskDetailBackdrop');
    modal = document.getElementById('azureTaskDetailModal');
  }

  const body = document.getElementById('azureTaskDetailBody');
  const title = document.getElementById('azureTaskDetailTitle');
  const subtitle = document.getElementById('azureTaskDetailSubtitle');
  const azureLinkBtn = document.getElementById('azureTaskOpenInDevOpsBtn');
  const azureCopyLinkBtn = document.getElementById('azureTaskCopyLinkBtn');
  if (!backdrop || !modal || !body || !title || !subtitle || !azureLinkBtn || !azureCopyLinkBtn) return;

  const taskId = String(row['ID'] || row['Id'] || '').trim();
  const taskTitle = String(row['Titulo'] || row['Título'] || row['System.Title'] || 'Detalle de tarea').trim();

  title.textContent = taskId ? `Task ${taskId}` : 'Detalle de tarea';
  subtitle.textContent = taskTitle || 'Sin titulo';

  const orgUrl = String(AzureConfig.orgUrl || '').trim();
  const project = String(AzureConfig.project || '').trim();
  let openUrl = '';
  if (taskId && orgUrl && project) {
    openUrl = `${orgUrl}/${encodeURIComponent(project)}/_workitems/edit/${encodeURIComponent(taskId)}`;
    azureLinkBtn.href = openUrl;
    azureLinkBtn.style.display = '';
    azureLinkBtn.setAttribute('aria-disabled', 'false');

    azureCopyLinkBtn.dataset.url = openUrl;
    azureCopyLinkBtn.style.display = '';
    azureCopyLinkBtn.disabled = false;
    azureCopyLinkBtn.innerHTML = '<i class="fas fa-link"></i> Copiar URL';
  } else {
    azureLinkBtn.removeAttribute('href');
    azureLinkBtn.style.display = 'none';
    azureLinkBtn.setAttribute('aria-disabled', 'true');

    azureCopyLinkBtn.dataset.url = '';
    azureCopyLinkBtn.style.display = 'none';
    azureCopyLinkBtn.disabled = true;
  }

  const entries = Object.entries(row).filter(([key]) => !String(key).startsWith('__'));
  body.innerHTML = entries.map(([key, value]) => {
    const printable = value === null || value === undefined || value === '' ? '—' : String(value);
    return `
      <div class="azure-task-detail-item">
        <span class="azure-task-detail-key">${escHtml(key)}</span>
        <span class="azure-task-detail-value">${escHtml(printable)}</span>
      </div>
    `;
  }).join('');

  backdrop.classList.add('active');
  modal.classList.add('active');
  document.body.style.overflow = 'hidden';
}

function createAzureTaskDetailModal() {
  const backdrop = document.createElement('div');
  backdrop.id = 'azureTaskDetailBackdrop';
  backdrop.className = 'daily-tasks-modal-backdrop';

  const modal = document.createElement('div');
  modal.id = 'azureTaskDetailModal';
  modal.className = 'daily-tasks-modal azure-task-detail-modal';
  modal.innerHTML = `
    <button class="daily-tasks-modal-close" id="azureTaskDetailClose" title="Cerrar">
      <i class="fas fa-times"></i>
    </button>
    <div class="daily-tasks-modal-header">
      <div class="daily-tasks-modal-title" id="azureTaskDetailTitle">Detalle de tarea</div>
      <div class="daily-tasks-modal-badge" style="background: rgba(245,158,11,.2); color: #fcd34d; border-color: rgba(245,158,11,.3);">Azure</div>
    </div>
    <div class="daily-tasks-modal-divider"></div>
    <div class="azure-task-detail-subtitle" id="azureTaskDetailSubtitle"></div>
    <div class="azure-task-detail-actions">
      <a id="azureTaskOpenInDevOpsBtn" class="btn btn-azure btn-sm" target="_blank" rel="noopener noreferrer" style="display:none;">
        <i class="fab fa-microsoft"></i> Abrir en Azure DevOps
      </a>
      <button id="azureTaskCopyLinkBtn" class="btn btn-outline btn-sm" type="button" style="display:none;">
        <i class="fas fa-link"></i> Copiar URL
      </button>
    </div>
    <div class="daily-tasks-modal-content azure-task-detail-body" id="azureTaskDetailBody"></div>
  `;

  document.body.appendChild(backdrop);
  document.body.appendChild(modal);

  const closeModal = () => {
    backdrop.classList.remove('active');
    modal.classList.remove('active');
    document.body.style.overflow = '';
  };

  backdrop.addEventListener('click', closeModal);
  const closeBtn = document.getElementById('azureTaskDetailClose');
  const copyBtn = document.getElementById('azureTaskCopyLinkBtn');
  if (closeBtn) {
    closeBtn.addEventListener('click', closeModal);
  }

  if (copyBtn) {
    copyBtn.addEventListener('click', async () => {
      const url = String(copyBtn.dataset.url || '').trim();
      if (!url) {
        showToast('No hay URL disponible para copiar.', 'error');
        return;
      }

      let copied = false;
      try {
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
          await navigator.clipboard.writeText(url);
          copied = true;
        }
      } catch (_) {
        copied = false;
      }

      if (!copied) {
        const helper = document.createElement('textarea');
        helper.value = url;
        helper.setAttribute('readonly', 'true');
        helper.style.position = 'fixed';
        helper.style.opacity = '0';
        document.body.appendChild(helper);
        helper.select();
        try {
          copied = document.execCommand('copy');
        } catch (_) {
          copied = false;
        }
        document.body.removeChild(helper);
      }

      if (copied) {
        copyBtn.innerHTML = '<i class="fas fa-check"></i> Copiado';
        showToast('URL copiada para compartir en Teams/Correo.', 'success');
        setTimeout(() => {
          copyBtn.innerHTML = '<i class="fas fa-link"></i> Copiar URL';
        }, 1300);
      } else {
        showToast('No se pudo copiar automaticamente. Intenta con Ctrl+C.', 'error');
      }
    });
  }

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && modal.classList.contains('active')) {
      closeModal();
    }
  });
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

// ─────────────────────────────────────────
// AZURE DEVOPS INTEGRATION
// ─────────────────────────────────────────
function openAzureConnectModal() {
  const savedOrg = safeGetItem('azure_org') || '';
  const savedProject = safeGetItem('azure_project') || '';

  DOM.azureOrgInput.value = savedOrg;
  DOM.azureProjectInput.value = savedProject;
  DOM.azurePatInput.value = '';
  DOM.azurePatServerStatus.textContent = 'Estado PAT servidor: verificando...';
  DOM.azureConnectError.textContent = '';

  checkServerPatStatus();

  const onConfirm = () => {
    const org = DOM.azureOrgInput.value.trim();
    const project = normalizeAzureProjectName(DOM.azureProjectInput.value);
    const pat = DOM.azurePatInput.value.trim();

    // Validation
    if (!org || !project) {
      DOM.azureConnectError.textContent = 'Organización y proyecto son requeridos.';
      return;
    }

    // Extract org name from URL if user pasted full URL
    const orgMatch = org.match(/dev\.azure\.com\/([^\/]+)/);
    const cleanOrg = orgMatch ? orgMatch[1] : org;

    if (!cleanOrg.match(/^[a-zA-Z0-9_-]+$/)) {
      DOM.azureConnectError.textContent = '⚠️ Nombre de organización inválido.';
      return;
    }

    if (pat && (pat.includes(' ') || pat.includes('\n'))) {
      DOM.azureConnectError.textContent = '⚠️ El PAT tiene espacios. Revisa que lo copiaste correctamente.';
      return;
    }

    // Save credentials
    AzureConfig.orgUrl = `https://dev.azure.com/${cleanOrg}`;
    AzureConfig.project = project;
    AzureConfig.isConnected = true;

    safeSetItem('azure_org', cleanOrg);
    safeSetItem('azure_project', project);

    cleanup();
    updateAzureConnectedStrip();
    fetchAzureWorkItems();
  };

  const onDeletePat = async () => {
    DOM.azureConnectError.textContent = '';
    const confirmed = window.confirm('¿Seguro que deseas borrar el PAT guardado en el servidor?');
    if (!confirmed) return;

    try {
      DOM.azureDeletePatBtn.disabled = true;

      const response = await callAzureBackend({ action: 'delete_pat' });
      if (!response || response.ok !== true) {
        const msg = response && response.message
          ? response.message
          : 'No se pudo borrar el PAT del servidor.';
        throw new Error(msg);
      }

      DOM.azurePatInput.value = '';
      DOM.azurePatServerStatus.textContent = 'Estado PAT servidor: no configurado.';
      showToast('PAT borrado del servidor correctamente.', 'success');
    } catch (error) {
      DOM.azureConnectError.textContent = error.message;
    } finally {
      DOM.azureDeletePatBtn.disabled = false;
    }
  };

  const onCancel = () => {
    cleanup();
  };

  const onKeyDown = e => {
    if (e.key === 'Enter') onConfirm();
    if (e.key === 'Escape') onCancel();
  };

  const cleanup = () => {
    DOM.azureConnectConfirmBtn.removeEventListener('click', onConfirm);
    DOM.azureDeletePatBtn.removeEventListener('click', onDeletePat);
    DOM.azureConnectCancelBtn.removeEventListener('click', onCancel);
    DOM.azureConnectBackdrop.removeEventListener('click', onCancel);
    document.removeEventListener('keydown', onKeyDown);

    DOM.azureConnectBackdrop.classList.remove('active');
    DOM.azureConnectModal.classList.remove('active');
    document.body.style.overflow = '';
  };

  DOM.azureConnectConfirmBtn.addEventListener('click', onConfirm);
  DOM.azureDeletePatBtn.addEventListener('click', onDeletePat);
  DOM.azureConnectCancelBtn.addEventListener('click', onCancel);
  DOM.azureConnectBackdrop.addEventListener('click', onCancel);
  document.addEventListener('keydown', onKeyDown);

  DOM.azureConnectBackdrop.classList.add('active');
  DOM.azureConnectModal.classList.add('active');
  document.body.style.overflow = 'hidden';
  DOM.azureOrgInput.focus();
}

function showAzureLoadingModal(message = 'Obteniendo tareas...') {
  DOM.azureLoadingStatus.textContent = message;
  DOM.azureLoadingText.textContent = 'Esto puede tomar un momento...';
  DOM.azureLoadingBackdrop.style.display = '';
  DOM.azureLoadingModal.style.display = '';
  document.body.style.overflow = 'hidden';
}

function setAzureLoadingProgress(percent, options = {}) {
  const safePercent = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
  const {
    summaryMessage = '',
    modalStatus = '',
    modalDetail = '',
    active = true
  } = options;

  if (DOM.azureSummaryProgress) {
    DOM.azureSummaryProgress.classList.toggle('active', Boolean(active));
  }

  if (DOM.azureSummaryProgressFill) {
    DOM.azureSummaryProgressFill.style.width = `${safePercent}%`;
  }

  if (DOM.azureSummaryProgressPercent) {
    DOM.azureSummaryProgressPercent.textContent = `${safePercent}%`;
  }

  if (DOM.azureLoadingProgressFill) {
    DOM.azureLoadingProgressFill.style.width = `${safePercent}%`;
  }

  if (DOM.azureLoadingPercent) {
    DOM.azureLoadingPercent.textContent = `${safePercent}%`;
  }

  if (DOM.azureTasksSummaryText && summaryMessage) {
    DOM.azureTasksSummaryText.textContent = summaryMessage;
  }

  if (DOM.azureLoadingStatus && modalStatus) {
    DOM.azureLoadingStatus.textContent = modalStatus;
  }

  if (DOM.azureLoadingText && modalDetail) {
    DOM.azureLoadingText.textContent = modalDetail;
  }
}

function getAzurePayloadStagePercent(rowCount) {
  const safeRows = Math.max(0, Number(rowCount) || 0);

  if (safeRows <= 200) return 78;
  if (safeRows <= 800) return 84;
  if (safeRows <= 2000) return 88;
  if (safeRows <= 5000) return 92;
  return 94;
}

function hideAzureLoadingModal() {
  DOM.azureLoadingBackdrop.style.display = 'none';
  DOM.azureLoadingModal.style.display = 'none';
  document.body.style.overflow = '';
}

function normalizeAzureProjectName(projectName) {
  let value = String(projectName || '').trim();

  // Prevent double-encoding issues like "Olimpia%2520Agil".
  for (let i = 0; i < 2; i++) {
    try {
      const decoded = decodeURIComponent(value);
      if (decoded === value) break;
      value = decoded;
    } catch (_) {
      break;
    }
  }

  return value;
}

async function callAzureBackend(payload, timeoutMs = 65000) {
  const endpoints = getAzureBackendCandidates();
  let lastError = null;

  for (const endpoint of endpoints) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const backendResponse = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        signal: controller.signal
      });

      const data = await backendResponse.json().catch(() => null);
      if (!backendResponse.ok || !data) {
        const msg = data && data.message
          ? data.message
          : `Error HTTP ${backendResponse.status}. El backend no devolvio JSON valido.`;
        throw new Error(`${msg} URL: ${endpoint}`);
      }

      return data;
    } catch (fetchErr) {
      if (fetchErr.name === 'AbortError') {
        lastError = new Error(`Tiempo de espera agotado (${Math.round(timeoutMs / 1000)}s). URL: ${endpoint}`);
      } else {
        lastError = new Error(`No se pudo conectar al backend. URL: ${endpoint}. Detalle: ${fetchErr.message}`);
      }
    } finally {
      clearTimeout(timer);
    }
  }

  throw lastError || new Error('No se pudo contactar api/azure.php en ninguna ruta conocida.');
}

async function checkServerPatStatus() {
  try {
    const response = await callAzureBackend({ action: 'status' }, 15000);
    const suffix = response.backendVersion
      ? ` (v${response.backendVersion})`
      : '';

    if (response.hasPat) {
      DOM.azurePatServerStatus.textContent = `Estado PAT servidor: configurado${suffix}.`;
    } else {
      DOM.azurePatServerStatus.textContent = `Estado PAT servidor: no configurado${suffix}.`;
    }
  } catch (err) {
    DOM.azurePatServerStatus.textContent = `Estado PAT servidor: error — ${err.message}`;
    console.error('[checkServerPatStatus]', err);
  }
}

async function fetchAzureWorkItems() {
  if (!AzureConfig.isConnected) {
    showToast('Azure no está conectado.', 'error');
    return;
  }

  showAzureLoadingModal('Obteniendo tareas de Azure DevOps...');
  setAzureLoadingProgress(20, {
    active: true,
    summaryMessage: 'Conectando con Azure DevOps... (20%)',
    modalStatus: 'Conectando con Azure DevOps...',
    modalDetail: 'Validando acceso y preparando consulta.'
  });
  showAzureSkeleton();

  try {
    const projectName = normalizeAzureProjectName(AzureConfig.project);
    const pat = DOM.azurePatInput ? DOM.azurePatInput.value.trim() : '';

    setAzureLoadingProgress(60, {
      active: true,
      summaryMessage: 'Consultando tareas en Azure... (60%)',
      modalStatus: 'Consultando backend de Azure...',
      modalDetail: 'Extrayendo items del proyecto y consolidando respuesta.'
    });

    const payload = await callAzureBackend({
      org: AzureConfig.orgUrl.replace('https://dev.azure.com/', '').trim(),
      project: projectName,
      pat,
      maxItems: AZURE_DEFAULT_MAX_ITEMS
    });

    if (!payload || payload.ok !== true) {
      const msg = payload && payload.message
        ? payload.message
        : 'No se pudo obtener respuesta valida del backend.';
      throw new Error(msg);
    }

    if (!Array.isArray(payload.rows) || payload.rows.length === 0) {
      setAzureLoadingProgress(100, {
        active: false,
        summaryMessage: `${AzureConfig.project} — sin tareas disponibles`,
        modalStatus: 'Completado',
        modalDetail: 'No se encontraron tareas para el proyecto.'
      });
      hideAzureLoadingModal();
      showToast('No se encontraron tareas en el proyecto.', 'info');
      return;
    }

    const payloadStagePercent = getAzurePayloadStagePercent(payload.rows.length);
    setAzureLoadingProgress(payloadStagePercent, {
      active: true,
      summaryMessage: `Payload recibido: ${payload.rows.length} tareas (${payloadStagePercent}%)`,
      modalStatus: `Payload recibido: ${payload.rows.length} tareas`,
      modalDetail: 'Calculando estructura de datos según el tamaño de la carga.'
    });

    if (payload.limitApplied) {
      showToast(`Se cargaron ${payload.rows.length} tareas recientes (límite aplicado: ${payload.limitApplied}).`, 'info');
    }

    await processAzureData(payload.rows, () => {
      setAzureLoadingProgress(95, {
        active: true,
        summaryMessage: `Renderizando tabla... (95%)`,
        modalStatus: 'Renderizando tabla final...',
        modalDetail: 'Construyendo encabezados y filas para mostrar resultados.'
      });
    });

    setAzureLoadingProgress(100, {
      active: false,
      summaryMessage: `${AzureConfig.project} — ${payload.rows.length} tarea${payload.rows.length !== 1 ? 's' : ''}`,
      modalStatus: 'Completado',
      modalDetail: 'Carga finalizada correctamente.'
    });
    hideAzureLoadingModal();

  } catch (error) {
    setAzureLoadingProgress(0, {
      active: false,
      summaryMessage: 'No se pudieron cargar tareas de Azure.',
      modalStatus: 'Error durante la carga',
      modalDetail: 'Revisa credenciales, conectividad o estado del backend.'
    });
    hideAzureLoadingModal();

    const backendHint = DOM.azurePatServerStatus && DOM.azurePatServerStatus.textContent.includes('(v')
      ? ` Backend: ${DOM.azurePatServerStatus.textContent}`
      : '';

    if (error instanceof TypeError && String(error.message).includes('Failed to fetch')) {
      showToast('No se pudo conectar al backend (api/azure.php). En Hostinger verifica que PHP este activo y el archivo exista en /api.', 'error');
    } else {
      showToast(`Error: ${error.message}.${backendHint}`, 'error');
    }

    console.error('Azure fetch error:', error);
  }
}

function convertAzureWorkItemsToRows(workItems) {
  const rows = [];

  workItems.forEach(wi => {
    const fields = wi.fields || {};
    
    const row = {
      'ID': fields['System.Id'] || '',
      'Título': fields['System.Title'] || '',
      'Tipo': fields['System.WorkItemType'] || '',
      'Estado': fields['System.State'] || '',
      'Asignado a': fields['System.AssignedTo']?.displayName || '',
      'Estimación Original': fields['Microsoft.VSTS.Scheduling.OriginalEstimate'] || 0,
      'Trabajo Completado': fields['Microsoft.VSTS.Scheduling.CompletedWork'] || 0,
      'Trabajo Restante': fields['Microsoft.VSTS.Scheduling.RemainingWork'] || 0,
      'Etiquetas': fields['System.Tags'] || '',
      'Ruta de Área': fields['System.AreaPath'] || '',
      'Iteración': fields['System.IterationPath'] || ''
    };

    rows.push(row);
  });

  return rows;
}

async function processAzureData(rows, onBeforeRender = null) {
  if (!rows || rows.length === 0) {
    showToast('No hay datos para procesar.', 'info');
    return;
  }

  // Persist to session so the data survives a page refresh
  saveAzureRowsToSession(rows);
  AzureState.rows = rows;

  // Show the dedicated Azure panel
  await renderAzurePanel(rows, onBeforeRender);
  showToast(`✓ ${rows.length} tareas cargadas desde Azure DevOps`, 'success');
}

// Restore Azure connection if it exists
function restoreAzureConnection() {
  const org = safeGetItem('azure_org');
  const project = safeGetItem('azure_project');

  if (org && project) {
    AzureConfig.orgUrl = `https://dev.azure.com/${org}`;
    AzureConfig.project = project;
    AzureConfig.isConnected = true;

    // Show the connected strip
    updateAzureConnectedStrip();

    // Restore cached rows if any
    const cached = loadAzureRowsFromSession();
    if (cached && cached.length > 0) {
      AzureState.rows = cached;
      renderAzurePanel(cached);
    }
  }
}

// ─────────────────────────────────────────
// AZURE SESSION STORAGE HELPERS
// ─────────────────────────────────────────
function saveAzureRowsToSession(rows) {
  try {
    sessionStorage.setItem(AZURE_ROWS_SESSION_KEY, JSON.stringify(rows));
  } catch (e) {
    console.warn('No se pudo guardar en sessionStorage:', e.message);
  }
}

function loadAzureRowsFromSession() {
  try {
    const raw = sessionStorage.getItem(AZURE_ROWS_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : null;
  } catch (e) {
    return null;
  }
}

// ─────────────────────────────────────────
// AZURE CONNECTED STRIP
// ─────────────────────────────────────────
function updateAzureConnectedStrip() {
  if (!DOM.azureConnectedStrip) return;
  if (AzureConfig.isConnected) {
    DOM.azureConnectedStrip.style.display = 'flex';
    if (DOM.azureConnectedProject) {
      DOM.azureConnectedProject.textContent = AzureConfig.project || '';
    }
    // Update label depending on whether tasks are already shown
    const hasData = AzureState.rows.length > 0;
    if (DOM.azureShowTasksBtnLabel) {
      DOM.azureShowTasksBtnLabel.textContent = hasData ? 'Recargar tareas' : 'Mostrar tareas';
    }
  } else {
    DOM.azureConnectedStrip.style.display = 'none';
  }
}

// ─────────────────────────────────────────
// AZURE TASKS PANEL RENDER
// ─────────────────────────────────────────
function showAzureSkeleton() {
  if (!DOM.azureTasksSection) return;
  DOM.azureTasksSection.style.display = '';
  DOM.azureSkeleton.style.display = 'flex';
  DOM.azureTasksResults.innerHTML = '';
  DOM.azureTasksEmpty.style.display = 'none';
  DOM.azureResultsMeta.style.display = 'none';
}

function hideAzureSkeleton() {
  if (DOM.azureSkeleton) DOM.azureSkeleton.style.display = 'none';
}

function renderAzurePanel(rows, onBeforeRender = null) {
  if (!DOM.azureTasksSection) return Promise.resolve();

  return new Promise(resolve => {
    // Show section + skeleton briefly for effect
    showAzureSkeleton();

    // Populate the state dropdown with unique states from data
    populateStateFilter(rows);
    populateTypeFilter(rows);

    // Small delay to show skeleton shimmer before rendering table
    setTimeout(() => {
      if (typeof onBeforeRender === 'function') {
        onBeforeRender();
      }

      hideAzureSkeleton();
      renderAzureTable(rows);
      DOM.azureTasksSection.style.display = '';
      updateAzureConnectedStrip();
      resolve();
    }, 600);
  });
}

function populateStateFilter(rows) {
  if (!DOM.azureFilterState) return;
  const currentVal = DOM.azureFilterState.value;
  const states = [...new Set(rows.map(r => String(r['Estado'] || r['State'] || '').trim()).filter(Boolean))].sort();

  DOM.azureFilterState.innerHTML = [
    '<option value="">Todos los estados</option>',
    '<option value="bucket:new">New</option>',
    '<option value="bucket:developing">Developing</option>',
    '<option value="bucket:done">Done</option>',
    '<option value="bucket:resolved">Resolved</option>'
  ].join('');

  states.forEach(s => {
    const opt = document.createElement('option');
    opt.value = `exact:${s.toLowerCase()}`;
    opt.textContent = s;
    DOM.azureFilterState.appendChild(opt);
  });

  if (currentVal && Array.from(DOM.azureFilterState.options).some(o => o.value === currentVal)) {
    DOM.azureFilterState.value = currentVal;
  }
}

function populateTypeFilter(rows) {
  if (!DOM.azureFilterType) return;
  const currentVal = DOM.azureFilterType.value;
  const types = [...new Set(rows.map(r => String(r['Tipo'] || r['Type'] || '').trim()).filter(Boolean))].sort();

  DOM.azureFilterType.innerHTML = '<option value="">Todos los tipos</option>';
  types.forEach(type => {
    const opt = document.createElement('option');
    opt.value = type.toLowerCase();
    opt.textContent = type;
    DOM.azureFilterType.appendChild(opt);
  });

  if (currentVal && Array.from(DOM.azureFilterType.options).some(o => o.value === currentVal)) {
    DOM.azureFilterType.value = currentVal;
  }
}

function renderAzureTable(rows) {
  DOM.azureTasksResults.innerHTML = '';

  if (!rows || rows.length === 0) {
    DOM.azureTasksEmpty.style.display = '';
    DOM.azureResultsMeta.style.display = 'none';
    return;
  }

  DOM.azureTasksEmpty.style.display = 'none';

  // Results count badge
  DOM.azureResultsMeta.style.display = 'flex';
  DOM.azureResultsCount.textContent = `${rows.length} tarea${rows.length !== 1 ? 's' : ''}`;

  // Summary text
  if (DOM.azureTasksSummaryText) {
    DOM.azureTasksSummaryText.textContent = `${AzureConfig.project} — ${rows.length} tarea${rows.length !== 1 ? 's' : ''}`;
  }

  const headers = Object.keys(rows[0]);
  const colMap  = identifyColumns(headers);
  const table   = buildTable(rows, headers, colMap);
  DOM.azureTasksResults.appendChild(table);
}

// ─────────────────────────────────────────
// AZURE FILTERS
// ─────────────────────────────────────────
function applyAzureFilters() {
  const userQuery = (DOM.azureFilterUser.value || '').trim().toLowerCase();
  const stateQuery = (DOM.azureFilterState.value || '').trim().toLowerCase();
  const typeQuery = (DOM.azureFilterType && DOM.azureFilterType.value ? DOM.azureFilterType.value : '').trim().toLowerCase();
  const fromDate = (DOM.azureFilterDateFrom && DOM.azureFilterDateFrom.value ? DOM.azureFilterDateFrom.value : '').trim();
  const toDate = (DOM.azureFilterDateTo && DOM.azureFilterDateTo.value ? DOM.azureFilterDateTo.value : '').trim();

  const filtered = AzureState.rows.filter(row => {
    const assignedTo = String(row['Asignado a'] || row['Assigned To'] || row['assignedTo'] || '').toLowerCase();
    const assignedEmail = String(row['Asignado correo'] || row['Assigned To Email'] || row['assignedEmail'] || '').toLowerCase();
    const state      = String(row['Estado']    || row['State']       || row['state']      || '').toLowerCase();
    const type       = String(row['Tipo'] || row['Type'] || '').toLowerCase();
    const rowDate = getRowDateForFiltering(row);

    const userMatch  = !userQuery || assignedTo.includes(userQuery) || assignedEmail.includes(userQuery);
    const stateMatch = matchesStateFilter(stateQuery, state);
    const typeMatch = !typeQuery || type === typeQuery;
    const fromMatch = !fromDate || (rowDate !== null && rowDate >= fromDate);
    const toMatch = !toDate || (rowDate !== null && rowDate <= toDate);

    return userMatch && stateMatch && typeMatch && fromMatch && toMatch;
  });

  renderAzureTable(filtered);
}

function getRowDateForFiltering(row) {
  const candidates = [
    row['Actualizado'],
    row['Changed Date'],
    row['ChangedDate'],
    row['Creado'],
    row['Created Date'],
    row['CreatedDate'],
    row['Fecha'],
    row['Date']
  ];

  for (const value of candidates) {
    const parsed = parseDateLike(value);
    if (parsed && parsed.key) {
      return parsed.key;
    }
  }

  return null;
}

function matchesStateFilter(filterValue, stateValue) {
  if (!filterValue) return true;

  const normalizedState = String(stateValue || '').trim().toLowerCase();
  if (!normalizedState) return false;

  if (filterValue.startsWith('exact:')) {
    return normalizedState === filterValue.replace('exact:', '');
  }

  if (filterValue === 'bucket:done') {
    return ['done', 'closed', 'completed', 'completado', 'cerrado', 'terminado'].includes(normalizedState);
  }

  if (filterValue === 'bucket:developing') {
    return ['active', 'in progress', 'en curso', 'en progreso', 'activo', 'developing', 'development'].includes(normalizedState);
  }

  if (filterValue === 'bucket:new') {
    return ['new', 'to do', 'nuevo', 'a hacer', 'pendiente'].includes(normalizedState);
  }

  if (filterValue === 'bucket:resolved') {
    return ['resolved', 'resuelto'].includes(normalizedState);
  }

  return normalizedState === filterValue;
}

function resetAzureFilters() {
  if (DOM.azureFilterUser)  DOM.azureFilterUser.value  = '';
  if (DOM.azureFilterState) DOM.azureFilterState.value = '';
  if (DOM.azureFilterType) DOM.azureFilterType.value = '';
  if (DOM.azureFilterDateFrom) DOM.azureFilterDateFrom.value = '';
  if (DOM.azureFilterDateTo) DOM.azureFilterDateTo.value = '';
  renderAzureTable(AzureState.rows);
  // Update count badge
  if (DOM.azureResultsCount) {
    DOM.azureResultsCount.textContent = `${AzureState.rows.length} tarea${AzureState.rows.length !== 1 ? 's' : ''}`;
  }
}

// ─────────────────────────────────────────
// DAILY TASKS MODAL
// ─────────────────────────────────────────
function showDailyTasksModal(isoKey, dateLabel, totalHours, fileData) {
  console.log('showDailyTasksModal called with:', { isoKey, dateLabel, totalHours, hasFileData: !!fileData });
  
  try {
    createDailyTasksModalElements();
    renderDailyTasksContent(isoKey, dateLabel, totalHours, fileData);
    
    const backdrop = document.getElementById('daily-tasks-modal-backdrop');
    const modal = document.getElementById('daily-tasks-modal');
    
    console.log('Modal elements found:', { backdrop: !!backdrop, modal: !!modal });
    
    if (backdrop && modal) {
      backdrop.classList.add('active');
      modal.classList.add('active');
      document.body.style.overflow = 'hidden';
      console.log('Modal shown successfully');
    } else {
      console.error('Modal elements not found');
    }
  } catch (error) {
    console.error('Error in showDailyTasksModal:', error);
  }
}

function closeDailyTasksModal() {
  const backdrop = document.getElementById('daily-tasks-modal-backdrop');
  const modal = document.getElementById('daily-tasks-modal');
  if (backdrop && modal) {
    backdrop.classList.remove('active');
    modal.classList.remove('active');
    document.body.style.overflow = '';
  }
}

function createDailyTasksModalElements() {
  if (document.getElementById('daily-tasks-modal')) return; // Already exists

  // Backdrop
  const backdrop = document.createElement('div');
  backdrop.id = 'daily-tasks-modal-backdrop';
  backdrop.className = 'daily-tasks-modal-backdrop';
  backdrop.addEventListener('click', closeDailyTasksModal);

  // Modal container
  const modal = document.createElement('div');
  modal.id = 'daily-tasks-modal';
  modal.className = 'daily-tasks-modal';

  // Close button
  const closeBtn = document.createElement('button');
  closeBtn.id = 'daily-tasks-modal-close';
  closeBtn.className = 'daily-tasks-modal-close';
  closeBtn.innerHTML = '<i class="fas fa-times"></i>';
  closeBtn.addEventListener('click', closeDailyTasksModal);

  // Header
  const header = document.createElement('div');
  header.className = 'daily-tasks-modal-header';
  header.innerHTML = `
    <div class="daily-tasks-modal-title">Tareas del día</div>
    <div class="daily-tasks-modal-badge">0h</div>
  `;

  // Divider
  const divider = document.createElement('div');
  divider.className = 'daily-tasks-modal-divider';

  // Content area
  const content = document.createElement('div');
  content.id = 'daily-tasks-modal-content';
  content.className = 'daily-tasks-modal-content';

  // Build modal structure
  modal.appendChild(closeBtn);
  modal.appendChild(header);
  modal.appendChild(divider);
  modal.appendChild(content);

  document.body.appendChild(backdrop);
  document.body.appendChild(modal);
}

function renderDailyTasksContent(isoKey, dateLabel, totalHours, fileData) {
  try {
    const contentDiv = document.getElementById('daily-tasks-modal-content');
    const badgeDiv = document.querySelector('.daily-tasks-modal-badge');
    const titleDiv = document.querySelector('.daily-tasks-modal-title');
    
    console.log('renderDailyTasksContent - contentDiv found:', !!contentDiv);
    
    if (!contentDiv) {
      console.error('Content div not found');
      return;
    }
    
    if (badgeDiv) badgeDiv.textContent = fmtHours(totalHours);
    if (titleDiv) titleDiv.textContent = `Tareas del ${escHtml(dateLabel)}`;

    const { rows, colMap } = fileData || {};
    console.log('fileData extracted:', { hasRows: !!rows, hasColMap: !!colMap, hasWorkDate: colMap?.workDate });
    
    if (!rows || !colMap) {
      console.error('Missing rows or colMap');
      contentDiv.innerHTML = '<div class="daily-tasks-empty"><i class="fas fa-inbox"></i>No hay datos disponibles.</div>';
      return;
    }

    const tasksForDay = getDailyTasksForDateDirect(isoKey, rows, colMap);
    console.log('Tasks for day:', tasksForDay.length, 'tasks found');
    
    if (!tasksForDay.length) {
      contentDiv.innerHTML = '<div class="daily-tasks-empty"><i class="fas fa-inbox"></i>Sin tareas en este día.</div>';
      return;
    }

    // Create view toggle
    const viewToggle = document.createElement('div');
    viewToggle.className = 'daily-tasks-view-toggle';
    viewToggle.innerHTML = `
      <button class="daily-view-btn daily-view-mindmap active" data-view="mindmap" title="Vista Mapa Mental">
        <i class="fas fa-sitemap"></i> Mapa Mental
      </button>
      <button class="daily-view-btn daily-view-list" data-view="list" title="Vista Lista">
        <i class="fas fa-list"></i> Lista
      </button>
    `;

    // Create custom mindmap view
    const mindmapView = buildCustomMindmapView(dateLabel, totalHours, tasksForDay);
    
    // Create list container
    const listContainer = document.createElement('div');
    listContainer.className = 'daily-tasks-list';
    listContainer.id = 'daily-list-view';
    renderDailyTasksList(listContainer, tasksForDay);
    
    // Setup toggle
    viewToggle.querySelectorAll('.daily-view-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const view = btn.dataset.view;
        const isMindmap = view === 'mindmap';
        
        viewToggle.querySelectorAll('.daily-view-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        
        mindmapView.style.display = isMindmap ? 'block' : 'none';
        listContainer.style.display = isMindmap ? 'none' : 'flex';
      });
    });

    contentDiv.innerHTML = '';
    contentDiv.appendChild(viewToggle);
    contentDiv.appendChild(mindmapView);
    contentDiv.appendChild(listContainer);
    listContainer.style.display = 'none';
    
    console.log('renderDailyTasksContent completed successfully');
  } catch (error) {
    console.error('Error in renderDailyTasksContent:', error);
  }
}

function buildCustomMindmapView(dateLabel, totalHours, tasks) {
  const container = document.createElement('div');
  container.className = 'custom-mindmap-view';

  // Group tasks by state
  const grouped = {};
  tasks.forEach(task => {
    const state = normalizeState(task.state);
    if (!grouped[state]) {
      grouped[state] = [];
    }
    grouped[state].push(task);
  });
  
  // State configuration
  const stateConfig = {
    'Por Hacer': { branchClass: 'branch-todo' },
    'En Progreso': { branchClass: 'branch-progress' },
    'Resuelto': { branchClass: 'branch-resolved' },
    'Completado': { branchClass: 'branch-done' }
  };

  const totalTasks = tasks.length;

  const tree = document.createElement('div');
  tree.className = 'mindmap-tree';
  container.appendChild(tree);

  const leftSide = document.createElement('div');
  leftSide.className = 'mindmap-side mindmap-side-left';

  const center = document.createElement('div');
  center.className = 'mindmap-center';
  center.innerHTML = `
    <div class="mindmap-root-node">
      <div class="mindmap-root-label">DIA</div>
      <div class="mindmap-root-date">${escHtml(dateLabel)}</div>
      <div class="mindmap-root-meta">
        <span class="mindmap-root-hours">${fmtHours(totalHours)}</span>
        <span class="mindmap-root-tasks">${totalTasks} tareas</span>
      </div>
    </div>
  `;

  const rightSide = document.createElement('div');
  rightSide.className = 'mindmap-side mindmap-side-right';

  tree.appendChild(leftSide);
  tree.appendChild(center);
  tree.appendChild(rightSide);

  const stateOrder = ['Por Hacer', 'En Progreso', 'Resuelto', 'Completado'];
  const activeStates = stateOrder.filter(state => grouped[state] && grouped[state].length);
  const leftStates = [];
  const rightStates = [];

  activeStates.forEach((state, index) => {
    if (index % 2 === 0) leftStates.push(state);
    else rightStates.push(state);
  });

  const createBranch = (state, side) => {
    const stateTasks = grouped[state];
    const stateHours = stateTasks.reduce((acc, t) => acc + t.hours, 0);
    const cfg = stateConfig[state];

    const tasksByType = {};
    stateTasks.forEach(task => {
      const typeKey = (task.type || 'Sin tipo').trim() || 'Sin tipo';
      if (!tasksByType[typeKey]) tasksByType[typeKey] = [];
      tasksByType[typeKey].push(task);
    });

    const branch = document.createElement('div');
    branch.className = `mindmap-branch ${cfg.branchClass} side-${side}`;

    const branchNode = document.createElement('div');
    branchNode.className = 'mindmap-branch-node';
    branchNode.innerHTML = `
      <span class="mindmap-branch-name">${escHtml(state)}</span>
      <span class="mindmap-branch-count">${stateTasks.length}</span>
      <span class="mindmap-branch-hours">${fmtHours(stateHours)}</span>
    `;
    branch.appendChild(branchNode);

    const branchChildren = document.createElement('div');
    branchChildren.className = 'mindmap-branch-children';

    Object.keys(tasksByType).sort().forEach(typeName => {
      const typeTasks = tasksByType[typeName];
      const typeHours = typeTasks.reduce((acc, t) => acc + t.hours, 0);

      const typeNode = document.createElement('div');
      typeNode.className = 'mindmap-type-node';
      typeNode.innerHTML = `
        <div class="mindmap-type-header">
          <span class="mindmap-type-name">${escHtml(typeName)}</span>
          <span class="mindmap-type-hours">${fmtHours(typeHours)}</span>
        </div>
      `;

      const leaves = document.createElement('div');
      leaves.className = 'mindmap-leaves';

      typeTasks.forEach(task => {
        const leaf = document.createElement('div');
        leaf.className = 'mindmap-leaf-node';

        const taskTitle = escHtml(task.title).substring(0, 90);
        const taskHours = task.hours > 0 ? fmtHours(task.hours) : '0h';

        leaf.innerHTML = `
          <div class="mindmap-leaf-main">
            <span class="mindmap-leaf-title">${taskTitle}</span>
          </div>
          <div class="mindmap-leaf-meta">
            ${task.assignedTo ? `<span class="mindmap-leaf-assignee">${escHtml(task.assignedTo)}</span>` : ''}
            <span class="mindmap-leaf-hours">${taskHours}</span>
          </div>
        `;

        leaves.appendChild(leaf);
      });

      typeNode.appendChild(leaves);
      branchChildren.appendChild(typeNode);
    });

    branch.appendChild(branchChildren);
    return branch;
  };

  leftStates.forEach(state => leftSide.appendChild(createBranch(state, 'left')));
  rightStates.forEach(state => rightSide.appendChild(createBranch(state, 'right')));
  
  return container;
}

function renderDailyTasksList(container, tasks) {
  container.innerHTML = '';
  tasks.forEach(task => {
    const item = document.createElement('div');
    item.className = 'daily-task-item';

    const icon = getTypeIcon(task.type);
    const hours = fmtHours(task.hours);
    const stateBadge = task.state ? `<span class="state-badge state-${(task.state || '').toLowerCase().replace(/\s+/g, '-')}">${escHtml(task.state)}</span>` : '';

    item.innerHTML = `
      <div class="daily-task-icon">${icon}</div>
      <div class="daily-task-main">
        <div class="daily-task-title">${escHtml(task.title)}</div>
        <div class="daily-task-meta">
          ${task.assignedTo ? `<span class="daily-task-meta-item"><i class="fas fa-user"></i> ${escHtml(task.assignedTo)}</span>` : ''}
          ${task.type ? `<span class="daily-task-meta-item"><i class="fas fa-tag"></i> ${escHtml(task.type)}</span>` : ''}
        </div>
      </div>
      <div class="daily-task-state">${stateBadge}</div>
      <div class="daily-task-hours">${hours}</div>
    `;

    container.appendChild(item);
  });
}

function getDailyTasksForDateDirect(isoKey, rows, colMap) {
  if (!rows || !colMap) return [];

  const tasks = [];
  
  if (!colMap.workDate) return tasks;

  rows.forEach(row => {
    const dateStr = row[colMap.workDate] || '';
    const parsedDate = parseDateLike(dateStr);
    
    if (parsedDate && parsedDate.key === isoKey) {
      const title = row[colMap.title] || '(Sin título)';
      const type = row[colMap.type] || '';
      const assignedTo = row[colMap.assignedTo] || '';
      const state = row[colMap.state] || '';
      const hoursStr = row[colMap.completedWork] || '0';
      const hours = parseHours(hoursStr);

      tasks.push({ title, type, assignedTo, state, hours });
    }
  });

  return tasks;
}

function buildMermaidMindmap(dateLabel, totalHours, tasks) {
  // Group tasks by state, then by type
  const grouped = {};
  
  tasks.forEach(task => {
    const state = normalizeState(task.state);
    const type = task.type || 'Sin tipo';
    
    if (!grouped[state]) {
      grouped[state] = {};
    }
    if (!grouped[state][type]) {
      grouped[state][type] = [];
    }
    grouped[state][type].push(task);
  });
  
  let mermaidCode = `mindmap
  root((📅 ${escapeMarkdown(dateLabel)}))
    💼 Total: ${fmtHours(totalHours)}
`;

  // State order
  const stateOrder = ['Por Hacer', 'En Progreso', 'Resuelto', 'Completado'];
  
  stateOrder.forEach(state => {
    if (grouped[state]) {
      mermaidCode += `    ${getStateIcon(state)} ${state}\n`;
      
      Object.keys(grouped[state]).forEach(type => {
        const typeTasks = grouped[state][type];
        const typeHours = typeTasks.reduce((acc, t) => acc + t.hours, 0);
        mermaidCode += `      ${getTypeIcon(type)} ${escapeMarkdown(type)} (${fmtHours(typeHours)})\n`;
        
        typeTasks.forEach(task => {
          const safeTitle = escapeMarkdown(task.title).substring(0, 40);
          mermaidCode += `        📌 ${safeTitle}${task.hours > 0 ? ` [${fmtHours(task.hours)}]` : ''}\n`;
        });
      });
    }
  });

  return mermaidCode;
}

function normalizeState(state) {
  const s = String(state || '').toLowerCase().trim();
  if (['done', 'closed', 'completed', 'completado', 'cerrado', 'terminado'].includes(s)) return 'Completado';
  if (['active', 'in progress', 'en curso', 'en progreso', 'activo'].includes(s)) return 'En Progreso';
  if (['resolved', 'resuelto'].includes(s)) return 'Resuelto';
  return 'Por Hacer';
}

function getStateIcon(state) {
  switch(state) {
    case 'Completado': return '✅';
    case 'En Progreso': return '⚡';
    case 'Resuelto': return '✔️';
    default: return '📋';
  }
}

function escapeMarkdown(text) {
  return String(text || '')
    .replace(/[*_`#\-\[\]()]/g, ' ')
    .trim();
}

function getTypeIcon(typeStr) {
  if (!typeStr) return '<i class="fas fa-circle-dot"></i>';
  
  const type = String(typeStr).toLowerCase();
  if (type.includes('bug')) return '<i class="fas fa-bug"></i>';
  if (type.includes('feature')) return '<i class="fas fa-star"></i>';
  if (type.includes('task')) return '<i class="fas fa-check-circle"></i>';
  if (type.includes('epic')) return '<i class="fas fa-flag"></i>';
  if (type.includes('story')) return '<i class="fas fa-book"></i>';
  return '<i class="fas fa-circle-dot"></i>';
}

// ─────────────────────────────────────────
// KANBAN VIEW
// ─────────────────────────────────────────
function groupByState(rows, colMap) {
  const groups = {
    pending: [],    // To Do, New, Pendiente
    active: [],     // In Progress, En curso, Activo
    resolved: [],   // Resolved, Resuelto
    done: []        // Done, Completed, Terminado
  };

  rows.forEach(row => {
    const state = String(row[colMap.state] || '').toLowerCase().trim();
    
    if (['done', 'closed', 'completed', 'completado', 'cerrado', 'terminado'].includes(state)) {
      groups.done.push(row);
    } else if (['active', 'in progress', 'en curso', 'en progreso', 'activo'].includes(state)) {
      groups.active.push(row);
    } else if (['resolved', 'resuelto'].includes(state)) {
      groups.resolved.push(row);
    } else {
      groups.pending.push(row);
    }
  });

  return groups;
}

function buildKanban(rows, headers, colMap) {
  const groups = groupByState(rows, colMap);
  const hourCols = new Set([colMap.completedWork, colMap.originalEstimate, colMap.remainingWork].filter(Boolean));

  const container = document.createElement('div');
  container.className = 'kanban-container';

  const columns = [
    { key: 'pending', title: '📋 Por Hacer', icon: 'fa-square-check', color: '#6b7280' },
    { key: 'active', title: '⚡ En Progreso', icon: 'fa-spinner', color: '#3b82f6' },
    { key: 'resolved', title: '✔️ Resuelto', icon: 'fa-check', color: '#06b6d4' },
    { key: 'done', title: '✅ Completado', icon: 'fa-circle-check', color: '#22c55e' }
  ];

  columns.forEach(col => {
    const column = document.createElement('div');
    column.className = 'kanban-column';
    column.style.borderTopColor = col.color;

    const colHeader = document.createElement('div');
    colHeader.className = 'kanban-column-header';
    colHeader.innerHTML = `
      <div class="kanban-col-title">
        <span>${col.title}</span>
        <span class="kanban-col-count">${groups[col.key].length}</span>
      </div>
      <div class="kanban-col-meta">${fmtHours(groups[col.key].reduce((acc, r) => acc + parseHours(r[colMap.completedWork] || '0'), 0))} horas</div>
    `;

    const colBody = document.createElement('div');
    colBody.className = 'kanban-column-body';

    if (groups[col.key].length === 0) {
      const empty = document.createElement('div');
      empty.className = 'kanban-empty';
      empty.innerHTML = '<i class="fas fa-inbox"></i><p>Vacío</p>';
      colBody.appendChild(empty);
    } else {
      groups[col.key].forEach(row => {
        const card = buildKanbanCard(row, headers, colMap, hourCols);
        colBody.appendChild(card);
      });
    }

    column.appendChild(colHeader);
    column.appendChild(colBody);
    container.appendChild(column);
  });

  return container;
}

function buildKanbanCard(row, headers, colMap, hourCols) {
  const card = document.createElement('div');
  card.className = 'kanban-card';

  const title = row[colMap.title] || '(Sin título)';
  const state = row[colMap.state] || '';
  const assignedTo = row[colMap.assignedTo] || '';
  const type = row[colMap.type] || '';
  const hours = parseHours(row[colMap.completedWork] || '0');
  const estimate = parseHours(row[colMap.originalEstimate] || '0');

  const stateBadgeClass = (() => {
    const s = String(state).toLowerCase();
    if (['done', 'closed', 'completed', 'completado', 'cerrado', 'terminado'].includes(s)) return 'state-done';
    if (['active', 'in progress', 'en curso', 'en progreso', 'activo'].includes(s)) return 'state-active';
    if (['resolved', 'resuelto'].includes(s)) return 'state-resolved';
    return 'state-new';
  })();

  const typeIcon = getTypeIcon(type);

  card.innerHTML = `
    <div class="kanban-card-head">
      <div class="kanban-card-type">${typeIcon}</div>
      <div class="kanban-card-state"><span class="state-badge ${stateBadgeClass}">${escHtml(state || 'N/A')}</span></div>
    </div>
    <div class="kanban-card-title">${escHtml(title)}</div>
    ${type ? `<div class="kanban-card-type-label">${escHtml(type)}</div>` : ''}
    ${assignedTo ? `<div class="kanban-card-assigned"><i class="fas fa-user-circle"></i> ${escHtml(assignedTo)}</div>` : ''}
    <div class="kanban-card-footer">
      ${hours > 0 ? `<span class="kanban-card-hours"><i class="fas fa-hourglass-end"></i> ${fmtHours(hours)}</span>` : ''}
      ${estimate > 0 ? `<span class="kanban-card-estimate"><i class="fas fa-clock"></i> ${fmtHours(estimate)}</span>` : ''}
    </div>
  `;

  return card;
}

function buildViewToggle() {
  const container = document.createElement('div');
  container.className = 'view-toggle-container';
  container.innerHTML = `
    <div class="view-toggle">
      <button class="view-btn view-btn-table active" data-view="table" title="Vista tabla">
        <i class="fas fa-table"></i> Tabla
      </button>
      <button class="view-btn view-btn-kanban" data-view="kanban" title="Vista Kanban">
        <i class="fas fa-columns"></i> Kanban
      </button>
    </div>
  `;
  return container;
}
