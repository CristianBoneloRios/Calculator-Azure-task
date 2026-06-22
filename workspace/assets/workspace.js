'use strict';

const workspaceToastContainer = document.getElementById('workspaceToastContainer');
const workspacePage = document.body.dataset.page;

document.addEventListener('DOMContentLoaded', () => {
  if (workspacePage === 'index') {
    initSummaryModuleModal();
    loadSummary();
  }
  if (workspacePage === 'profile') {
    initProfilePage();
  }
  if (workspacePage === 'notes') {
    initNotesPage();
  }
  if (workspacePage === 'tasks') {
    initTasksPage();
  }
  if (workspacePage === 'goals') {
    initGoalsPage();
  }
  if (workspacePage === 'calendar') {
    initCalendarPage();
  }
  if (workspacePage === 'generacion_documentos') {
    initDocumentGenerationPage();
  }
});

const summaryModuleMeta = {
  tasks: {
    kicker: 'Tareas',
    title: 'Tareas priorizadas',
    description: 'Visualiza tareas pendientes y ejecutadas para enfocar tu dia con claridad.',
    href: 'tasks.php',
    points: ['Organiza prioridades del dia', 'Controla estado y vencimientos', 'Actualiza avance en segundos']
  },
  goals: {
    kicker: 'Metas',
    title: 'Metas activas',
    description: 'Monitorea objetivos en curso y detecta rapidamente donde concentrar esfuerzo.',
    href: 'goals.php',
    points: ['Seguimiento porcentual por objetivo', 'Estado activo o completado', 'Alineacion con tareas clave']
  },
  calendar: {
    kicker: 'Agenda',
    title: 'Calendario operativo',
    description: 'Consulta compromisos y sesiones para mantener sincronizacion con el equipo.',
    href: 'calendar.php',
    points: ['Eventos del dia en un vistazo', 'Sincronizacion con Teams/Outlook', 'Bloques de tiempo por prioridad']
  },
  notes: {
    kicker: 'Notas',
    title: 'Notas importantes',
    description: 'Centraliza ideas y recordatorios para no perder informacion critica.',
    href: 'notes.php',
    points: ['Contexto rapido de trabajo', 'Notas clave por prioridad', 'Referencia para decisiones diarias']
  }
};

async function apiRequest(action, options = {}) {
  // En desarrollo (localhost), usar mock-dashboard.php
  const apiFile = window.location.hostname === 'localhost' ? 'mock-dashboard.php' : 'dashboard.php';
  const response = await fetch(`../api/${apiFile}?action=${encodeURIComponent(action)}`, {
    method: options.method || 'GET',
    credentials: 'same-origin',
    headers: options.isFormData ? {} : { 'Content-Type': 'application/json' },
    body: options.body || null,
  });

  let data = null;
  let rawBody = '';

  try {
    data = await response.json();
  } catch (_) {
    try {
      rawBody = await response.text();
    } catch (_) {
      rawBody = '';
    }
  }

  if (!response.ok || !data || data.ok === false) {
    const fallbackMessage = rawBody
      ? `HTTP ${response.status}: ${rawBody.slice(0, 180)}`
      : `HTTP ${response.status}: Error inesperado del servidor.`;
    throw new Error((data && data.message) || fallbackMessage);
  }

  return data;
}

function showWorkspaceToast(message, variant = 'primary') {
  if (!workspaceToastContainer) {
    return;
  }

  const icons = {
    danger:  'fa-circle-xmark',
    success: 'fa-circle-check',
    warning: 'fa-triangle-exclamation',
    primary: 'fa-circle-info',
  };
  const icon = icons[variant] || icons.primary;

  const toast = document.createElement('div');
  toast.className = `ws-toast ${variant}`;
  toast.innerHTML = `<i class="fas ${icon}"></i><span>${message}</span>`;
  workspaceToastContainer.appendChild(toast);

  setTimeout(() => {
    toast.classList.add('hiding');
    setTimeout(() => toast.remove(), 350);
  }, 2600);
}

function renderEmptyState(container, message) {
  container.innerHTML = `<div class="workspace-empty">${message}</div>`;
}

async function loadSummary() {
  const summaryCards = document.getElementById('summaryCards');
  const summaryTasks = document.getElementById('summaryTasks');
  const summaryGoals = document.getElementById('summaryGoals');
  const summaryEvents = document.getElementById('summaryEvents');
  const summaryNotes = document.getElementById('summaryNotes');

  if (summaryCards) {
    summaryCards.innerHTML = `
      <article class="workspace-stat summary-stat-card is-loading"></article>
      <article class="workspace-stat summary-stat-card is-loading"></article>
      <article class="workspace-stat summary-stat-card is-loading"></article>
      <article class="workspace-stat summary-stat-card is-loading"></article>
      <article class="workspace-stat summary-stat-card is-loading"></article>`;
  }

  [summaryTasks, summaryGoals, summaryEvents, summaryNotes].forEach(list => {
    if (!list) return;
    list.innerHTML = `
      <div class="workspace-list-item skeleton-item"></div>
      <div class="workspace-list-item skeleton-item"></div>
      <div class="workspace-list-item skeleton-item"></div>`;
  });

  try {
    const data = await apiRequest('summary');
    summaryCards.innerHTML = `
      <article class="workspace-stat summary-stat-card summary-stat-blue" data-module="notes" style="--ws-delay:0ms"><span class="workspace-tag">Notas</span><strong>${data.summary.notes}</strong><p class="workspace-muted">Ideas activas y referencias</p></article>
      <article class="workspace-stat summary-stat-card summary-stat-green" data-module="tasks" style="--ws-delay:40ms"><span class="workspace-tag">Hoy</span><strong>${data.summary.tasks_today}</strong><p class="workspace-muted">Tareas programadas para hoy</p></article>
      <article class="workspace-stat summary-stat-card summary-stat-yellow" data-module="goals" style="--ws-delay:80ms"><span class="workspace-tag">Metas</span><strong>${data.summary.goals_active}</strong><p class="workspace-muted">Objetivos en curso</p></article>
      <article class="workspace-stat summary-stat-card summary-stat-blue" data-module="calendar" style="--ws-delay:120ms"><span class="workspace-tag">Agenda</span><strong>${data.summary.events_upcoming}</strong><p class="workspace-muted">Eventos por atender</p></article>
      <article class="workspace-stat summary-stat-card summary-stat-green" data-module="calendar" style="--ws-delay:160ms"><span class="workspace-tag">Teams hoy</span><strong>${data.summary.teams_hours_today_label}</strong><p class="workspace-muted">${data.summary.teams_sessions_today} sesiones sincronizadas desde Power Automate</p></article>`;

    renderList(summaryTasks, data.tasks, item => `<div class="workspace-list-item"><strong>${item.title}</strong><p>${item.task_date} · ${item.status} · ${item.priority}</p></div>`, 'Aun no tienes tareas para hoy.');
    renderList(summaryGoals, data.goals, item => `<div class="workspace-list-item"><strong>${item.title}</strong><p>${item.progress_percent}% completado · ${item.status}</p></div>`, 'Todavia no hay metas registradas.');
    renderList(summaryEvents, data.events, item => `<div class="workspace-list-item"><strong>${item.title}</strong><p>${item.start_at} → ${item.end_at}</p></div>`, 'No hay eventos proximos.');
    renderList(summaryNotes, data.notes, item => `<div class="workspace-list-item"><strong>${item.title}</strong><p>${item.content.slice(0, 120)}</p></div>`, 'No hay notas importantes guardadas.');
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}

function renderList(container, items, template, emptyMessage) {
  if (!container) return;
  if (!items || items.length === 0) {
    renderEmptyState(container, emptyMessage);
    return;
  }

  container.innerHTML = items.map(template).join('');
}

async function initDocumentGenerationPage() {
  const form = document.getElementById('docGenerationForm');
  const filesInput = document.getElementById('docFiles');
  const typeInput = document.getElementById('docGenerationType');
  const descriptionInput = document.getElementById('docDescription');
  const statusBox = document.getElementById('docGenerationStatus');
  const list = document.getElementById('docJobsList');
  const reloadBtn = document.getElementById('docHistoryReloadBtn');
  const webhookInput = document.getElementById('docWebhookUrl');
  const saveWebhookBtn = document.getElementById('docSaveWebhookBtn');
  const securityState = document.getElementById('docSecurityState');
  const securityQuickForm = document.getElementById('docSecurityQuickForm');
  const securityQuickInput = document.getElementById('docQuickAccessKey');
  const securityQuickRemoveBtn = document.getElementById('docQuickRemoveKeyBtn');

  const accessBackdrop = document.getElementById('docAccessBackdrop');
  const accessCloseBtn = document.getElementById('docAccessClose');
  const accessKeyInput = document.getElementById('docAccessKeyInput');
  const accessVerifyBtn = document.getElementById('docAccessVerifyBtn');

  const docApiRequest = async (action, options = {}) => {
    const response = await fetch(`../api/generacion_documentos.php?action=${encodeURIComponent(action)}`, {
      method: options.method || 'GET',
      credentials: 'same-origin',
      headers: options.isFormData ? {} : { 'Content-Type': 'application/json' },
      body: options.body || null,
    });

    let payload = null;
    try {
      payload = await response.json();
    } catch (_) {
      payload = null;
    }

    if (!response.ok || !payload || payload.ok === false) {
      const message = payload && payload.message
        ? payload.message
        : `HTTP ${response.status}: Error inesperado en modulo de documentos.`;
      const error = new Error(message);
      error.httpStatus = response.status;
      error.payload = payload || null;
      throw error;
    }

    return payload;
  };

  const openAccessModal = () => {
    if (!accessBackdrop) return;
    accessBackdrop.hidden = false;
    document.body.classList.add('ws-modal-open');
    if (accessKeyInput) {
      accessKeyInput.value = '';
      accessKeyInput.focus();
    }
  };

  const closeAccessModal = () => {
    if (!accessBackdrop) return;
    accessBackdrop.hidden = true;
    document.body.classList.remove('ws-modal-open');
  };

  const setSecurityStateText = security => {
    if (!securityState) return;

    if (!security || !security.configured || !security.enabled) {
      securityState.textContent = 'No hay clave secundaria activa. Puedes configurar una para proteger tus documentos.';
      return;
    }

    securityState.textContent = security.verified
      ? 'Acceso a documentos verificado en esta sesion.'
      : 'Clave secundaria activa. Debes verificarla para ver y descargar documentos.';
  };

  const loadConfigAndSecurity = async () => {
    const config = await docApiRequest('get_config');
    const security = await docApiRequest('security_status');

    if (webhookInput) {
      webhookInput.value = config.webhook_url || '';
    }

    setSecurityStateText(security);

    if (security.configured && security.enabled && !security.verified) {
      openAccessModal();
      return false;
    }

    return true;
  };

  const renderJobs = jobs => {
    if (!list) {
      return;
    }

    if (!jobs || jobs.length === 0) {
      renderEmptyState(list, 'Aun no hay documentos generados. Sube un archivo para iniciar.');
      return;
    }

    list.innerHTML = jobs.map(job => {
      const status = String(job.status || 'pending');
      const badgeClass = `doc-job-status ${status}`;
      const canDownload = status === 'completed' && job.output_file_url;

      return `
        <article class="workspace-list-item doc-job-item">
          <div class="doc-job-head">
            <strong>${job.input_file_name || 'Archivo sin nombre'}</strong>
            <span class="${badgeClass}">${status}</span>
          </div>
          <p>${job.generation_type || 'manual'} · ${job.created_at || ''}</p>
          ${job.error_message ? `<p class="doc-job-error">${job.error_message}</p>` : ''}
          <div class="workspace-actions">
            ${canDownload
              ? `<button type="button" class="btn btn-sm btn-primary" data-download-job="${job.id}"><i class="fas fa-download"></i> Descargar resultado</button>`
              : '<span class="workspace-muted">Sin archivo disponible aun.</span>'}
            <button type="button" class="btn btn-sm btn-outline-danger" data-delete-job="${job.id}"><i class="fas fa-trash"></i> Eliminar</button>
          </div>
        </article>`;
    }).join('');

    list.querySelectorAll('[data-download-job]').forEach(button => {
      button.addEventListener('click', async () => {
        try {
          const result = await docApiRequest('download', {
            method: 'POST',
            body: JSON.stringify({ id: Number(button.dataset.downloadJob || 0) }),
          });

          if (result.download_url) {
            window.open(result.download_url, '_blank', 'noopener');
          }
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-delete-job]').forEach(button => {
      button.addEventListener('click', async () => {
        const jobId = Number(button.dataset.deleteJob || 0);
        if (jobId <= 0) return;

        if (!confirm('¿Estás seguro de que deseas eliminar este documento del historial?')) {
          return;
        }

        try {
          await docApiRequest('delete_job', {
            method: 'POST',
            body: JSON.stringify({ id: jobId }),
          });

          showWorkspaceToast('Documento eliminado correctamente.', 'success');
          await loadHistory();
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });
  };

  const loadHistory = async () => {
    try {
      if (statusBox) {
        statusBox.textContent = 'Consultando historial de documentos...';
      }

      const result = await docApiRequest('history');
      renderJobs(result.jobs || []);

      if (statusBox) {
        statusBox.textContent = 'Historial actualizado.';
      }
    } catch (error) {
      if (error.httpStatus === 403) {
        openAccessModal();
      }
      if (statusBox) {
        statusBox.textContent = error.message;
      }
      renderEmptyState(list, 'No se pudo cargar el historial por falta de verificacion o error de servidor.');
    }
  };

  const bootAllowed = await loadConfigAndSecurity();
  if (bootAllowed) {
    await loadHistory();
  }

  saveWebhookBtn?.addEventListener('click', async () => {
    const webhookUrl = String(webhookInput?.value || '').trim();
    try {
      await docApiRequest('set_webhook_url', {
        method: 'POST',
        body: JSON.stringify({ webhook_url: webhookUrl }),
      });
      showWorkspaceToast('URL de Power Automate guardada.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  securityQuickForm?.addEventListener('submit', async event => {
    event.preventDefault();
    const key = String(securityQuickInput?.value || '').trim();

    try {
      await docApiRequest('set_access_key', {
        method: 'POST',
        body: JSON.stringify({ access_key: key }),
      });
      if (securityQuickInput) securityQuickInput.value = '';
      showWorkspaceToast('Clave secundaria guardada.', 'success');
      await loadConfigAndSecurity();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  securityQuickRemoveBtn?.addEventListener('click', async () => {
    try {
      await docApiRequest('remove_access_key', {
        method: 'POST',
        body: JSON.stringify({}),
      });
      showWorkspaceToast('Clave secundaria eliminada.', 'success');
      await loadConfigAndSecurity();
      await loadHistory();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  form?.addEventListener('submit', async event => {
    event.preventDefault();

    const files = filesInput?.files;
    if (!files || files.length === 0) {
      showWorkspaceToast('Adjunta al menos un archivo para generar el documento.', 'warning');
      return;
    }

    const data = new FormData();
    Array.from(files).forEach(file => data.append('files[]', file));
    data.append('generation_type', String(typeInput?.value || 'manual'));
    data.append('description', String(descriptionInput?.value || '').trim());

    try {
      if (statusBox) {
        statusBox.textContent = 'Enviando archivos a Power Automate para generacion...';
      }

      await docApiRequest('create', {
        method: 'POST',
        body: data,
        isFormData: true,
      });

      if (filesInput) filesInput.value = '';
      if (descriptionInput) descriptionInput.value = '';

      showWorkspaceToast('Solicitud enviada. Revisa el historial para ver resultados.', 'success');
      await loadHistory();
    } catch (error) {
      if (error.httpStatus === 403) {
        openAccessModal();
      }

      if (statusBox) {
        statusBox.textContent = error.message;
      }
      showWorkspaceToast(error.message, 'danger');
    }
  });

  reloadBtn?.addEventListener('click', loadHistory);

  accessVerifyBtn?.addEventListener('click', async () => {
    const key = String(accessKeyInput?.value || '').trim();
    if (key.length < 6) {
      showWorkspaceToast('La clave secundaria debe tener al menos 6 caracteres.', 'warning');
      accessKeyInput?.focus();
      return;
    }

    try {
      await docApiRequest('verify_access_key', {
        method: 'POST',
        body: JSON.stringify({ access_key: key }),
      });

      closeAccessModal();
      showWorkspaceToast('Acceso verificado para documentos sensibles.', 'success');
      await loadConfigAndSecurity();
      await loadHistory();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  accessCloseBtn?.addEventListener('click', closeAccessModal);
  accessBackdrop?.addEventListener('click', event => {
    if (event.target === accessBackdrop) {
      closeAccessModal();
    }
  });
}

async function initProfilePage() {
  const profileForm = document.getElementById('profileForm');
  const publicProfileForm = document.getElementById('publicProfileForm');
  const photoForm = document.getElementById('photoForm');
  const preview = document.getElementById('profilePhotoPreview');
  const previewFallback = document.getElementById('profilePhotoFallback');
  const identityPhoto = document.getElementById('profileIdentityPhoto');
  const identityFallback = document.getElementById('profileIdentityFallback');
  const identityName = document.getElementById('profileIdentityName');
  const identityEmail = document.getElementById('profileIdentityEmail');
  const photoInput = document.getElementById('profilePhotoInput');
  const uploadDropzone = document.getElementById('profileUploadDropzone');
  const photoResetBtn = document.getElementById('profilePhotoResetBtn');
  const photoMeta = document.getElementById('profilePhotoMeta');
  const photoSubmitBtn = document.getElementById('profilePhotoSubmitBtn');
  const passwordInput = document.getElementById('profilePassword');
  const passwordConfirmInput = document.getElementById('profilePasswordConfirm');

  const twoFAStateBadge = document.getElementById('profile2FAStateBadge');
  const twoFATag = document.getElementById('profile2FATag');
  const twoFAGenerateBtn = document.getElementById('profile2FAGenerateBtn');
  const twoFADisableBtn = document.getElementById('profile2FADisableBtn');
  const twoFASecretBox = document.getElementById('profile2FASecretBox');
  const twoFASecretInput = document.getElementById('profile2FASecret');
  const twoFACopyBtn = document.getElementById('profile2FACopyBtn');
  const twoFAQRWrap = document.getElementById('profile2FAQRWrap');
  const twoFAQR = document.getElementById('profile2FAQR');
  const twoFAConfirmBox = document.getElementById('profile2FAConfirmBox');
  const twoFACode = document.getElementById('profile2FACode');
  const twoFAEnableBtn = document.getElementById('profile2FAEnableBtn');
  const twoFAToggle = document.getElementById('profile2FAToggle');
  const twoFAHint = document.getElementById('profile2FAHint');
  const docSecurityForm = document.getElementById('profileDocSecurityForm');
  const docSecurityInput = document.getElementById('profileDocAccessKey');
  const docSecurityConfirmInput = document.getElementById('profileDocAccessKeyConfirm');
  const docSecurityState = document.getElementById('profileDocSecurityState');
  const docSecurityTag = document.getElementById('profileDocSecurityTag');
  const docSecurityRemoveBtn = document.getElementById('profileDocSecurityRemoveBtn');
  const developerAvatar = document.getElementById('profileDeveloperAvatar');
  const developerAvatarFallback = document.getElementById('profileDeveloperAvatarFallback');
  const developerName = document.getElementById('profileDeveloperName');
  const developerRole = document.getElementById('profileDeveloperRole');
  const developerOwnerEmail = document.getElementById('profileDeveloperOwnerEmail');
  const developerOwnerTag = document.getElementById('profileDeveloperOwnerTag');
  const developerState = document.getElementById('profileDeveloperState');
  const developerPhotoForm = document.getElementById('profileDeveloperPhotoForm');
  const developerPhotoInput = document.getElementById('profileDeveloperPhotoInput');
  const developerPhotoSaveBtn = document.getElementById('profileDeveloperPhotoSaveBtn');
  const developerOwnerForm = document.getElementById('profileDeveloperOwnerForm');
  const developerOwnerSelect = document.getElementById('profileDeveloperOwnerUserId');
  const developerOwnerSaveBtn = document.getElementById('profileDeveloperOwnerSaveBtn');
  const developerPromoteAdminForm = document.getElementById('profileDeveloperPromoteAdminForm');
  const developerPromoteAdminEmail = document.getElementById('profileDeveloperPromoteAdminEmail');
  const developerPromoteAdminBtn = document.getElementById('profileDeveloperPromoteAdminBtn');

  const defaultPhotoMetaLabel = 'PNG/JPG/WebP hasta 5MB';
  let currentPhotoUrl = '';
  let stagedPhotoFile = null;
  let stagedPhotoObjectUrl = '';
  let is2FAEnabled = false;
  let is2FASetupPending = false;
  let isProgrammaticToggleUpdate = false;
  let canManageDeveloperProfile = false;

  const getAvatarFallbackText = name => {
    const safeName = String(name || '').trim();
    if (!safeName) return 'U';
    const [first, second] = safeName.split(/\s+/);
    return `${(first || 'U').charAt(0)}${(second || '').charAt(0)}`.toUpperCase();
  };

  const setPhotoView = (src, fallbackText) => {
    const hasPhoto = Boolean(src);

    if (preview) {
      if (hasPhoto) {
        preview.src = src;
        preview.style.display = 'block';
      } else {
        preview.removeAttribute('src');
        preview.style.display = 'none';
      }
    }

    if (identityPhoto) {
      if (hasPhoto) {
        identityPhoto.src = src;
        identityPhoto.style.display = 'block';
      } else {
        identityPhoto.removeAttribute('src');
        identityPhoto.style.display = 'none';
      }
    }

    if (previewFallback) {
      previewFallback.style.display = hasPhoto ? 'none' : 'grid';
      previewFallback.textContent = hasPhoto ? '' : fallbackText;
    }

    if (identityFallback) {
      identityFallback.style.display = hasPhoto ? 'none' : 'grid';
      identityFallback.textContent = hasPhoto ? '' : fallbackText;
    }

    const topbarChip = document.querySelector('.workspace-user-chip');
    if (topbarChip) {
      const leadingNode = topbarChip.querySelector('img, .workspace-user-fallback');
      if (leadingNode) {
        leadingNode.outerHTML = hasPhoto
          ? `<img src="${src}" alt="Foto de perfil">`
          : '<span class="workspace-user-fallback"><i class="fas fa-user"></i></span>';
      }
    }
  };

  const syncWorkspaceIdentity = (fullName, email) => {
    const safeName = String(fullName || '').trim() || 'Usuario del Workspace';
    const safeEmail = String(email || '').trim();

    if (identityName) identityName.textContent = safeName;
    if (identityEmail) identityEmail.textContent = safeEmail;

    const headerBadgeLabel = document.querySelector('.header-badge span');
    if (headerBadgeLabel) headerBadgeLabel.textContent = safeName;

    const chipName = document.querySelector('.workspace-user-chip strong');
    if (chipName) chipName.textContent = safeName;

    const chipEmail = document.querySelector('.workspace-user-chip small');
    if (chipEmail) chipEmail.textContent = safeEmail;

    if (!currentPhotoUrl && !stagedPhotoFile) {
      setPhotoView('', getAvatarFallbackText(safeName));
    }
  };

  const set2FAVisualState = enabled => {
    is2FAEnabled = Boolean(enabled);

    if (twoFAStateBadge) {
      twoFAStateBadge.classList.toggle('is-enabled', is2FAEnabled);
      twoFAStateBadge.innerHTML = is2FAEnabled
        ? '<i class="fas fa-shield-check"></i> 2FA activado'
        : '<i class="fas fa-shield"></i> 2FA desactivado';
    }

    if (twoFATag) {
      twoFATag.classList.toggle('is-enabled', is2FAEnabled);
      twoFATag.innerHTML = is2FAEnabled
        ? '<i class="fas fa-lock"></i> Cuenta reforzada'
        : '<i class="fas fa-lock-open"></i> Configuracion requerida';
    }

    if (twoFADisableBtn) {
      twoFADisableBtn.disabled = !is2FAEnabled;
    }

    if (twoFAToggle) {
      isProgrammaticToggleUpdate = true;
      twoFAToggle.checked = is2FAEnabled;
      isProgrammaticToggleUpdate = false;
    }

    if (twoFAHint) {
      twoFAHint.textContent = is2FAEnabled
        ? 'Tu cuenta ya esta protegida con 2FA. Puedes desactivar cuando lo necesites.'
        : 'Activa para generar QR y registrar la app autenticadora.';
    }
  };

  const reset2FASetupView = () => {
    is2FASetupPending = false;
    if (twoFASecretBox) twoFASecretBox.hidden = true;
    if (twoFAQRWrap) twoFAQRWrap.hidden = true;
    if (twoFAConfirmBox) twoFAConfirmBox.hidden = true;
    if (twoFASecretInput) twoFASecretInput.value = '';
    if (twoFAQR) twoFAQR.removeAttribute('src');
    if (twoFACode) twoFACode.value = '';
  };

  const setDocumentSecurityVisualState = security => {
    if (!docSecurityTag || !docSecurityState) {
      return;
    }

    const configured = Boolean(security && security.configured);
    docSecurityTag.innerHTML = configured
      ? '<i class="fas fa-lock"></i> Clave configurada'
      : '<i class="fas fa-lock-open"></i> Sin clave';

    docSecurityState.textContent = configured
      ? `Clave secundaria activa${security.last_verified_at ? ` · ultima verificacion ${security.last_verified_at}` : ''}.`
      : 'Aun no tienes clave secundaria para documentos sensibles.';
  };

  const loadDocumentSecurityStatus = async () => {
    if (!docSecurityState) {
      return;
    }

    try {
      const result = await apiRequest('document_security_status');
      setDocumentSecurityVisualState(result.security || null);
    } catch (error) {
      docSecurityState.textContent = error.message || 'No se pudo consultar el estado de seguridad de documentos.';
    }
  };

  const setDeveloperControlsEnabled = enabled => {
    if (developerPhotoInput) developerPhotoInput.disabled = !enabled;
    if (developerPhotoSaveBtn) developerPhotoSaveBtn.disabled = !enabled;
    if (developerOwnerSelect) developerOwnerSelect.disabled = !enabled;
    if (developerOwnerSaveBtn) developerOwnerSaveBtn.disabled = !enabled;
    if (developerPromoteAdminEmail) developerPromoteAdminEmail.disabled = !enabled;
    if (developerPromoteAdminBtn) developerPromoteAdminBtn.disabled = !enabled;
  };

  const renderDeveloperProfile = (developerProfile, adminUsers = [], canManage = false) => {
    canManageDeveloperProfile = Boolean(canManage);

    const photoUrl = String(developerProfile?.photo_url || '').trim();
    if (developerAvatar) {
      if (photoUrl) {
        developerAvatar.src = photoUrl;
        developerAvatar.style.display = 'block';
      } else {
        developerAvatar.removeAttribute('src');
        developerAvatar.style.display = 'none';
      }
    }

    if (developerAvatarFallback) {
      developerAvatarFallback.style.display = photoUrl ? 'none' : 'inline-block';
    }

    if (developerName) {
      developerName.textContent = String(developerProfile?.display_name || 'Desarrollador');
    }

    if (developerRole) {
      developerRole.textContent = String(developerProfile?.role_label || 'Developer');
    }

    if (developerOwnerEmail) {
      developerOwnerEmail.textContent = `Propietario: ${String(developerProfile?.owner_email || 'n/a')}`;
    }

    if (developerOwnerTag) {
      developerOwnerTag.innerHTML = canManageDeveloperProfile
        ? '<i class="fas fa-user-shield"></i> Eres admin propietario'
        : '<i class="fas fa-lock"></i> Solo admin propietario';
      developerOwnerTag.classList.toggle('is-enabled', canManageDeveloperProfile);
    }

    if (developerOwnerSelect) {
      const currentOwnerId = Number(developerProfile?.owner_user_id || 0);
      const admins = Array.isArray(adminUsers) ? adminUsers : [];
      developerOwnerSelect.innerHTML = admins.length
        ? admins.map(admin => `<option value="${Number(admin.id)}">${String(admin.full_name || admin.email)} (${String(admin.email || '')})</option>`).join('')
        : '<option value="">No hay admins disponibles</option>';

      if (currentOwnerId > 0) {
        developerOwnerSelect.value = String(currentOwnerId);
      }
    }

    if (developerState) {
      developerState.textContent = canManageDeveloperProfile
        ? 'Puedes actualizar la foto del desarrollador o transferir la propiedad a otro admin.'
        : 'No tienes permisos para editar este perfil. Solo el admin propietario puede hacerlo.';
    }

    setDeveloperControlsEnabled(canManageDeveloperProfile);
  };

  const validatePhotoFile = file => {
    if (!file) {
      return 'Selecciona una imagen primero.';
    }

    const allowedMimeTypes = ['image/jpeg', 'image/png', 'image/webp'];
    if (!allowedMimeTypes.includes(String(file.type || '').toLowerCase())) {
      return 'Solo se permiten archivos PNG, JPG o WebP.';
    }

    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
      return 'La imagen supera 5MB. Usa una mas liviana.';
    }

    return '';
  };

  const stagePhoto = file => {
    const error = validatePhotoFile(file);
    if (error) {
      showWorkspaceToast(error, 'warning');
      return;
    }

    if (stagedPhotoObjectUrl) {
      URL.revokeObjectURL(stagedPhotoObjectUrl);
      stagedPhotoObjectUrl = '';
    }

    stagedPhotoFile = file;
    const objectUrl = URL.createObjectURL(file);
    stagedPhotoObjectUrl = objectUrl;
    const fallbackText = getAvatarFallbackText(identityName?.textContent || '');
    setPhotoView(objectUrl, fallbackText);

    if (photoMeta) {
      const sizeInMb = (file.size / (1024 * 1024)).toFixed(2);
      photoMeta.textContent = `${file.name} · ${sizeInMb} MB`;
    }
  };

  uploadDropzone?.addEventListener('click', () => {
    photoInput?.click();
  });

  uploadDropzone?.addEventListener('keydown', event => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      photoInput?.click();
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    uploadDropzone?.addEventListener(eventName, event => {
      event.preventDefault();
      uploadDropzone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    uploadDropzone?.addEventListener(eventName, event => {
      event.preventDefault();
      uploadDropzone.classList.remove('drag-over');

      if (eventName === 'drop' && event.dataTransfer?.files?.length) {
        stagePhoto(event.dataTransfer.files[0]);
      }
    });
  });

  photoInput?.addEventListener('change', event => {
    const file = event.target.files?.[0] || null;
    if (!file) return;
    stagePhoto(file);
  });

  photoResetBtn?.addEventListener('click', () => {
    stagedPhotoFile = null;
    if (photoInput) photoInput.value = '';
    const fallbackText = getAvatarFallbackText(identityName?.textContent || '');
    setPhotoView(currentPhotoUrl, fallbackText);
    if (photoMeta) {
      photoMeta.textContent = defaultPhotoMetaLabel;
    }
    if (stagedPhotoObjectUrl) {
      URL.revokeObjectURL(stagedPhotoObjectUrl);
      stagedPhotoObjectUrl = '';
    }
  });

  try {
    const data = await apiRequest('profile_get');

    document.getElementById('profileFullName').value = data.user.full_name || '';
    document.getElementById('profileEmail').value = data.user.email || '';
    document.getElementById('publicDisplayName').value = data.public_profile?.display_name || '';
    document.getElementById('publicRoleTitle').value = data.public_profile?.role_title || '';
    document.getElementById('publicCompanyName').value = data.public_profile?.company_name || '';
    document.getElementById('publicBio').value = data.public_profile?.bio || '';

    syncWorkspaceIdentity(data.user.full_name, data.user.email);

    currentPhotoUrl = data.user.profile_photo_path
      ? `../${data.user.profile_photo_path}`
      : (data.public_profile?.photo_url || '');

    setPhotoView(currentPhotoUrl, getAvatarFallbackText(data.user.full_name));
    set2FAVisualState(Boolean(data.user.two_factor_enabled));
    reset2FASetupView();
    renderDeveloperProfile(data.developer_profile, data.admin_users, Boolean(data.can_manage_developer_profile));
    await loadDocumentSecurityStatus();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }

  profileForm?.addEventListener('submit', async event => {
    event.preventDefault();
    if (passwordInput && passwordConfirmInput && passwordInput.value !== passwordConfirmInput.value) {
      showWorkspaceToast('Las contrasenas no coinciden.', 'warning');
      return;
    }

    try {
      const result = await apiRequest('profile_update', {
        method: 'POST',
        body: JSON.stringify({
          full_name: document.getElementById('profileFullName').value,
          email: document.getElementById('profileEmail').value,
          password: passwordInput?.value || '',
        }),
      });

      syncWorkspaceIdentity(result.user?.full_name, result.user?.email);

      if (passwordInput) passwordInput.value = '';
      if (passwordConfirmInput) passwordConfirmInput.value = '';
      showWorkspaceToast('Perfil personal actualizado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  publicProfileForm?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('public_profile_update', {
        method: 'POST',
        body: JSON.stringify({
          display_name: document.getElementById('publicDisplayName').value,
          role_title: document.getElementById('publicRoleTitle').value,
          company_name: document.getElementById('publicCompanyName').value,
          bio: document.getElementById('publicBio').value,
        }),
      });
      showWorkspaceToast('Perfil publico actualizado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  photoForm?.addEventListener('submit', async event => {
    event.preventDefault();

    const fileToUpload = stagedPhotoFile || photoInput?.files?.[0] || null;
    const error = validatePhotoFile(fileToUpload);
    if (error) {
      showWorkspaceToast(error, 'warning');
      return;
    }

    const formData = new FormData();
    formData.append('photo', fileToUpload);
    formData.append('make_public_profile', document.getElementById('makePublicProfilePhoto').checked ? '1' : '0');

    try {
      if (photoSubmitBtn) {
        photoSubmitBtn.disabled = true;
      }

      const data = await apiRequest('profile_photo_upload', {
        method: 'POST',
        body: formData,
        isFormData: true,
      });

      currentPhotoUrl = data.photo_url || '';
      stagedPhotoFile = null;
      if (photoInput) photoInput.value = '';
      if (stagedPhotoObjectUrl) {
        URL.revokeObjectURL(stagedPhotoObjectUrl);
        stagedPhotoObjectUrl = '';
      }
      setPhotoView(currentPhotoUrl, getAvatarFallbackText(identityName?.textContent || ''));
      if (photoMeta) {
        photoMeta.textContent = defaultPhotoMetaLabel;
      }
      showWorkspaceToast('Foto de perfil actualizada.', 'success');
    } catch (uploadError) {
      showWorkspaceToast(uploadError.message, 'danger');
    } finally {
      if (photoSubmitBtn) {
        photoSubmitBtn.disabled = false;
      }
    }
  });

  developerPhotoForm?.addEventListener('submit', async event => {
    event.preventDefault();

    if (!canManageDeveloperProfile) {
      showWorkspaceToast('Solo el admin propietario puede modificar la foto del desarrollador.', 'warning');
      return;
    }

    const fileToUpload = developerPhotoInput?.files?.[0] || null;
    const error = validatePhotoFile(fileToUpload);
    if (error) {
      showWorkspaceToast(error, 'warning');
      return;
    }

    const formData = new FormData();
    formData.append('photo', fileToUpload);

    try {
      if (developerPhotoSaveBtn) developerPhotoSaveBtn.disabled = true;

      const result = await apiRequest('developer_profile_photo_upload', {
        method: 'POST',
        body: formData,
        isFormData: true,
      });

      if (developerPhotoInput) developerPhotoInput.value = '';
      const refreshed = await apiRequest('profile_get');
      renderDeveloperProfile(refreshed.developer_profile, refreshed.admin_users, Boolean(refreshed.can_manage_developer_profile));
      showWorkspaceToast(result.message || 'Foto del desarrollador actualizada.', 'success');
    } catch (uploadError) {
      showWorkspaceToast(uploadError.message, 'danger');
    } finally {
      if (developerPhotoSaveBtn) developerPhotoSaveBtn.disabled = !canManageDeveloperProfile;
    }
  });

  developerOwnerForm?.addEventListener('submit', async event => {
    event.preventDefault();

    if (!canManageDeveloperProfile) {
      showWorkspaceToast('Solo el admin propietario puede transferir la propiedad.', 'warning');
      return;
    }

    const targetUserId = Number(developerOwnerSelect?.value || 0);
    if (targetUserId <= 0) {
      showWorkspaceToast('Selecciona un admin valido.', 'warning');
      return;
    }

    try {
      if (developerOwnerSaveBtn) developerOwnerSaveBtn.disabled = true;
      const result = await apiRequest('developer_profile_transfer_owner', {
        method: 'POST',
        body: JSON.stringify({ target_user_id: targetUserId }),
      });

      showWorkspaceToast(result.message || 'Propiedad transferida.', 'success');

      const refreshed = await apiRequest('profile_get');
      renderDeveloperProfile(refreshed.developer_profile, refreshed.admin_users, Boolean(refreshed.can_manage_developer_profile));
    } catch (transferError) {
      showWorkspaceToast(transferError.message, 'danger');
    } finally {
      if (developerOwnerSaveBtn) developerOwnerSaveBtn.disabled = !canManageDeveloperProfile;
    }
  });

  developerPromoteAdminForm?.addEventListener('submit', async event => {
    event.preventDefault();

    if (!canManageDeveloperProfile) {
      showWorkspaceToast('Solo el admin propietario puede crear nuevos admins.', 'warning');
      return;
    }

    const email = String(developerPromoteAdminEmail?.value || '').trim().toLowerCase();
    if (!email) {
      showWorkspaceToast('Ingresa el correo de la cuenta que quieres convertir a admin.', 'warning');
      return;
    }

    try {
      if (developerPromoteAdminBtn) developerPromoteAdminBtn.disabled = true;

      const result = await apiRequest('developer_profile_promote_admin', {
        method: 'POST',
        body: JSON.stringify({ email }),
      });

      showWorkspaceToast(result.message || 'Cuenta promovida a admin.', 'success');
      if (developerPromoteAdminEmail) developerPromoteAdminEmail.value = '';

      const refreshed = await apiRequest('profile_get');
      renderDeveloperProfile(refreshed.developer_profile, refreshed.admin_users, Boolean(refreshed.can_manage_developer_profile));
    } catch (promoteError) {
      showWorkspaceToast(promoteError.message, 'danger');
    } finally {
      if (developerPromoteAdminBtn) developerPromoteAdminBtn.disabled = !canManageDeveloperProfile;
    }
  });

  twoFAGenerateBtn?.addEventListener('click', async () => {
    try {
      const data = await apiRequest('profile_2fa_generate', {
        method: 'POST',
        body: JSON.stringify({}),
      });

      is2FASetupPending = true;
      if (twoFASecretBox) twoFASecretBox.hidden = false;
      if (twoFASecretInput) twoFASecretInput.value = data.secret || '';
      if (twoFAQRWrap) twoFAQRWrap.hidden = false;
      if (twoFAQR) twoFAQR.src = data.qr_url || '';
      if (twoFAConfirmBox) twoFAConfirmBox.hidden = false;
      if (twoFACode) {
        twoFACode.value = '';
        twoFACode.focus();
      }

      showWorkspaceToast('QR generado. Escanealo y confirma con el codigo.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
      reset2FASetupView();
    }
  });

  twoFACopyBtn?.addEventListener('click', async () => {
    if (!twoFASecretInput?.value) return;
    try {
      await navigator.clipboard.writeText(twoFASecretInput.value);
      showWorkspaceToast('Clave 2FA copiada.', 'success');
    } catch (error) {
      showWorkspaceToast('No se pudo copiar automaticamente. Copiala manualmente.', 'warning');
    }
  });

  twoFAEnableBtn?.addEventListener('click', async () => {
    const code = String(twoFACode?.value || '').replace(/\D/g, '').slice(0, 6);
    if (twoFACode) {
      twoFACode.value = code;
    }

    if (code.length !== 6) {
      showWorkspaceToast('Ingresa un codigo 2FA de 6 digitos.', 'warning');
      twoFACode?.focus();
      return;
    }

    try {
      const result = await apiRequest('profile_2fa_enable', {
        method: 'POST',
        body: JSON.stringify({ code }),
      });

      set2FAVisualState(Boolean(result.user?.two_factor_enabled));
      reset2FASetupView();
      showWorkspaceToast(result.message || '2FA activado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  twoFACode?.addEventListener('input', event => {
    const cleanCode = String(event.target.value || '').replace(/\D/g, '').slice(0, 6);
    event.target.value = cleanCode;
  });

  twoFADisableBtn?.addEventListener('click', async () => {
    if (!is2FAEnabled && !is2FASetupPending) {
      reset2FASetupView();
      showWorkspaceToast('2FA ya estaba desactivado.', 'primary');
      return;
    }

    try {
      const result = await apiRequest('profile_2fa_disable', {
        method: 'POST',
        body: JSON.stringify({}),
      });

      set2FAVisualState(Boolean(result.user?.two_factor_enabled));
      reset2FASetupView();
      showWorkspaceToast(result.message || '2FA desactivado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  twoFAToggle?.addEventListener('change', async event => {
    if (isProgrammaticToggleUpdate) {
      return;
    }

    const turnOn = Boolean(event.target.checked);

    if (turnOn) {
      if (is2FAEnabled) {
        showWorkspaceToast('2FA ya esta activado.', 'primary');
        return;
      }

      twoFAGenerateBtn?.click();
      return;
    }

    if (is2FAEnabled) {
      twoFADisableBtn?.click();
      return;
    }

    reset2FASetupView();
    showWorkspaceToast('Configuracion 2FA cancelada.', 'primary');
  });

  docSecurityForm?.addEventListener('submit', async event => {
    event.preventDefault();

    const accessKey = String(docSecurityInput?.value || '').trim();
    const accessKeyConfirm = String(docSecurityConfirmInput?.value || '').trim();

    if (accessKey.length < 6) {
      showWorkspaceToast('La clave secundaria debe tener al menos 6 caracteres.', 'warning');
      docSecurityInput?.focus();
      return;
    }

    if (accessKey !== accessKeyConfirm) {
      showWorkspaceToast('Las claves secundarias no coinciden.', 'warning');
      docSecurityConfirmInput?.focus();
      return;
    }

    try {
      await apiRequest('document_security_set', {
        method: 'POST',
        body: JSON.stringify({ access_key: accessKey }),
      });

      if (docSecurityInput) docSecurityInput.value = '';
      if (docSecurityConfirmInput) docSecurityConfirmInput.value = '';
      showWorkspaceToast('Clave secundaria guardada.', 'success');
      await loadDocumentSecurityStatus();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  docSecurityRemoveBtn?.addEventListener('click', async () => {
    try {
      await apiRequest('document_security_remove', {
        method: 'POST',
        body: JSON.stringify({}),
      });

      if (docSecurityInput) docSecurityInput.value = '';
      if (docSecurityConfirmInput) docSecurityConfirmInput.value = '';
      showWorkspaceToast('Clave secundaria eliminada.', 'success');
      await loadDocumentSecurityStatus();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });
}

async function initNotesPage() {
  const form = document.getElementById('noteForm');
  const list = document.getElementById('notesList');
  const searchInput = document.getElementById('notesSearchInput');
  const colorFilter = document.getElementById('notesColorFilter');
  const scopeFilter = document.getElementById('notesScopeFilter');
  const headerMetrics = document.getElementById('notesHeaderMetrics');

  const state = {
    notes: [],
    activeNoteId: null,
    search: '',
    color: 'all',
    scope: 'all',
    commentsByNote: new Map(),
    sharesByNote: new Map(),
  };

  const sanitize = value => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

  const nl2br = value => sanitize(value).replace(/\n/g, '<br>');

  const formatStamp = value => {
    if (!value) return 'sin fecha';
    const date = new Date(value.replace(' ', 'T'));
    if (Number.isNaN(date.getTime())) return 'sin fecha';
    return date.toLocaleString('es-ES', { dateStyle: 'short', timeStyle: 'short' });
  };

  const excerpt = text => {
    const raw = String(text ?? '').trim();
    if (raw.length <= 190) return raw;
    return `${raw.slice(0, 190)}...`;
  };

  const getFilteredNotes = () => state.notes.filter(note => {
    const matchesSearch = state.search === ''
      || note.title.toLowerCase().includes(state.search)
      || note.content.toLowerCase().includes(state.search);
    const matchesColor = state.color === 'all' || note.color === state.color;
    const matchesScope = state.scope === 'all' || note.note_scope === state.scope;
    return matchesSearch && matchesColor && matchesScope;
  });

  const buildCommentNodes = comments => {
    const sorted = Array.isArray(comments) ? [...comments] : [];
    sorted.sort((a, b) => {
      const first = new Date(String(a.created_at ?? '').replace(' ', 'T')).getTime() || 0;
      const second = new Date(String(b.created_at ?? '').replace(' ', 'T')).getTime() || 0;
      if (first === second) return Number(a.id) - Number(b.id);
      return first - second;
    });

    const byParent = new Map();
    sorted.forEach(comment => {
      const parent = Number(comment.parent_comment_id || 0);
      if (!byParent.has(parent)) {
        byParent.set(parent, []);
      }
      byParent.get(parent).push(comment);
    });

    const renderNode = (comment, depth) => {
      const children = byParent.get(Number(comment.id)) || [];
      const childHtml = children.map(child => renderNode(child, depth + 1)).join('');
      return `
        <article class="note-comment" style="--comment-depth:${Math.min(depth, 4)};">
          <div class="note-comment-head">
            <strong>${sanitize(comment.author_name || comment.author_email || 'Colaborador')}</strong>
            <span>${formatStamp(comment.created_at)}</span>
          </div>
          <p class="note-comment-content">${nl2br(comment.content)}</p>
          <div class="note-comment-actions">
            <button class="btn btn-sm btn-outline-primary" data-toggle-reply="${comment.id}">Responder</button>
            <button class="btn btn-sm btn-outline-danger" data-delete-comment="${comment.id}">Eliminar</button>
          </div>
          <form class="note-reply-form" data-reply-form="${comment.id}" hidden>
            <textarea class="form-control" rows="2" maxlength="3000" placeholder="Responder comentario..."></textarea>
            <div class="workspace-actions">
              <button type="submit" class="btn btn-sm btn-primary">Publicar</button>
              <button type="button" class="btn btn-sm btn-outline-secondary" data-cancel-reply="${comment.id}">Cancelar</button>
            </div>
          </form>
          ${childHtml ? `<div class="note-comment-children">${childHtml}</div>` : ''}
        </article>`;
    };

    return (byParent.get(0) || []).map(root => renderNode(root, 0)).join('');
  };

  const renderShares = note => {
    if (Number(note.is_owner || 0) !== 1) {
      return '<p class="workspace-muted">Esta nota es compartida contigo; no puedes gestionar invitaciones.</p>';
    }

    const shares = state.sharesByNote.get(Number(note.id)) || [];
    const items = shares.length
      ? shares.map(share => `
        <li>
          <span>${sanitize(share.invited_email)}</span>
          <small>${Number(share.is_active) === 1 ? 'Activa' : 'Revocada'}</small>
          ${Number(share.is_active) === 1
            ? `<button class="btn btn-sm btn-outline-danger" data-revoke-share="${share.id}">Revocar</button>`
            : ''}
        </li>`).join('')
      : '<li class="workspace-muted">No hay invitaciones registradas.</li>';

    return `
      <form class="note-share-form" data-share-form="${note.id}">
        <label class="form-label">Compartir por correo</label>
        <div class="note-share-inline">
          <input type="email" class="form-control" placeholder="correo@dominio.com" maxlength="190" required>
          <button type="submit" class="btn btn-sm btn-outline-primary">Invitar</button>
        </div>
      </form>
      <ul class="note-share-list">${items}</ul>`;
  };

  const renderNoteDetails = note => {
    if (Number(state.activeNoteId) !== Number(note.id)) return '';

    const comments = state.commentsByNote.get(Number(note.id)) || [];
    const commentsHtml = comments.length
      ? buildCommentNodes(comments)
      : '<p class="workspace-muted">Aun no hay comentarios. Inicia el hilo de seguimiento.</p>';

    return `
      <section class="workspace-note-detail">
        <div class="workspace-note-detail-grid">
          <div class="workspace-note-content-full">${nl2br(note.content)}</div>
          <div class="workspace-note-share-box">
            <h5>Invitaciones</h5>
            ${renderShares(note)}
          </div>
        </div>
        <div class="workspace-note-comments">
          <h5>Comentarios jerarquicos</h5>
          <form class="note-comment-form" data-comment-form="${note.id}">
            <textarea class="form-control" rows="2" maxlength="3000" placeholder="Escribe un comentario principal..." required></textarea>
            <button type="submit" class="btn btn-sm btn-primary">Comentar</button>
          </form>
          <div class="note-comments-thread">${commentsHtml}</div>
        </div>
      </section>`;
  };

  const renderNotes = () => {
    if (!list) return;

    const filtered = getFilteredNotes();
    if (headerMetrics) {
      headerMetrics.textContent = `${filtered.length} notas visibles · ${state.notes.length} total`;
    }

    if (!filtered.length) {
      list.innerHTML = '<p class="workspace-empty">No hay notas con esos filtros. Ajusta busqueda, color o tipo.</p>';
      return;
    }

    list.innerHTML = filtered.map((note, index) => `
      <article class="workspace-note compact" data-color="${sanitize(note.color || 'blue')}" style="--stagger:${index};">
        <div class="workspace-note-top">
          <div>
            <h4>${sanitize(note.title)}</h4>
            <p>${sanitize(excerpt(note.content))}</p>
          </div>
          <div class="workspace-note-meta">
            ${Number(note.is_pinned) === 1 ? '<span class="workspace-tag note-tag-pin"><i class="fas fa-thumbtack"></i> Fijada</span>' : ''}
            <span class="workspace-tag note-tag-scope">${note.note_scope === 'shared' ? 'Compartida' : 'Propia'}</span>
          </div>
        </div>
        <div class="workspace-note-foot">
          <small>Actualizada ${formatStamp(note.updated_at)}</small>
          <div class="workspace-actions">
            <span class="workspace-tag"><i class="fas fa-comments"></i> ${Number(note.comments_count || 0)}</span>
            <button class="btn btn-sm btn-outline-primary" data-open-note="${note.id}">${Number(state.activeNoteId) === Number(note.id) ? 'Ocultar' : 'Abrir'}</button>
            ${Number(note.can_edit || 0) === 1
              ? `<button class="btn btn-sm btn-outline-danger" data-delete-note="${note.id}"><i class="fas fa-trash"></i></button>`
              : ''}
          </div>
        </div>
        ${renderNoteDetails(note)}
      </article>
    `).join('');

    list.querySelectorAll('[data-open-note]').forEach(button => {
      button.addEventListener('click', async () => {
        const noteId = Number(button.dataset.openNote || 0);
        if (state.activeNoteId === noteId) {
          state.activeNoteId = null;
          renderNotes();
          return;
        }

        state.activeNoteId = noteId;
        try {
          await loadNoteDetails(noteId);
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
        renderNotes();
      });
    });

    list.querySelectorAll('[data-delete-note]').forEach(button => {
      button.addEventListener('click', async () => {
        try {
          await apiRequest('note_delete', { method: 'POST', body: JSON.stringify({ id: button.dataset.deleteNote }) });
          showWorkspaceToast('Nota eliminada.', 'success');
          await loadNotes();
          renderNotes();
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-comment-form]').forEach(formElement => {
      formElement.addEventListener('submit', async event => {
        event.preventDefault();
        const noteId = Number(formElement.dataset.commentForm || 0);
        const textarea = formElement.querySelector('textarea');
        const content = String(textarea?.value || '').trim();
        if (!content) return;

        try {
          await apiRequest('note_comment_add', {
            method: 'POST',
            body: JSON.stringify({ note_id: noteId, content }),
          });
          if (textarea) textarea.value = '';
          await loadNoteDetails(noteId);
          await loadNotes();
          renderNotes();
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-toggle-reply]').forEach(button => {
      button.addEventListener('click', () => {
        const commentId = button.dataset.toggleReply;
        const replyForm = list.querySelector(`[data-reply-form="${commentId}"]`);
        if (!replyForm) return;
        replyForm.hidden = !replyForm.hidden;
      });
    });

    list.querySelectorAll('[data-cancel-reply]').forEach(button => {
      button.addEventListener('click', () => {
        const commentId = button.dataset.cancelReply;
        const replyForm = list.querySelector(`[data-reply-form="${commentId}"]`);
        if (!replyForm) return;
        const textarea = replyForm.querySelector('textarea');
        if (textarea) textarea.value = '';
        replyForm.hidden = true;
      });
    });

    list.querySelectorAll('[data-reply-form]').forEach(replyForm => {
      replyForm.addEventListener('submit', async event => {
        event.preventDefault();
        const commentId = Number(replyForm.dataset.replyForm || 0);
        const noteContainer = replyForm.closest('.workspace-note');
        const noteId = Number(noteContainer?.querySelector('[data-open-note]')?.dataset.openNote || 0);
        const textarea = replyForm.querySelector('textarea');
        const content = String(textarea?.value || '').trim();
        if (!noteId || !commentId || !content) return;

        try {
          await apiRequest('note_comment_add', {
            method: 'POST',
            body: JSON.stringify({ note_id: noteId, parent_comment_id: commentId, content }),
          });
          if (textarea) textarea.value = '';
          replyForm.hidden = true;
          await loadNoteDetails(noteId);
          await loadNotes();
          renderNotes();
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-delete-comment]').forEach(button => {
      button.addEventListener('click', async () => {
        const noteContainer = button.closest('.workspace-note');
        const noteId = Number(noteContainer?.querySelector('[data-open-note]')?.dataset.openNote || 0);
        const commentId = Number(button.dataset.deleteComment || 0);
        if (!commentId || !noteId) return;

        try {
          await apiRequest('note_comment_delete', {
            method: 'POST',
            body: JSON.stringify({ comment_id: commentId }),
          });
          await loadNoteDetails(noteId);
          await loadNotes();
          renderNotes();
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-share-form]').forEach(shareForm => {
      shareForm.addEventListener('submit', async event => {
        event.preventDefault();
        const noteId = Number(shareForm.dataset.shareForm || 0);
        const input = shareForm.querySelector('input[type="email"]');
        const email = String(input?.value || '').trim();
        if (!noteId || !email) return;

        try {
          await apiRequest('note_share_invite', {
            method: 'POST',
            body: JSON.stringify({ note_id: noteId, email }),
          });
          if (input) input.value = '';
          await loadNoteDetails(noteId);
          renderNotes();
          showWorkspaceToast('Invitacion registrada.', 'success');
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });

    list.querySelectorAll('[data-revoke-share]').forEach(button => {
      button.addEventListener('click', async () => {
        const noteContainer = button.closest('.workspace-note');
        const noteId = Number(noteContainer?.querySelector('[data-open-note]')?.dataset.openNote || 0);
        const shareId = Number(button.dataset.revokeShare || 0);
        if (!noteId || !shareId) return;

        try {
          await apiRequest('note_share_revoke', {
            method: 'POST',
            body: JSON.stringify({ share_id: shareId }),
          });
          await loadNoteDetails(noteId);
          renderNotes();
          showWorkspaceToast('Invitacion revocada.', 'success');
        } catch (error) {
          showWorkspaceToast(error.message, 'danger');
        }
      });
    });
  };

  const loadNotes = async () => {
    const data = await apiRequest('notes_list');
    state.notes = Array.isArray(data.notes) ? data.notes : [];

    if (state.activeNoteId !== null) {
      const activeExists = state.notes.some(note => Number(note.id) === Number(state.activeNoteId));
      if (!activeExists) {
        state.activeNoteId = null;
      }
    }
  };

  const loadNoteDetails = async noteId => {
    const currentNote = state.notes.find(item => Number(item.id) === Number(noteId));
    if (!currentNote) return;

    const commentsPayload = await apiRequest('note_comments_list', {
      method: 'POST',
      body: JSON.stringify({ note_id: noteId }),
    });
    state.commentsByNote.set(Number(noteId), commentsPayload.comments || []);

    if (Number(currentNote.is_owner || 0) === 1) {
      const sharesPayload = await apiRequest('note_shares_list', {
        method: 'POST',
        body: JSON.stringify({ note_id: noteId }),
      });
      state.sharesByNote.set(Number(noteId), sharesPayload.shares || []);
    }
  };

  form?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('note_save', {
        method: 'POST',
        body: JSON.stringify({
          title: document.getElementById('noteTitle').value,
          content: document.getElementById('noteContent').value,
          color: document.getElementById('noteColor').value,
          is_pinned: document.getElementById('notePinned').checked,
        })
      });
      form.reset();
      showWorkspaceToast('Nota guardada.', 'success');
      await loadNotes();
      renderNotes();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  searchInput?.addEventListener('input', () => {
    state.search = String(searchInput.value || '').trim().toLowerCase();
    renderNotes();
  });

  colorFilter?.addEventListener('change', () => {
    state.color = colorFilter.value || 'all';
    renderNotes();
  });

  scopeFilter?.addEventListener('change', () => {
    state.scope = scopeFilter.value || 'all';
    renderNotes();
  });

  try {
    await loadNotes();
    renderNotes();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}

async function initTasksPage() {
  const form = document.getElementById('taskForm');
  const list = document.getElementById('tasksList');

  const load = async () => {
    const data = await apiRequest('tasks_list');
    renderList(list, data.tasks, task => `
      <div class="workspace-list-item">
        <div class="workspace-actions justify-content-between">
          <strong>${task.title}</strong>
          <div class="workspace-actions">
            <button class="btn btn-sm btn-outline-success" data-toggle-task="${task.id}"><i class="fas fa-check"></i></button>
            <button class="btn btn-sm btn-outline-danger" data-delete-task="${task.id}"><i class="fas fa-trash"></i></button>
          </div>
        </div>
        <p>${task.task_date} · ${task.priority} · ${task.status}${task.description ? ` · ${task.description}` : ''}</p>
      </div>`, 'No hay tareas cargadas.');

    list.querySelectorAll('[data-toggle-task]').forEach(button => {
      button.addEventListener('click', async () => {
        await apiRequest('task_toggle', { method: 'POST', body: JSON.stringify({ id: button.dataset.toggleTask }) });
        await load();
      });
    });
    list.querySelectorAll('[data-delete-task]').forEach(button => {
      button.addEventListener('click', async () => {
        await apiRequest('task_delete', { method: 'POST', body: JSON.stringify({ id: button.dataset.deleteTask }) });
        await load();
      });
    });
  };

  form?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('task_save', {
        method: 'POST',
        body: JSON.stringify({
          title: document.getElementById('taskTitle').value,
          task_date: document.getElementById('taskDate').value,
          description: document.getElementById('taskDescription').value,
          status: document.getElementById('taskStatus').value,
          priority: document.getElementById('taskPriority').value,
          due_time: document.getElementById('taskDueTime').value,
        })
      });
      form.reset();
      document.getElementById('taskDate').valueAsDate = new Date();
      await load();
      showWorkspaceToast('Tarea guardada.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  document.getElementById('taskDate').valueAsDate = new Date();
  try {
    await load();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}

async function initGoalsPage() {
  const form = document.getElementById('goalForm');
  const list = document.getElementById('goalsList');

  const load = async () => {
    const data = await apiRequest('goals_list');
    renderList(list, data.goals, goal => `
      <div class="workspace-list-item">
        <div class="workspace-actions justify-content-between">
          <strong>${goal.title}</strong>
          <button class="btn btn-sm btn-outline-danger" data-delete-goal="${goal.id}"><i class="fas fa-trash"></i></button>
        </div>
        <p>${goal.progress_percent}% · ${goal.status}${goal.target_date ? ` · ${goal.target_date}` : ''}</p>
      </div>`, 'No hay metas registradas.');

    list.querySelectorAll('[data-delete-goal]').forEach(button => {
      button.addEventListener('click', async () => {
        await apiRequest('goal_delete', { method: 'POST', body: JSON.stringify({ id: button.dataset.deleteGoal }) });
        await load();
      });
    });
  };

  form?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('goal_save', {
        method: 'POST',
        body: JSON.stringify({
          title: document.getElementById('goalTitle').value,
          description: document.getElementById('goalDescription').value,
          target_date: document.getElementById('goalTargetDate').value,
          progress_percent: document.getElementById('goalProgress').value,
          status: document.getElementById('goalStatus').value,
        })
      });
      form.reset();
      await load();
      showWorkspaceToast('Meta guardada.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  try {
    await load();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}

async function initCalendarPage() {
  const calGrid = document.getElementById('calGrid');
  const calMiniGrid = document.getElementById('calMiniGrid');
  const calMonthTitle = document.getElementById('calMonthTitle');
  const calMiniMonthLabel = document.getElementById('calMiniMonthLabel');
  const calAgendaDate = document.getElementById('calAgendaDate');
  const calAgendaList = document.getElementById('calAgendaList');
  const calPrevBtn = document.getElementById('calPrevBtn');
  const calNextBtn = document.getElementById('calNextBtn');
  const calTodayBtn = document.getElementById('calTodayBtn');
  const calMiniPrev = document.getElementById('calMiniPrev');
  const calMiniNext = document.getElementById('calMiniNext');
  const calNewEventBtn = document.getElementById('calNewEventBtn');
  const calSyncNowBtn = document.getElementById('calSyncNowBtn');
  const calSyncStatus = document.getElementById('calSyncStatus');

  const calFilterManual = document.getElementById('calFilterManual');
  const calFilterTask = document.getElementById('calFilterTask');
  const calFilterMeeting = document.getElementById('calFilterMeeting');
  const calFilterSession = document.getElementById('calFilterSession');
  const calFilterTeams = document.getElementById('calFilterTeams');

  const calDayPopup = document.getElementById('calDayPopup');
  const calPopupDate = document.getElementById('calPopupDate');
  const calPopupEvents = document.getElementById('calPopupEvents');
  const calPopupClose = document.getElementById('calPopupClose');
  const calQuickForm = document.getElementById('calQuickForm');
  const calQuickTitle = document.getElementById('calQuickTitle');
  const calQuickType = document.getElementById('calQuickType');
  const calQuickDate = document.getElementById('calQuickDate');
  const calQuickStart = document.getElementById('calQuickStart');
  const calQuickEnd = document.getElementById('calQuickEnd');
  const calQuickUrl = document.getElementById('calQuickUrl');
  const calQuickCancel = document.getElementById('calQuickCancel');
  const calTeamsSyncRow = document.getElementById('calTeamsSyncRow');
  const calTeamsSyncCheck = document.getElementById('calTeamsSyncCheck');

  const rotateKeyButton = document.getElementById('powerAutomateRotateKey');
  const externalEmailInput = document.getElementById('powerAutomateExternalEmail');
  const webhookUrlInput = document.getElementById('powerAutomateWebhookUrl');
  const headerNameInput = document.getElementById('powerAutomateHeaderName');
  const tokenInput = document.getElementById('powerAutomateToken');
  const statusBox = document.getElementById('powerAutomateStatus');
  const teamsTodaySummary = document.getElementById('teamsTodaySummary');
  const outboundUrlInput = document.getElementById('powerAutomateOutboundUrl');
  const saveOutboundUrlBtn = document.getElementById('saveOutboundUrlBtn');

  if (!calGrid || !calMonthTitle || !calAgendaList) {
    showWorkspaceToast('La vista de calendario no cargo correctamente.', 'danger');
    return;
  }

  const today = new Date();
  const MONTHS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const WEEK_DAYS = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo'];

  let currentYear = today.getFullYear();
  let currentMonth = today.getMonth();
  let selectedDate = formatDate(today);
  let activeDayEl = null;
  let allEvents = [];
  let isPAConfigured = false;

  const setSyncStatus = (message, variant = 'neutral', isLoading = false) => {
    if (!calSyncStatus) return;
    calSyncStatus.textContent = message;
    calSyncStatus.classList.remove('sync-ok', 'sync-error', 'sync-loading');
    if (variant === 'success') calSyncStatus.classList.add('sync-ok');
    if (variant === 'error') calSyncStatus.classList.add('sync-error');
    if (isLoading) calSyncStatus.classList.add('sync-loading');

    if (calSyncNowBtn) {
      calSyncNowBtn.disabled = isLoading;
      const icon = calSyncNowBtn.querySelector('i');
      if (icon) {
        icon.className = isLoading ? 'fas fa-rotate fa-spin' : 'fas fa-rotate';
      }
    }
  };

  const TYPE_META = {
    task: { label: 'Tarea', icon: 'fa-check-square', cls: 'task' },
    meeting: { label: 'Reunion', icon: 'fa-video', cls: 'meeting' },
    session: { label: 'Sesion', icon: 'fa-clock', cls: 'session' },
    teams: { label: 'Teams', icon: 'fa-users', cls: 'teams' },
    power_automate_teams: { label: 'Teams', icon: 'fa-users', cls: 'teams' },
    manual: { label: 'Evento', icon: 'fa-calendar-day', cls: 'manual' },
  };

  function formatDate(dateObj) {
    return `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
  }

  function escHtml(value) {
    return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function sourceTypeForFilter(rawType) {
    const sourceType = rawType === 'power_automate_teams' ? 'teams' : (rawType || 'manual');
    if (!['manual', 'task', 'meeting', 'session', 'teams'].includes(sourceType)) {
      return 'manual';
    }
    return sourceType;
  }

  function buildFiltersState() {
    return {
      manual: Boolean(calFilterManual?.checked),
      task: Boolean(calFilterTask?.checked),
      meeting: Boolean(calFilterMeeting?.checked),
      session: Boolean(calFilterSession?.checked),
      teams: Boolean(calFilterTeams?.checked),
    };
  }

  function getFilteredEvents() {
    const filters = buildFiltersState();
    return (allEvents || []).filter(eventItem => filters[sourceTypeForFilter(eventItem.source_type)]);
  }

  function buildEventsMap(events) {
    const map = {};
    events.forEach(eventItem => {
      const day = (eventItem.start_at || '').slice(0, 10);
      if (!day) return;
      if (!map[day]) map[day] = [];
      map[day].push(eventItem);
    });

    Object.keys(map).forEach(day => {
      map[day].sort((a, b) => String(a.start_at || '').localeCompare(String(b.start_at || '')));
    });

    return map;
  }

  function monthMatrix(year, month) {
    const first = new Date(year, month, 1);
    const firstWeekdayMonday = (first.getDay() + 6) % 7;
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const prevMonthDays = new Date(year, month, 0).getDate();
    const cells = [];

    for (let i = firstWeekdayMonday - 1; i >= 0; i--) {
      const day = prevMonthDays - i;
      const prevMonth = month === 0 ? 11 : month - 1;
      const prevYear = month === 0 ? year - 1 : year;
      cells.push({
        date: `${prevYear}-${String(prevMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
        day,
        otherMonth: true,
      });
    }

    for (let day = 1; day <= daysInMonth; day++) {
      cells.push({
        date: `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
        day,
        otherMonth: false,
      });
    }

    const total = Math.ceil(cells.length / 7) * 7;
    const nextMonth = month === 11 ? 0 : month + 1;
    const nextYear = month === 11 ? year + 1 : year;
    for (let day = 1; cells.length < total; day++) {
      cells.push({
        date: `${nextYear}-${String(nextMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
        day,
        otherMonth: true,
      });
    }

    return cells;
  }

  function renderMainCalendar() {
    const filteredEvents = getFilteredEvents();
    const eventMap = buildEventsMap(filteredEvents);
    const cells = monthMatrix(currentYear, currentMonth);
    const todayString = formatDate(today);

    calMonthTitle.textContent = `${MONTHS[currentMonth]} ${currentYear}`;

    calGrid.innerHTML = cells.map(cell => {
      const dayEvents = eventMap[cell.date] || [];
      const chips = dayEvents.slice(0, 3).map(eventItem => {
        const mappedType = sourceTypeForFilter(eventItem.source_type);
        const startTime = (eventItem.start_at || '').slice(11, 16);
        return `<div class="cal-event-chip ${mappedType}"><span class="time">${escHtml(startTime)}</span><span class="title">${escHtml(eventItem.title)}</span></div>`;
      }).join('');
      const more = dayEvents.length > 3 ? `<div class="cal-chip-more">+${dayEvents.length - 3} mas</div>` : '';
      const classes = ['cal-day'];
      if (cell.otherMonth) classes.push('other-month');
      if (cell.date === todayString) classes.push('today');
      if (cell.date === selectedDate) classes.push('selected');

      return `<div class="${classes.join(' ')}" data-date="${cell.date}" role="gridcell" tabindex="0" aria-label="${cell.date}">
        <span class="cal-day-num">${cell.day}</span>
        <div class="cal-day-dots">${chips}${more}</div>
      </div>`;
    }).join('');

    calGrid.querySelectorAll('.cal-day').forEach(dayCell => {
      dayCell.addEventListener('click', () => onDaySelected(dayCell.dataset.date, dayCell, eventMap));
      dayCell.addEventListener('keydown', event => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          onDaySelected(dayCell.dataset.date, dayCell, eventMap);
        }
      });
    });
  }

  function renderMiniCalendar() {
    if (!calMiniGrid || !calMiniMonthLabel) return;
    const cells = monthMatrix(currentYear, currentMonth);
    const todayString = formatDate(today);

    calMiniMonthLabel.textContent = `${MONTHS[currentMonth]} ${currentYear}`;
    calMiniGrid.innerHTML = cells.map(cell => {
      const classes = ['cal-mini-day'];
      if (cell.otherMonth) classes.push('other-month');
      if (cell.date === todayString) classes.push('today');
      if (cell.date === selectedDate) classes.push('selected');
      return `<button type="button" class="${classes.join(' ')}" data-date="${cell.date}">${cell.day}</button>`;
    }).join('');

    calMiniGrid.querySelectorAll('.cal-mini-day').forEach(dayButton => {
      dayButton.addEventListener('click', () => {
        const clickedDate = String(dayButton.dataset.date || '');
        if (!clickedDate) return;
        const [year, month] = clickedDate.split('-').map(Number);
        currentYear = year;
        currentMonth = month - 1;
        selectedDate = clickedDate;
        renderMainCalendar();
        renderMiniCalendar();
        renderAgenda();
      });
    });
  }

  function renderAgenda() {
    const filteredEvents = getFilteredEvents();
    const selectedEvents = filteredEvents
      .filter(eventItem => (eventItem.start_at || '').slice(0, 10) === selectedDate)
      .sort((a, b) => String(a.start_at || '').localeCompare(String(b.start_at || '')));

    const [year, month, day] = selectedDate.split('-').map(Number);
    const dateObj = new Date(year, month - 1, day);
    if (calAgendaDate) {
      calAgendaDate.textContent = `${WEEK_DAYS[(dateObj.getDay() + 6) % 7]} ${day} ${MONTHS[month - 1]}`;
    }

    if (selectedEvents.length === 0) {
      calAgendaList.innerHTML = '<div class="workspace-empty">No hay nada programado para este dia. Usa Nuevo para crear tarea, reunion o sesion.</div>';
      return;
    }

    calAgendaList.innerHTML = selectedEvents.map(eventItem => {
      const mappedType = sourceTypeForFilter(eventItem.source_type);
      const meta = TYPE_META[mappedType] || TYPE_META.manual;
      const startTime = (eventItem.start_at || '').slice(11, 16);
      const endTime = (eventItem.end_at || '').slice(11, 16);
      return `<div class="cal-agenda-item ${meta.cls}">
        <span class="cal-agenda-time">${escHtml(startTime)}${endTime ? ' - ' + escHtml(endTime) : ''}</span>
        <div class="cal-agenda-content">
          <strong>${escHtml(eventItem.title)}</strong>
          <small>${escHtml(meta.label)}${eventItem.location ? ' · ' + escHtml(eventItem.location) : ''}</small>
        </div>
      </div>`;
    }).join('');
  }

  function onDaySelected(date, dayElement, eventMap) {
    selectedDate = date;
    if (activeDayEl) activeDayEl.classList.remove('selected');
    activeDayEl = dayElement;
    activeDayEl.classList.add('selected');
    renderMiniCalendar();
    renderAgenda();
    openDayPopup(date, eventMap[date] || [], dayElement);
  }

  function closeQuickForm() {
    if (!calQuickForm) return;
    calQuickForm.hidden = true;
    if (calQuickTitle) calQuickTitle.value = '';
    if (calQuickUrl) calQuickUrl.value = '';
    if (calDayPopup) {
      calDayPopup.querySelectorAll('.cal-type-btn').forEach(button => button.classList.remove('active'));
    }
  }

  function closeDayPopup() {
    if (calDayPopup) calDayPopup.hidden = true;
    closeQuickForm();
  }

  function positionPopup(triggerElement) {
    if (!calDayPopup || !triggerElement) return;
    const triggerRect = triggerElement.getBoundingClientRect();
    const popupWidth = 310;
    const popupHeight = calDayPopup.scrollHeight || 380;

    let left = triggerRect.left;
    let top = triggerRect.bottom + 8;

    if (left + popupWidth > window.innerWidth - 12) left = window.innerWidth - popupWidth - 12;
    if (left < 12) left = 12;
    if (top + popupHeight > window.innerHeight - 12) top = triggerRect.top - popupHeight - 8;
    if (top < 12) top = 12;

    calDayPopup.style.left = `${left}px`;
    calDayPopup.style.top = `${top}px`;
  }

  function openDayPopup(date, dayEvents, triggerElement) {
    if (!calDayPopup || !calPopupEvents || !calPopupDate) return;

    const [year, month, day] = date.split('-').map(Number);
    calPopupDate.textContent = `${day} de ${MONTHS[month - 1]} ${year}`;
    if (calQuickDate) calQuickDate.value = date;

    closeQuickForm();

    if (!dayEvents.length) {
      calPopupEvents.innerHTML = '<p class="cal-popup-empty"><i class="fas fa-calendar-day"></i> Sin eventos este dia</p>';
    } else {
      calPopupEvents.innerHTML = dayEvents.map(eventItem => {
        const mappedType = sourceTypeForFilter(eventItem.source_type);
        const meta = TYPE_META[mappedType] || TYPE_META.manual;
        const startTime = (eventItem.start_at || '').slice(11, 16);
        const endTime = (eventItem.end_at || '').slice(11, 16);
        const canDelete = eventItem.source_type !== 'power_automate_teams';
        return `<div class="cal-popup-event-item">
          <span class="cal-event-icon ${meta.cls}"><i class="fas ${meta.icon}"></i></span>
          <div class="cal-popup-event-meta">
            <strong>${escHtml(eventItem.title)}</strong>
            <small>${escHtml(startTime)}${endTime ? ' -> ' + escHtml(endTime) : ''}</small>
          </div>
          ${canDelete ? `<button class="cal-popup-event-del" data-del-id="${eventItem.id}" title="Eliminar"><i class="fas fa-trash"></i></button>` : ''}
        </div>`;
      }).join('');

      calPopupEvents.querySelectorAll('[data-del-id]').forEach(deleteButton => {
        deleteButton.addEventListener('click', async () => {
          try {
            await apiRequest('calendar_delete', { method: 'POST', body: JSON.stringify({ id: deleteButton.dataset.delId }) });
            showWorkspaceToast('Evento eliminado.', 'success');
            closeDayPopup();
            await load();
          } catch (error) {
            showWorkspaceToast(error.message, 'danger');
          }
        });
      });
    }

    if (calTeamsSyncRow) calTeamsSyncRow.style.display = isPAConfigured ? 'flex' : 'none';
    calDayPopup.hidden = false;
    positionPopup(triggerElement);
  }

  if (calDayPopup) {
    calDayPopup.querySelectorAll('.cal-type-btn').forEach(typeButton => {
      typeButton.addEventListener('click', () => {
        if (!calDayPopup) return;
        calDayPopup.querySelectorAll('.cal-type-btn').forEach(button => button.classList.remove('active'));
        typeButton.classList.add('active');
        if (calQuickType) calQuickType.value = String(typeButton.dataset.type || 'manual');
        if (calQuickForm) calQuickForm.hidden = false;
        calQuickTitle?.focus();
      });
    });
  }

  calQuickCancel?.addEventListener('click', closeQuickForm);
  calPopupClose?.addEventListener('click', closeDayPopup);

  calQuickForm?.addEventListener('submit', async event => {
    event.preventDefault();

    const day = calQuickDate?.value || selectedDate;
    const start = calQuickStart?.value || '09:00';
    const end = calQuickEnd?.value || '10:00';
    const eventType = calQuickType?.value || 'manual';

    try {
      const saveResult = await apiRequest('calendar_save', {
        method: 'POST',
        body: JSON.stringify({
          title: calQuickTitle?.value || 'Nuevo evento',
          description: '',
          start_at: `${day} ${start}:00`,
          end_at: `${day} ${end}:00`,
          location: '',
          meeting_url: calQuickUrl?.value || '',
          source_type: eventType,
        }),
      });

      const outbound = saveResult?.outbound_sync;
      if (outbound && outbound.attempted) {
        if (outbound.ok) {
          showWorkspaceToast('Guardado y enviado a Outlook/Teams.', 'success');
        } else {
          showWorkspaceToast(outbound.message || 'Guardado local. Power Automate no respondio.', 'warning');
        }
      } else {
        showWorkspaceToast('Evento guardado.', 'success');
      }

      closeDayPopup();
      await load();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') closeDayPopup();
  });

  document.addEventListener('click', event => {
    if (!calDayPopup || calDayPopup.hidden) return;
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (!calDayPopup.contains(target) && !target.closest('.cal-day') && !target.closest('#calNewEventBtn')) {
      closeDayPopup();
    }
  });

  calPrevBtn?.addEventListener('click', () => {
    currentMonth -= 1;
    if (currentMonth < 0) {
      currentMonth = 11;
      currentYear -= 1;
    }
    renderMainCalendar();
    renderMiniCalendar();
  });

  calNextBtn?.addEventListener('click', () => {
    currentMonth += 1;
    if (currentMonth > 11) {
      currentMonth = 0;
      currentYear += 1;
    }
    renderMainCalendar();
    renderMiniCalendar();
  });

  calTodayBtn?.addEventListener('click', () => {
    currentYear = today.getFullYear();
    currentMonth = today.getMonth();
    selectedDate = formatDate(today);
    renderMainCalendar();
    renderMiniCalendar();
    renderAgenda();
  });

  calMiniPrev?.addEventListener('click', () => calPrevBtn?.click());
  calMiniNext?.addEventListener('click', () => calNextBtn?.click());

  [calFilterManual, calFilterTask, calFilterMeeting, calFilterSession, calFilterTeams].forEach(filter => {
    filter?.addEventListener('change', () => {
      renderMainCalendar();
      renderMiniCalendar();
      renderAgenda();
    });
  });

  calNewEventBtn?.addEventListener('click', () => {
    const selectedCell = calGrid.querySelector(`.cal-day[data-date="${selectedDate}"]`);
    openDayPopup(selectedDate, getFilteredEvents().filter(eventItem => (eventItem.start_at || '').slice(0, 10) === selectedDate), selectedCell || calGrid);
  });

  calSyncNowBtn?.addEventListener('click', async () => {
    try {
      setSyncStatus('Solicitando sincronizacion con Power Automate...', 'neutral', true);
      const result = await apiRequest('calendar_sync_request', {
        method: 'POST',
        body: JSON.stringify({}),
      });

      setSyncStatus(result.message || 'Solicitud enviada. Esperando eventos...', 'success', false);
      showWorkspaceToast(result.message || 'Solicitud de sincronizacion enviada.', 'success');

      setTimeout(async () => {
        try {
          await load();
          setSyncStatus('Sincronizacion completada. Datos actualizados.', 'success', false);
        } catch (error) {
          setSyncStatus('No se pudo refrescar despues de sincronizar.', 'error', false);
        }
      }, 1500);
    } catch (error) {
      setSyncStatus(error.message || 'Error al sincronizar con Power Automate.', 'error', false);
      showWorkspaceToast(error.message || 'Error al sincronizar.', 'danger');
    }
  });

  const renderPowerAutomateConfig = config => {
    if (!config) {
      return;
    }

    if (externalEmailInput) externalEmailInput.value = config.external_account_email || '';
    if (webhookUrlInput) webhookUrlInput.value = config.webhook_url || '';
    if (headerNameInput) headerNameInput.value = config.header_name || 'X-Power-Automate-Key';
    if (outboundUrlInput) outboundUrlInput.value = config.outbound_webhook_url || '';
    if (statusBox) {
      statusBox.textContent = config.configured
        ? `Clave activa. Estado: ${config.sync_status || 'configured'}${config.last_synced_at ? ` · ultimo sync ${config.last_synced_at}` : ''}${config.token_preview ? ` · ${config.token_preview}` : ''}`
        : 'La integracion aun no tiene clave activa.';
    }
  };

  const load = async () => {
    const data = await apiRequest('calendar_list');
    allEvents = data.events || [];
    renderMainCalendar();
    renderMiniCalendar();
    renderAgenda();

    renderPowerAutomateConfig(data.power_automate);

    if (teamsTodaySummary) {
      teamsTodaySummary.textContent = data.teams_today && data.teams_today.sessions > 0
        ? `Hoy se acumularon ${data.teams_today.minutes} min en ${data.teams_today.sessions} sesiones de Teams sincronizadas.`
        : 'Sin datos sincronizados hoy.';
    }

    if (data.power_automate?.configured) {
      const syncState = data.power_automate?.sync_status || 'configured';
      setSyncStatus(`Estado Power Automate: ${syncState}${data.power_automate?.last_synced_at ? ` · ultimo sync ${data.power_automate.last_synced_at}` : ''}`, 'success', false);
    } else {
      setSyncStatus('Sincronizacion inactiva. Configura Power Automate para habilitarla.', 'error', false);
    }
  };

  rotateKeyButton?.addEventListener('click', async () => {
    try {
      const data = await apiRequest('power_automate_config_rotate', {
        method: 'POST',
        body: JSON.stringify({
          external_account_email: externalEmailInput?.value || '',
        })
      });

      if (tokenInput) tokenInput.value = data.config.token || '';
      if (webhookUrlInput) webhookUrlInput.value = data.config.webhook_url || '';
      if (headerNameInput) headerNameInput.value = data.config.header_name || 'X-Power-Automate-Key';
      if (statusBox) {
        statusBox.textContent = 'Nueva clave generada. Copiala ahora y guardala como encabezado fijo en Power Automate.';
      }
      showWorkspaceToast('Clave generada. Copiala antes de salir.', 'success');
      await load();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  saveOutboundUrlBtn?.addEventListener('click', async () => {
    const outboundUrl = outboundUrlInput?.value.trim() || '';
    try {
      await apiRequest('power_automate_set_outbound', {
        method: 'POST',
        body: JSON.stringify({ outbound_webhook_url: outboundUrl }),
      });
      showWorkspaceToast('URL de salida guardada.', 'success');
      await load();
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  try {
    await load();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}

function initSummaryModuleModal() {
  const backdrop = document.getElementById('summaryModuleBackdrop');
  const closeBtn = document.getElementById('summaryModuleClose');
  const kicker = document.getElementById('summaryModuleKicker');
  const title = document.getElementById('summaryModuleTitle');
  const description = document.getElementById('summaryModuleDescription');
  const points = document.getElementById('summaryModulePoints');
  const link = document.getElementById('summaryModuleLink');

  if (!backdrop || !closeBtn || !kicker || !title || !description || !points || !link) {
    return;
  }

  const closeModal = () => {
    backdrop.hidden = true;
    document.body.classList.remove('ws-modal-open');
  };

  const openModal = moduleKey => {
    const module = summaryModuleMeta[moduleKey];
    if (!module) {
      return;
    }

    kicker.textContent = module.kicker;
    title.textContent = module.title;
    description.textContent = module.description;
    points.innerHTML = module.points.map(point => `<p><i class="fas fa-check-circle"></i>${point}</p>`).join('');
    link.href = module.href;
    link.textContent = 'Ir directamente al modulo';

    backdrop.hidden = false;
    document.body.classList.add('ws-modal-open');
  };

  document.addEventListener('click', event => {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }

    const trigger = target.closest('.summary-module-trigger, .summary-stat-card[data-module]');
    if (!trigger) {
      return;
    }

    const moduleKey = String(trigger.getAttribute('data-module') || '').trim();
    if (!moduleKey || !summaryModuleMeta[moduleKey]) {
      return;
    }

    event.preventDefault();
    openModal(moduleKey);
  });

  closeBtn.addEventListener('click', closeModal);
  backdrop.addEventListener('click', event => {
    if (event.target === backdrop) {
      closeModal();
    }
  });

  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && !backdrop.hidden) {
      closeModal();
    }
  });
}