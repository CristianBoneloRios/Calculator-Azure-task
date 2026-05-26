'use strict';

const workspaceToastContainer = document.getElementById('workspaceToastContainer');
const workspacePage = document.body.dataset.page;

document.addEventListener('DOMContentLoaded', () => {
  if (workspacePage === 'index') {
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
});

async function apiRequest(action, options = {}) {
  const response = await fetch(`../api/dashboard.php?action=${encodeURIComponent(action)}`, {
    method: options.method || 'GET',
    credentials: 'same-origin',
    headers: options.isFormData ? {} : { 'Content-Type': 'application/json' },
    body: options.body || null,
  });

  const data = await response.json();
  if (!response.ok || data.ok === false) {
    throw new Error(data.message || 'Error inesperado.');
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

  try {
    const data = await apiRequest('summary');
    summaryCards.innerHTML = `
      <article class="workspace-stat"><span class="workspace-tag">Notas</span><strong>${data.summary.notes}</strong><p class="workspace-muted">Ideas activas y referencias</p></article>
      <article class="workspace-stat"><span class="workspace-tag">Hoy</span><strong>${data.summary.tasks_today}</strong><p class="workspace-muted">Tareas programadas para hoy</p></article>
      <article class="workspace-stat"><span class="workspace-tag">Metas</span><strong>${data.summary.goals_active}</strong><p class="workspace-muted">Objetivos en curso</p></article>
      <article class="workspace-stat"><span class="workspace-tag">Agenda</span><strong>${data.summary.events_upcoming}</strong><p class="workspace-muted">Eventos por atender</p></article>
      <article class="workspace-stat"><span class="workspace-tag">Teams hoy</span><strong>${data.summary.teams_hours_today_label}</strong><p class="workspace-muted">${data.summary.teams_sessions_today} sesiones sincronizadas desde Power Automate</p></article>`;

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

async function initProfilePage() {
  const profileForm = document.getElementById('profileForm');
  const publicProfileForm = document.getElementById('publicProfileForm');
  const photoForm = document.getElementById('photoForm');
  const preview = document.getElementById('profilePhotoPreview');

  try {
    const data = await apiRequest('profile_get');
    document.getElementById('profileFullName').value = data.user.full_name || '';
    document.getElementById('profileEmail').value = data.user.email || '';
    document.getElementById('publicDisplayName').value = data.public_profile?.display_name || '';
    document.getElementById('publicRoleTitle').value = data.public_profile?.role_title || '';
    document.getElementById('publicCompanyName').value = data.public_profile?.company_name || '';
    document.getElementById('publicBio').value = data.public_profile?.bio || '';
    if (preview && data.user.profile_photo_path) {
      preview.src = `../${data.user.profile_photo_path}`;
    } else if (preview && data.public_profile?.photo_url) {
      preview.src = data.public_profile.photo_url;
    }
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }

  profileForm?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('profile_update', {
        method: 'POST',
        body: JSON.stringify({
          full_name: document.getElementById('profileFullName').value,
          email: document.getElementById('profileEmail').value,
          password: document.getElementById('profilePassword').value,
        })
      });
      document.getElementById('profilePassword').value = '';
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
        })
      });
      showWorkspaceToast('Perfil publico actualizado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

  photoForm?.addEventListener('submit', async event => {
    event.preventDefault();
    const input = document.getElementById('profilePhotoInput');
    if (!input.files.length) {
      showWorkspaceToast('Selecciona una imagen primero.', 'warning');
      return;
    }

    const formData = new FormData();
    formData.append('photo', input.files[0]);
    formData.append('make_public_profile', document.getElementById('makePublicProfilePhoto').checked ? '1' : '0');

    try {
      const response = await fetch('../api/dashboard.php?action=profile_photo_upload', {
        method: 'POST',
        credentials: 'same-origin',
        body: formData,
      });
      const data = await response.json();
      if (!response.ok || data.ok === false) {
        throw new Error(data.message || 'No se pudo subir la foto.');
      }
      preview.src = data.photo_url;
      input.value = '';
      showWorkspaceToast('Foto de perfil actualizada.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });
}

async function initNotesPage() {
  const form = document.getElementById('noteForm');
  const list = document.getElementById('notesList');

  const load = async () => {
    const data = await apiRequest('notes_list');
    renderList(list, data.notes, note => `
      <article class="workspace-note" data-color="${note.color}">
        <div class="workspace-actions justify-content-between">
          <strong>${note.title}</strong>
          <button class="btn btn-sm btn-outline-danger" data-delete-note="${note.id}"><i class="fas fa-trash"></i></button>
        </div>
        <p>${note.content.replace(/\n/g, '<br>')}</p>
      </article>`, 'No hay notas importantes registradas.');

    list.querySelectorAll('[data-delete-note]').forEach(button => {
      button.addEventListener('click', async () => {
        await apiRequest('note_delete', { method: 'POST', body: JSON.stringify({ id: button.dataset.deleteNote }) });
        showWorkspaceToast('Nota eliminada.', 'success');
        await load();
      });
    });
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
  const form = document.getElementById('calendarForm');
  const list = document.getElementById('calendarList');
  const sources = document.getElementById('calendarSources');
  const rotateKeyButton = document.getElementById('powerAutomateRotateKey');
  const externalEmailInput = document.getElementById('powerAutomateExternalEmail');
  const webhookUrlInput = document.getElementById('powerAutomateWebhookUrl');
  const headerNameInput = document.getElementById('powerAutomateHeaderName');
  const tokenInput = document.getElementById('powerAutomateToken');
  const statusBox = document.getElementById('powerAutomateStatus');
  const teamsTodaySummary = document.getElementById('teamsTodaySummary');

  const renderPowerAutomateConfig = config => {
    if (!config) {
      return;
    }

    if (externalEmailInput) externalEmailInput.value = config.external_account_email || '';
    if (webhookUrlInput) webhookUrlInput.value = config.webhook_url || '';
    if (headerNameInput) headerNameInput.value = config.header_name || 'X-Power-Automate-Key';
    if (statusBox) {
      statusBox.textContent = config.configured
        ? `Clave activa. Estado: ${config.sync_status || 'configured'}${config.last_synced_at ? ` · ultimo sync ${config.last_synced_at}` : ''}${config.token_preview ? ` · ${config.token_preview}` : ''}`
        : 'La integracion aun no tiene clave activa.';
    }
  };

  const load = async () => {
    const data = await apiRequest('calendar_list');
    renderList(list, data.events, event => `
      <div class="workspace-list-item">
        <div class="workspace-actions justify-content-between">
          <strong>${event.title}</strong>
          <button class="btn btn-sm btn-outline-danger" data-delete-event="${event.id}"><i class="fas fa-trash"></i></button>
        </div>
        <p>${event.start_at} → ${event.end_at}${event.location ? ` · ${event.location}` : ''}${event.source_type ? ` · ${event.source_type}` : ''}</p>
      </div>`, 'No hay eventos programados.');
    renderList(sources, data.sources, source => `<div class="workspace-list-item"><strong>${source.provider}</strong><p>${source.sync_status}${source.external_account_email ? ` · ${source.external_account_email}` : ''}${source.last_synced_at ? ` · ultimo sync ${source.last_synced_at}` : ''}</p></div>`, 'Teams y Calendar aun no estan conectados.');
    renderPowerAutomateConfig(data.power_automate);
    if (teamsTodaySummary) {
      teamsTodaySummary.textContent = data.teams_today && data.teams_today.sessions > 0
        ? `Hoy se acumularon ${data.teams_today.minutes} min en ${data.teams_today.sessions} sesiones de Teams sincronizadas.`
        : 'Sin datos sincronizados hoy.';
    }

    list.querySelectorAll('[data-delete-event]').forEach(button => {
      button.addEventListener('click', async () => {
        await apiRequest('calendar_delete', { method: 'POST', body: JSON.stringify({ id: button.dataset.deleteEvent }) });
        await load();
      });
    });
  };

  form?.addEventListener('submit', async event => {
    event.preventDefault();
    try {
      await apiRequest('calendar_save', {
        method: 'POST',
        body: JSON.stringify({
          title: document.getElementById('calendarTitle').value,
          description: document.getElementById('calendarDescription').value,
          start_at: document.getElementById('calendarStartAt').value.replace('T', ' '),
          end_at: document.getElementById('calendarEndAt').value.replace('T', ' '),
          location: document.getElementById('calendarLocation').value,
          meeting_url: document.getElementById('calendarMeetingUrl').value,
          source_type: 'manual',
        })
      });
      form.reset();
      await load();
      showWorkspaceToast('Evento guardado.', 'success');
    } catch (error) {
      showWorkspaceToast(error.message, 'danger');
    }
  });

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

  try {
    await load();
  } catch (error) {
    showWorkspaceToast(error.message, 'danger');
  }
}