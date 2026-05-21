<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Calendario', 'calendar', $user);
?>
<div class="workspace-calendar-banner">
  <strong>Base de integracion lista para Teams y Microsoft Calendar</strong>
  <p class="mb-0 workspace-muted">La base de datos ya contempla fuentes, eventos y tokens de integracion. La sincronizacion automatica requiere registrar una aplicacion en Microsoft Graph y sus credenciales.</p>
</div>

<div class="workspace-grid">
  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Agenda</span>
        <h3>Nuevo evento</h3>
      </div>
    </div>
    <form id="calendarForm" class="workspace-form">
      <div>
        <label for="calendarTitle" class="form-label">Titulo</label>
        <input type="text" class="form-control" id="calendarTitle" required>
      </div>
      <div>
        <label for="calendarDescription" class="form-label">Descripcion</label>
        <textarea class="form-control" id="calendarDescription" rows="4"></textarea>
      </div>
      <div>
        <label for="calendarStartAt" class="form-label">Inicio</label>
        <input type="datetime-local" class="form-control" id="calendarStartAt" required>
      </div>
      <div>
        <label for="calendarEndAt" class="form-label">Fin</label>
        <input type="datetime-local" class="form-control" id="calendarEndAt" required>
      </div>
      <div>
        <label for="calendarLocation" class="form-label">Lugar</label>
        <input type="text" class="form-control" id="calendarLocation">
      </div>
      <div>
        <label for="calendarMeetingUrl" class="form-label">URL de reunion</label>
        <input type="url" class="form-control" id="calendarMeetingUrl" placeholder="https://teams.microsoft.com/...">
      </div>
      <button type="submit" class="btn btn-primary">Guardar evento</button>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Eventos</span>
        <h3>Calendario del sistema</h3>
      </div>
    </div>
    <div class="workspace-list" id="calendarList"></div>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Integraciones</span>
        <h3>Fuentes de sincronizacion</h3>
      </div>
    </div>
    <div class="workspace-list" id="calendarSources"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>