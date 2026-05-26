<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Calendario', 'calendar', $user);
?>
<div class="workspace-calendar-banner">
  <strong>Power Automate listo para consolidar sesiones de Teams sin Entra ID</strong>
  <p class="mb-0 workspace-muted">Genera una clave segura, configura un flujo diario en Microsoft 365 y envia solo el corte del dia a este backend para calcular el total de tiempo en sesiones.</p>
</div>

<div class="workspace-grid">
  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Power Automate</span>
        <h3>Webhook seguro para corte diario</h3>
      </div>
      <button type="button" class="btn btn-sm btn-primary" id="powerAutomateRotateKey">Generar clave</button>
    </div>

    <div class="workspace-form">
      <div>
        <label for="powerAutomateExternalEmail" class="form-label">Correo de la cuenta que ejecuta el flujo</label>
        <input type="email" class="form-control" id="powerAutomateExternalEmail" placeholder="usuario@empresa.com">
      </div>

      <div>
        <label for="powerAutomateWebhookUrl" class="form-label">URL del webhook</label>
        <input type="text" class="form-control" id="powerAutomateWebhookUrl" readonly>
      </div>

      <div>
        <label for="powerAutomateHeaderName" class="form-label">Encabezado de autenticacion</label>
        <input type="text" class="form-control" id="powerAutomateHeaderName" readonly>
      </div>

      <div>
        <label for="powerAutomateToken" class="form-label">Clave secreta</label>
        <input type="text" class="form-control" id="powerAutomateToken" readonly placeholder="Genera una nueva clave para verla aqui una sola vez">
      </div>

      <div class="workspace-callout" id="powerAutomateStatus">La integracion aun no tiene clave activa.</div>
      <div class="workspace-callout subtle">
        Flujo recomendado: recurrencia diaria -> obtener eventos de hoy desde Outlook Calendar -> seleccionar solo titulo, horas, ubicacion y URL -> POST JSON a este webhook con <code>date</code>, <code>source_email</code> y <code>sessions</code>.
      </div>
    </div>
  </article>

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
    <div class="workspace-callout subtle" id="teamsTodaySummary">Sin datos sincronizados hoy.</div>
    <div class="workspace-list" id="calendarSources"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>