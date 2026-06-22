<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Calendario', 'calendar', $user);
?>
<div class="workspace-grid">
  <article class="workspace-panel cal-workspace-full" id="calMainPanel">
    <div class="cal-shell">
      <aside class="cal-sidebar-left">
        <div class="cal-mini-head">
          <button class="btn btn-sm btn-outline" id="calMiniPrev" aria-label="Mes anterior mini"><i class="fas fa-chevron-left"></i></button>
          <strong id="calMiniMonthLabel">Cargando...</strong>
          <button class="btn btn-sm btn-outline" id="calMiniNext" aria-label="Mes siguiente mini"><i class="fas fa-chevron-right"></i></button>
        </div>

        <div class="cal-mini-week" aria-hidden="true">
          <span>L</span><span>M</span><span>X</span><span>J</span><span>V</span><span>S</span><span>D</span>
        </div>
        <div class="cal-mini-grid" id="calMiniGrid"></div>

        <div class="cal-sidebar-block">
          <h4>Mis calendarios</h4>
          <label class="cal-filter-item"><input type="checkbox" id="calFilterManual" checked> Calendario</label>
          <label class="cal-filter-item"><input type="checkbox" id="calFilterTask" checked> Tareas</label>
          <label class="cal-filter-item"><input type="checkbox" id="calFilterMeeting" checked> Reuniones</label>
          <label class="cal-filter-item"><input type="checkbox" id="calFilterSession" checked> Sesiones</label>
          <label class="cal-filter-item"><input type="checkbox" id="calFilterTeams" checked> Teams</label>
        </div>
      </aside>

      <section class="cal-center">
        <div class="workspace-panel-header cal-topbar">
          <div>
            <span class="eyebrow">Calendario</span>
            <h3 id="calMonthTitle">Cargando...</h3>
          </div>
          <div class="cal-nav">
            <button class="btn btn-sm btn-outline" id="calPrevBtn" aria-label="Mes anterior"><i class="fas fa-chevron-left"></i></button>
            <button class="btn btn-sm btn-outline" id="calTodayBtn">Hoy</button>
            <button class="btn btn-sm btn-outline" id="calNextBtn" aria-label="Mes siguiente"><i class="fas fa-chevron-right"></i></button>
          </div>
        </div>

        <div class="cal-grid-wrap">
          <div class="cal-week-header" aria-hidden="true">
            <span>Lunes</span><span>Martes</span><span>Miercoles</span><span>Jueves</span><span>Viernes</span><span>Sabado</span><span>Domingo</span>
          </div>
          <div class="cal-grid" id="calGrid" role="grid"></div>
        </div>

        <div class="cal-legend">
          <span class="cal-legend-item"><span class="cal-dot task"></span>Tarea</span>
          <span class="cal-legend-item"><span class="cal-dot meeting"></span>Reunion</span>
          <span class="cal-legend-item"><span class="cal-dot session"></span>Sesion</span>
          <span class="cal-legend-item"><span class="cal-dot teams"></span>Teams</span>
        </div>
      </section>

      <aside class="cal-sidebar-right">
        <div class="workspace-panel-header">
          <div>
            <span class="eyebrow">Agenda</span>
            <h3 id="calAgendaDate">Selecciona un dia</h3>
          </div>
        </div>

        <div class="workspace-actions cal-agenda-actions">
          <button class="btn btn-sm btn-primary" id="calNewEventBtn"><i class="fas fa-plus"></i> Nuevo</button>
          <button class="btn btn-sm btn-outline" id="calSyncNowBtn"><i class="fas fa-rotate"></i> Sincronizar</button>
        </div>

        <div class="workspace-callout subtle" id="calSyncStatus">Sincronizacion inactiva.</div>

        <div class="workspace-list cal-agenda-list" id="calAgendaList"></div>

        <div class="workspace-callout subtle" id="teamsTodaySummary">Cargando...</div>
      </aside>
    </div>
  </article>

  <article class="workspace-panel cal-workspace-full">
    <div class="workspace-panel-header" style="margin-bottom:10px">
      <div>
        <span class="eyebrow">Power Automate</span>
        <h3>Webhook &amp; Teams</h3>
      </div>
      <button type="button" class="btn btn-sm btn-primary" id="powerAutomateRotateKey">Generar clave</button>
    </div>

    <div class="workspace-form cal-pa-form">
      <div>
        <label for="powerAutomateExternalEmail" class="form-label">Correo del flujo entrante</label>
        <input type="email" class="form-control" id="powerAutomateExternalEmail" placeholder="usuario@empresa.com">
      </div>
      <div>
        <label for="powerAutomateWebhookUrl" class="form-label">URL webhook (Teams -> sistema)</label>
        <input type="text" class="form-control" id="powerAutomateWebhookUrl" readonly>
      </div>
      <div>
        <label for="powerAutomateHeaderName" class="form-label">Encabezado de autenticacion</label>
        <input type="text" class="form-control" id="powerAutomateHeaderName" readonly>
      </div>
      <div>
        <label for="powerAutomateToken" class="form-label">Clave secreta</label>
        <input type="text" class="form-control" id="powerAutomateToken" readonly placeholder="Genera una clave para verla una sola vez">
      </div>
      <div>
        <label for="powerAutomateOutboundUrl" class="form-label">
          URL flujo de salida <span class="cal-pa-label-hint">(sistema -> Teams)</span>
        </label>
        <input type="url" class="form-control" id="powerAutomateOutboundUrl" placeholder="https://prod-xx.logic.azure.com/...">
      </div>
      <button type="button" class="btn btn-sm btn-outline cal-pa-save-btn" id="saveOutboundUrlBtn">
        <i class="fas fa-save"></i> Guardar URL de salida
      </button>
      <div class="workspace-callout" id="powerAutomateStatus">Sin clave activa.</div>
    </div>
  </article>
</div>

<!-- ── Day popup ──────────────────────────────────────── -->
<div id="calDayPopup" class="cal-popup" hidden role="dialog" aria-modal="true" aria-labelledby="calPopupDate">
  <div class="cal-popup-inner">

    <div class="cal-popup-header">
      <strong id="calPopupDate"></strong>
      <button class="cal-popup-close" id="calPopupClose" aria-label="Cerrar"><i class="fas fa-xmark"></i></button>
    </div>

    <div id="calPopupEvents" class="cal-popup-events"></div>

    <p class="cal-popup-add-label">Agregar nuevo</p>
    <div class="cal-popup-actions">
      <button class="cal-type-btn task"    data-type="task">   <i class="fas fa-check-square"></i> Tarea</button>
      <button class="cal-type-btn meeting" data-type="meeting"><i class="fas fa-video"></i> Reunion</button>
      <button class="cal-type-btn session" data-type="session"><i class="fas fa-clock"></i> Sesion</button>
    </div>

    <form id="calQuickForm" class="cal-quick-form" hidden autocomplete="off">
      <input type="hidden" id="calQuickType">
      <input type="hidden" id="calQuickDate">
      <input type="text" class="form-control" id="calQuickTitle" placeholder="Titulo del evento" required>
      <div class="cal-time-row">
        <input type="time" class="form-control" id="calQuickStart" value="09:00">
        <span>→</span>
        <input type="time" class="form-control" id="calQuickEnd" value="10:00">
      </div>
      <input type="url" class="form-control" id="calQuickUrl" placeholder="URL de Teams (opcional)">
      <div class="cal-teams-sync-row" id="calTeamsSyncRow" style="display:none">
        <input type="checkbox" id="calTeamsSyncCheck">
        <i class="fas fa-users"></i>
        <label for="calTeamsSyncCheck">Sincronizar con Teams</label>
      </div>
      <div class="cal-quick-btns">
        <button type="submit" class="btn btn-sm btn-primary">Guardar</button>
        <button type="button" class="btn btn-sm btn-outline" id="calQuickCancel">Cancelar</button>
      </div>
    </form>

  </div>
</div>
<?php workspaceLayoutEnd(); ?>