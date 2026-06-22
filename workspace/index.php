<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Panel principal', 'index', $user);
?>
<div class="workspace-cards" id="summaryCards"></div>

<div class="workspace-grid">
  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Hoy</span>
        <h3>Tareas priorizadas</h3>
      </div>
      <a class="btn btn-sm btn-outline-light summary-module-trigger" data-module="tasks" href="tasks.php">Gestionar tareas</a>
    </div>
    <div class="workspace-list" id="summaryTasks"></div>
  </article>

  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">En foco</span>
        <h3>Metas activas</h3>
      </div>
      <a class="btn btn-sm btn-outline-light summary-module-trigger" data-module="goals" href="goals.php">Ver metas</a>
    </div>
    <div class="workspace-list" id="summaryGoals"></div>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Agenda</span>
        <h3>Eventos del calendario</h3>
      </div>
      <a class="btn btn-sm btn-outline-light summary-module-trigger" data-module="calendar" href="calendar.php">Abrir calendario</a>
    </div>
    <div class="workspace-list" id="summaryEvents"></div>
  </article>

  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Contexto</span>
        <h3>Notas importantes</h3>
      </div>
      <a class="btn btn-sm btn-outline-light summary-module-trigger" data-module="notes" href="notes.php">Abrir notas</a>
    </div>
    <div class="workspace-list" id="summaryNotes"></div>
  </article>
</div>

<div class="ws-module-modal-backdrop" id="summaryModuleBackdrop" hidden>
  <article class="ws-module-modal" role="dialog" aria-modal="true" aria-labelledby="summaryModuleTitle">
    <button type="button" class="ws-module-modal-close" id="summaryModuleClose" aria-label="Cerrar">
      <i class="fas fa-xmark"></i>
    </button>
    <span class="ws-module-modal-kicker" id="summaryModuleKicker">Modulo</span>
    <h3 id="summaryModuleTitle">Resumen del modulo</h3>
    <p id="summaryModuleDescription">Consulta los indicadores clave y continua trabajando con mayor contexto.</p>
    <div class="ws-module-modal-points" id="summaryModulePoints"></div>
    <a id="summaryModuleLink" class="ws-module-modal-link" href="calendar.php">Ir directamente al modulo</a>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>