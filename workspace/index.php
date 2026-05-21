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
      <a class="btn btn-sm btn-outline-light" href="tasks.php">Gestionar tareas</a>
    </div>
    <div class="workspace-list" id="summaryTasks"></div>
  </article>

  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">En foco</span>
        <h3>Metas activas</h3>
      </div>
      <a class="btn btn-sm btn-outline-light" href="goals.php">Ver metas</a>
    </div>
    <div class="workspace-list" id="summaryGoals"></div>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Agenda</span>
        <h3>Eventos del calendario</h3>
      </div>
      <a class="btn btn-sm btn-outline-light" href="calendar.php">Abrir calendario</a>
    </div>
    <div class="workspace-list" id="summaryEvents"></div>
  </article>

  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Contexto</span>
        <h3>Notas importantes</h3>
      </div>
      <a class="btn btn-sm btn-outline-light" href="notes.php">Abrir notas</a>
    </div>
    <div class="workspace-list" id="summaryNotes"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>