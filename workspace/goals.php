<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Metas', 'goals', $user);
?>
<div class="workspace-grid">
  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Direccion</span>
        <h3>Nueva meta</h3>
      </div>
    </div>
    <form id="goalForm" class="workspace-form">
      <div>
        <label for="goalTitle" class="form-label">Titulo</label>
        <input type="text" class="form-control" id="goalTitle" required>
      </div>
      <div>
        <label for="goalDescription" class="form-label">Descripcion</label>
        <textarea class="form-control" id="goalDescription" rows="5"></textarea>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="goalTargetDate" class="form-label">Fecha objetivo</label>
          <input type="date" class="form-control" id="goalTargetDate">
        </div>
        <div>
          <label for="goalProgress" class="form-label">Avance %</label>
          <input type="number" class="form-control" id="goalProgress" min="0" max="100" value="0">
        </div>
      </div>
      <div>
        <label for="goalStatus" class="form-label">Estado</label>
        <select id="goalStatus" class="form-select">
          <option value="active">Activa</option>
          <option value="paused">En pausa</option>
          <option value="completed">Completada</option>
        </select>
      </div>
      <button type="submit" class="btn btn-primary">Guardar meta</button>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Seguimiento</span>
        <h3>Metas registradas</h3>
      </div>
    </div>
    <div class="workspace-list" id="goalsList"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>