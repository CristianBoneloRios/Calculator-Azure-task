<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Notas importantes', 'notes', $user);
?>
<div class="workspace-grid">
  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Captura</span>
        <h3>Nueva nota</h3>
      </div>
    </div>
    <form id="noteForm" class="workspace-form">
      <div>
        <label for="noteTitle" class="form-label">Titulo</label>
        <input type="text" class="form-control" id="noteTitle" required>
      </div>
      <div>
        <label for="noteContent" class="form-label">Contenido</label>
        <textarea class="form-control" id="noteContent" rows="7" required></textarea>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="noteColor" class="form-label">Color</label>
          <select id="noteColor" class="form-select">
            <option value="blue">Azul</option>
            <option value="amber">Ambar</option>
            <option value="green">Verde</option>
            <option value="pink">Rosado</option>
          </select>
        </div>
        <div class="d-flex align-items-end">
          <div class="form-check mb-2">
            <input class="form-check-input" type="checkbox" id="notePinned">
            <label class="form-check-label" for="notePinned">Fijar arriba</label>
          </div>
        </div>
      </div>
      <button type="submit" class="btn btn-primary">Guardar nota</button>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Repositorio</span>
        <h3>Notas guardadas</h3>
      </div>
    </div>
    <div class="workspace-list" id="notesList"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>