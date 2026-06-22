<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Notas importantes', 'notes', $user);
?>
<div class="workspace-grid notes-module-grid">
  <article class="workspace-panel tall notes-compose-panel">
    <div class="workspace-panel-header notes-panel-header">
      <div>
        <span class="eyebrow">Captura</span>
        <h3>Nueva nota compacta</h3>
      </div>
      <span class="workspace-tag"><i class="fas fa-wand-magic-sparkles"></i> Nuevo estilo</span>
    </div>

    <form id="noteForm" class="workspace-form notes-form-compact">
      <div>
        <label for="noteTitle" class="form-label">Titulo</label>
        <input type="text" class="form-control" id="noteTitle" maxlength="180" required>
      </div>
      <div>
        <label for="noteContent" class="form-label">Contenido</label>
        <textarea class="form-control" id="noteContent" rows="6" maxlength="4000" required></textarea>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="noteColor" class="form-label">Color</label>
          <select id="noteColor" class="form-select">
            <option value="blue">Azul</option>
            <option value="yellow">Amarillo</option>
            <option value="purple">Morado</option>
            <option value="cyan">Cian</option>
          </select>
        </div>
        <div class="d-flex align-items-end">
          <div class="form-check mb-2">
            <input class="form-check-input" type="checkbox" id="notePinned">
            <label class="form-check-label" for="notePinned">Fijar arriba</label>
          </div>
        </div>
      </div>

      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary"><i class="fas fa-floppy-disk"></i> Guardar nota</button>
      </div>
    </form>

    <div class="workspace-callout subtle notes-help-box">
      <i class="fas fa-lightbulb"></i>
      Usa comentarios jerarquicos para seguimiento y comparte notas por invitacion de correo.
    </div>
  </article>

  <article class="workspace-panel wide notes-repository-panel">
    <div class="workspace-panel-header notes-panel-header">
      <div>
        <span class="eyebrow">Repositorio</span>
        <h3>Notas guardadas</h3>
      </div>
      <div class="notes-header-metrics" id="notesHeaderMetrics">0 notas</div>
    </div>

    <div class="notes-toolbar">
      <div class="notes-toolbar-search">
        <label for="notesSearchInput" class="form-label">Buscar</label>
        <input type="search" class="form-control" id="notesSearchInput" placeholder="Buscar por titulo o contenido...">
      </div>
      <div class="notes-toolbar-filter">
        <label for="notesColorFilter" class="form-label">Color</label>
        <select id="notesColorFilter" class="form-select">
          <option value="all">Todos</option>
          <option value="blue">Azul</option>
          <option value="yellow">Amarillo</option>
          <option value="purple">Morado</option>
          <option value="cyan">Cian</option>
        </select>
      </div>
      <div class="notes-toolbar-filter">
        <label for="notesScopeFilter" class="form-label">Tipo</label>
        <select id="notesScopeFilter" class="form-select">
          <option value="all">Todas</option>
          <option value="mine">Mias</option>
          <option value="shared">Compartidas</option>
        </select>
      </div>
    </div>

    <div class="workspace-list" id="notesList"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>