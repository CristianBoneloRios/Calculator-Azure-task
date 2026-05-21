<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Tareas del dia', 'tasks', $user);
?>
<div class="workspace-grid">
  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Operacion</span>
        <h3>Registrar tarea</h3>
      </div>
    </div>
    <form id="taskForm" class="workspace-form">
      <div>
        <label for="taskTitle" class="form-label">Titulo</label>
        <input type="text" class="form-control" id="taskTitle" required>
      </div>
      <div>
        <label for="taskDescription" class="form-label">Descripcion</label>
        <textarea class="form-control" id="taskDescription" rows="4"></textarea>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="taskDate" class="form-label">Fecha</label>
          <input type="date" class="form-control" id="taskDate" required>
        </div>
        <div>
          <label for="taskDueTime" class="form-label">Hora</label>
          <input type="time" class="form-control" id="taskDueTime">
        </div>
      </div>
      <div class="workspace-form-grid two">
        <div>
          <label for="taskPriority" class="form-label">Prioridad</label>
          <select id="taskPriority" class="form-select">
            <option value="low">Baja</option>
            <option value="medium" selected>Media</option>
            <option value="high">Alta</option>
          </select>
        </div>
        <div>
          <label for="taskStatus" class="form-label">Estado</label>
          <select id="taskStatus" class="form-select">
            <option value="pending">Pendiente</option>
            <option value="in_progress">En progreso</option>
            <option value="done">Hecha</option>
          </select>
        </div>
      </div>
      <button type="submit" class="btn btn-primary">Guardar tarea</button>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Backlog diario</span>
        <h3>Tareas registradas</h3>
      </div>
    </div>
    <div class="workspace-list" id="tasksList"></div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>