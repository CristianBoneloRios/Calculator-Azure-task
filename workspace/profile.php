<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Perfil y foto', 'profile', $user);
?>
<div class="workspace-grid">
  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Cuenta</span>
        <h3>Perfil del usuario</h3>
      </div>
    </div>
    <form id="profileForm" class="workspace-form">
      <div class="workspace-form-grid two">
        <div>
          <label for="profileFullName" class="form-label">Nombre completo</label>
          <input type="text" class="form-control" id="profileFullName" required>
        </div>
        <div>
          <label for="profileEmail" class="form-label">Correo</label>
          <input type="email" class="form-control" id="profileEmail" required>
        </div>
      </div>
      <div>
        <label for="profilePassword" class="form-label">Nueva contrasena</label>
        <input type="password" class="form-control" id="profilePassword" placeholder="Dejar vacio para no cambiarla">
      </div>
      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary">Guardar perfil</button>
      </div>
    </form>
  </article>

  <article class="workspace-panel tall">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Foto</span>
        <h3>Imagen de perfil</h3>
      </div>
    </div>
    <form id="photoForm" class="workspace-form">
      <img id="profilePhotoPreview" class="workspace-photo-preview" alt="Foto de perfil">
      <div>
        <label for="profilePhotoInput" class="form-label">Subir nueva foto</label>
        <input type="file" class="form-control" id="profilePhotoInput" accept="image/*">
      </div>
      <div class="form-check">
        <input class="form-check-input" type="checkbox" value="1" id="makePublicProfilePhoto" checked>
        <label class="form-check-label" for="makePublicProfilePhoto">Usar tambien esta foto para el bloque publico de Cristian</label>
      </div>
      <button type="submit" class="btn btn-outline-light">Actualizar foto</button>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Portada publica</span>
        <h3>Datos visibles en la pagina principal</h3>
      </div>
    </div>
    <form id="publicProfileForm" class="workspace-form">
      <div class="workspace-form-grid two">
        <div>
          <label for="publicDisplayName" class="form-label">Nombre visible</label>
          <input type="text" class="form-control" id="publicDisplayName" required>
        </div>
        <div>
          <label for="publicRoleTitle" class="form-label">Cargo visible</label>
          <input type="text" class="form-control" id="publicRoleTitle" required>
        </div>
      </div>
      <div>
        <label for="publicCompanyName" class="form-label">Empresa</label>
        <input type="text" class="form-control" id="publicCompanyName">
      </div>
      <div>
        <label for="publicBio" class="form-label">Descripcion</label>
        <textarea class="form-control" id="publicBio" rows="5"></textarea>
      </div>
      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary">Actualizar portada</button>
      </div>
    </form>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>