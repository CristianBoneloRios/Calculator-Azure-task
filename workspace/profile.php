<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Perfil y foto', 'profile', $user);
?>
<div class="workspace-grid profile-page-grid">
  <article class="workspace-panel wide profile-identity-panel">
    <div class="workspace-panel-header profile-panel-header">
      <div>
        <span class="eyebrow">Cuenta</span>
        <h3>Perfil profesional</h3>
      </div>
      <span class="workspace-tag"><i class="fas fa-user-shield"></i> Seguridad activa</span>
    </div>

    <div class="profile-identity-summary">
      <div class="profile-identity-avatar-wrap">
        <img id="profileIdentityPhoto" class="profile-identity-avatar" alt="Foto del usuario">
        <span id="profileIdentityFallback" class="profile-identity-fallback"><i class="fas fa-user"></i></span>
      </div>
      <div class="profile-identity-meta">
        <h4 id="profileIdentityName">Usuario del Workspace</h4>
        <p id="profileIdentityEmail">correo@ejemplo.com</p>
        <div class="profile-security-badges">
          <span class="profile-badge" id="profile2FAStateBadge"><i class="fas fa-shield"></i> 2FA pendiente</span>
          <span class="profile-badge subtle"><i class="fas fa-lock"></i> Sesion protegida</span>
        </div>
      </div>
    </div>

    <form id="profileForm" class="workspace-form profile-account-form">
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
      <div class="workspace-form-grid two">
        <div>
          <label for="profilePassword" class="form-label">Nueva contrasena</label>
          <input type="password" class="form-control" id="profilePassword" placeholder="Dejar vacio para no cambiarla" autocomplete="new-password">
        </div>
        <div>
          <label for="profilePasswordConfirm" class="form-label">Confirmar contrasena</label>
          <input type="password" class="form-control" id="profilePasswordConfirm" placeholder="Repite la nueva contrasena" autocomplete="new-password">
        </div>
      </div>
      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary"><i class="fas fa-floppy-disk"></i> Guardar perfil</button>
      </div>
    </form>
  </article>

  <article class="workspace-panel tall profile-photo-panel">
    <div class="workspace-panel-header profile-panel-header">
      <div>
        <span class="eyebrow">Foto</span>
        <h3>Avatar profesional</h3>
      </div>
      <span class="profile-file-meta" id="profilePhotoMeta">PNG/JPG/WebP hasta 5MB</span>
    </div>

    <form id="photoForm" class="workspace-form profile-photo-form">
      <div class="profile-photo-stage">
        <img id="profilePhotoPreview" class="profile-photo-preview-large" alt="Foto de perfil">
        <span id="profilePhotoFallback" class="profile-photo-fallback-large"><i class="fas fa-user"></i></span>
      </div>

      <div class="profile-upload-dropzone" id="profileUploadDropzone" role="button" tabindex="0" aria-label="Subir foto de perfil">
        <i class="fas fa-cloud-arrow-up"></i>
        <strong>Arrastra tu imagen o haz clic para seleccionar</strong>
        <p>Recomendado: cuadrada de 512x512 para mejor nitidez.</p>
        <input type="file" id="profilePhotoInput" accept="image/png,image/jpeg,image/webp" hidden>
      </div>

      <div class="form-check">
        <input class="form-check-input" type="checkbox" value="1" id="makePublicProfilePhoto" checked>
        <label class="form-check-label" for="makePublicProfilePhoto">Usar tambien esta foto para el bloque publico de Cristian</label>
      </div>

      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary" id="profilePhotoSubmitBtn"><i class="fas fa-upload"></i> Publicar foto</button>
        <button type="button" class="btn btn-outline-light" id="profilePhotoResetBtn"><i class="fas fa-rotate-left"></i> Revertir</button>
      </div>
    </form>
  </article>

  <article class="workspace-panel wide profile-security-panel">
    <div class="workspace-panel-header profile-panel-header">
      <div>
        <span class="eyebrow">Seguridad</span>
        <h3>Autenticacion de dos factores (2FA)</h3>
      </div>
      <span class="workspace-tag" id="profile2FATag"><i class="fas fa-lock"></i> Configuracion requerida</span>
    </div>

    <div class="profile-2fa-grid">
      <div class="profile-2fa-left">
        <p class="workspace-muted">Activa 2FA para proteger tu cuenta con un codigo temporal desde Google Authenticator, Authy u otra app TOTP.</p>
        <div class="profile-2fa-toggle-row">
          <label class="profile-2fa-toggle" for="profile2FAToggle">
            <input type="checkbox" id="profile2FAToggle">
            <span class="profile-2fa-toggle-slider"></span>
          </label>
          <div>
            <strong>Activar/Desactivar autenticacion 2FA</strong>
            <p id="profile2FAHint" class="workspace-muted">Al activar se genera QR para registrar la app autenticadora.</p>
          </div>
        </div>
        <div class="workspace-actions">
          <button type="button" class="btn btn-primary" id="profile2FAGenerateBtn"><i class="fas fa-qrcode"></i> Generar QR</button>
          <button type="button" class="btn btn-outline-light" id="profile2FADisableBtn"><i class="fas fa-shield-slash"></i> Desactivar 2FA</button>
        </div>
        <div class="profile-2fa-secret-box" id="profile2FASecretBox" hidden>
          <label class="form-label" for="profile2FASecret">Clave manual</label>
          <div class="profile-2fa-secret-input-wrap">
            <input type="text" class="form-control" id="profile2FASecret" readonly>
            <button type="button" class="btn btn-sm btn-outline-light" id="profile2FACopyBtn"><i class="fas fa-copy"></i> Copiar</button>
          </div>
        </div>
      </div>

      <div class="profile-2fa-right">
        <div class="profile-2fa-qr-wrap" id="profile2FAQRWrap" hidden>
          <img id="profile2FAQR" alt="Codigo QR para 2FA">
          <p class="profile-2fa-qr-note">Escanea este QR en Google Authenticator, Authy o Microsoft Authenticator.</p>
        </div>
        <div class="workspace-form" id="profile2FAConfirmBox" hidden>
          <div>
            <label for="profile2FACode" class="form-label">Codigo de verificacion</label>
            <input type="text" class="form-control" id="profile2FACode" inputmode="numeric" maxlength="6" placeholder="000000">
          </div>
          <button type="button" class="btn btn-primary" id="profile2FAEnableBtn"><i class="fas fa-check-circle"></i> Activar 2FA</button>
        </div>
      </div>
    </div>
  </article>

  <article class="workspace-panel wide profile-developer-panel">
    <div class="workspace-panel-header profile-panel-header">
      <div>
        <span class="eyebrow">Ajustes admin</span>
        <h3>Perfil del desarrollador (separado del usuario)</h3>
      </div>
      <span class="workspace-tag" id="profileDeveloperOwnerTag"><i class="fas fa-user-shield"></i> Solo admin propietario</span>
    </div>

    <p class="workspace-muted">Esta foto no depende del usuario logueado. Solo el admin propietario puede modificarla o transferir propiedad a otro admin.</p>

    <div class="profile-developer-grid">
      <div class="profile-developer-preview-card">
        <div class="about-me-btn" style="cursor:default; margin-bottom:0;">
          <div class="about-me-btn-avatar" id="profileDeveloperAvatarWrap">
            <img id="profileDeveloperAvatar" alt="Foto del desarrollador" style="display:none;">
            <i id="profileDeveloperAvatarFallback" class="fas fa-user"></i>
          </div>
          <div class="about-me-btn-info">
            <span class="about-me-btn-name" id="profileDeveloperName">Desarrollador</span>
            <span class="about-me-btn-role" id="profileDeveloperRole">Developer</span>
            <span class="workspace-muted" id="profileDeveloperOwnerEmail"></span>
          </div>
        </div>
      </div>

      <div class="profile-developer-controls">
        <form id="profileDeveloperPhotoForm" class="workspace-form">
          <div>
            <label for="profileDeveloperPhotoInput" class="form-label">Cambiar foto del desarrollador</label>
            <input type="file" class="form-control" id="profileDeveloperPhotoInput" accept="image/png,image/jpeg,image/webp">
          </div>
          <div class="workspace-actions">
            <button type="submit" class="btn btn-primary" id="profileDeveloperPhotoSaveBtn"><i class="fas fa-image"></i> Guardar foto desarrollador</button>
          </div>
        </form>

        <form id="profileDeveloperOwnerForm" class="workspace-form">
          <div>
            <label for="profileDeveloperOwnerUserId" class="form-label">Transferir propiedad a otro admin</label>
            <select class="form-select" id="profileDeveloperOwnerUserId"></select>
          </div>
          <div class="workspace-actions">
            <button type="submit" class="btn btn-outline-light" id="profileDeveloperOwnerSaveBtn"><i class="fas fa-exchange-alt"></i> Transferir propiedad</button>
          </div>
        </form>

        <form id="profileDeveloperPromoteAdminForm" class="workspace-form">
          <div>
            <label for="profileDeveloperPromoteAdminEmail" class="form-label">Crear nuevo admin desde cuenta existente</label>
            <input type="email" class="form-control" id="profileDeveloperPromoteAdminEmail" placeholder="correo@dominio.com">
          </div>
          <div class="workspace-actions">
            <button type="submit" class="btn btn-outline-light" id="profileDeveloperPromoteAdminBtn"><i class="fas fa-user-plus"></i> Convertir en admin</button>
          </div>
        </form>

        <div class="workspace-callout subtle" id="profileDeveloperState">Cargando configuracion de desarrollador...</div>
      </div>
    </div>
  </article>

  <article class="workspace-panel wide profile-doc-security-panel">
    <div class="workspace-panel-header profile-panel-header">
      <div>
        <span class="eyebrow">Documentos sensibles</span>
        <h3>Clave de acceso secundaria</h3>
      </div>
      <span class="workspace-tag" id="profileDocSecurityTag"><i class="fas fa-lock-open"></i> Sin clave</span>
    </div>

    <p class="workspace-muted">Configura una clave adicional para proteger documentos generados por IA y Power Automate.</p>

    <form id="profileDocSecurityForm" class="workspace-form">
      <div class="workspace-form-grid two">
        <div>
          <label for="profileDocAccessKey" class="form-label">Nueva clave secundaria</label>
          <input type="password" class="form-control" id="profileDocAccessKey" minlength="6" placeholder="Minimo 6 caracteres" autocomplete="new-password">
        </div>
        <div>
          <label for="profileDocAccessKeyConfirm" class="form-label">Confirmar clave secundaria</label>
          <input type="password" class="form-control" id="profileDocAccessKeyConfirm" minlength="6" placeholder="Repite la clave" autocomplete="new-password">
        </div>
      </div>
      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary" id="profileDocSecuritySaveBtn"><i class="fas fa-key"></i> Guardar clave secundaria</button>
        <button type="button" class="btn btn-outline-light" id="profileDocSecurityRemoveBtn"><i class="fas fa-trash"></i> Eliminar clave</button>
      </div>
      <div class="workspace-callout subtle" id="profileDocSecurityState">Cargando estado de seguridad de documentos...</div>
    </form>
  </article>

  <article class="workspace-panel wide">
    <div class="workspace-panel-header profile-panel-header">
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
        <button type="submit" class="btn btn-primary"><i class="fas fa-pen-to-square"></i> Actualizar portada</button>
      </div>
    </form>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>