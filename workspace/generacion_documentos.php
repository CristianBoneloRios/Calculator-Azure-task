<?php

declare(strict_types=1);

require_once __DIR__ . '/includes/layout.php';

ensureApplicationInstalled();
$user = requirePageAuth();

workspaceLayoutStart('Generacion automatica de documentos', 'generacion_documentos', $user);
?>
<div class="workspace-grid docs-page-grid">
  <article class="workspace-panel wide docs-upload-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">IA + Power Automate</span>
        <h3>Generador de manuales y guias</h3>
      </div>
      <span class="workspace-tag"><i class="fas fa-shield-halved"></i> Acceso protegido</span>
    </div>

    <form id="docGenerationForm" class="workspace-form docs-generation-form" enctype="multipart/form-data">
      <div>
        <label for="docFiles" class="form-label">Archivos de entrada</label>
        <input type="file" id="docFiles" class="form-control" name="files[]" multiple accept=".pdf,.docx,.txt,.jpg,.jpeg,.png,.mp3,.wav" required>
        <p class="workspace-muted">Formatos permitidos: PDF, DOCX, TXT, JPG, PNG, MP3 y WAV.</p>
      </div>

      <div class="workspace-form-grid two">
        <div>
          <label for="docGenerationType" class="form-label">Tipo de documento a generar</label>
          <select id="docGenerationType" class="form-select" required>
            <option value="manual">Manual</option>
            <option value="guia">Guia</option>
            <option value="informe">Informe</option>
          </select>
        </div>
        <div>
          <label for="docWebhookUrl" class="form-label">Webhook Power Automate</label>
          <div class="docs-inline-config">
            <input type="url" id="docWebhookUrl" class="form-control" placeholder="https://prod-xx.logic.azure.com/...">
            <button type="button" class="btn btn-sm btn-outline-light" id="docSaveWebhookBtn"><i class="fas fa-save"></i> Guardar</button>
          </div>
        </div>
      </div>

      <div>
        <label for="docDescription" class="form-label">Descripcion (opcional)</label>
        <textarea id="docDescription" class="form-control" rows="3" placeholder="Contexto para mejorar la generacion del documento"></textarea>
      </div>

      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary" id="docGenerateBtn"><i class="fas fa-sparkles"></i> Generar documento</button>
      </div>
      <div class="workspace-callout subtle" id="docGenerationStatus">Listo para iniciar generacion.</div>
    </form>
  </article>

  <article class="workspace-panel tall docs-security-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Seguridad</span>
        <h3>Clave secundaria</h3>
      </div>
    </div>

    <div class="workspace-callout" id="docSecurityState">Verificando estado de seguridad...</div>

    <form id="docSecurityQuickForm" class="workspace-form">
      <div>
        <label for="docQuickAccessKey" class="form-label">Clave de acceso documentos</label>
        <input type="password" class="form-control" id="docQuickAccessKey" minlength="6" placeholder="Minimo 6 caracteres">
      </div>
      <div class="workspace-actions">
        <button type="submit" class="btn btn-primary"><i class="fas fa-key"></i> Guardar clave</button>
        <button type="button" class="btn btn-outline-light" id="docQuickRemoveKeyBtn"><i class="fas fa-trash"></i> Eliminar</button>
      </div>
    </form>
  </article>

  <article class="workspace-panel wide docs-history-panel">
    <div class="workspace-panel-header">
      <div>
        <span class="eyebrow">Historial</span>
        <h3>Documentos generados</h3>
      </div>
      <button type="button" class="btn btn-sm btn-outline-light" id="docHistoryReloadBtn"><i class="fas fa-rotate"></i> Actualizar</button>
    </div>

    <div class="workspace-list" id="docJobsList"></div>
  </article>
</div>

<div class="ws-module-modal-backdrop" id="docAccessBackdrop" hidden>
  <article class="ws-module-modal" role="dialog" aria-modal="true" aria-labelledby="docAccessTitle">
    <button type="button" class="ws-module-modal-close" id="docAccessClose" aria-label="Cerrar">
      <i class="fas fa-xmark"></i>
    </button>
    <span class="ws-module-modal-kicker">Acceso protegido</span>
    <h3 id="docAccessTitle">Verificacion de documentos sensibles</h3>
    <p>Ingresa tu clave secundaria para habilitar el acceso al historial y descargas.</p>
    <div class="workspace-form" style="margin-top:12px;">
      <div>
        <label for="docAccessKeyInput" class="form-label">Clave secundaria</label>
        <input type="password" id="docAccessKeyInput" class="form-control" minlength="6" autocomplete="off">
      </div>
      <button type="button" class="btn btn-primary" id="docAccessVerifyBtn"><i class="fas fa-lock-open"></i> Verificar acceso</button>
    </div>
  </article>
</div>
<?php workspaceLayoutEnd(); ?>
