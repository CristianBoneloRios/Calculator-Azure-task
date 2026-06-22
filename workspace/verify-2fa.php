<?php

declare(strict_types=1);

try {
  require_once __DIR__ . '/../api/app.php';
  ensureApplicationInstalled();

  // Verificar que el usuario está en el proceso de verificación 2FA
  if (!isset($_SESSION['_2fa_user_id'])) {
    header('Location: login.php');
    exit;
  }
} catch (Throwable $throwable) {
  $bootstrapError = 'No se pudo inicializar: ' . $throwable->getMessage();
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Verificar 2FA | Azure Task Suite</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="https://cdn-uicons.flaticon.com/uicons-bold-rounded/css/uicons-bold-rounded.css" />
  <link rel="stylesheet" href="../styles.css">
</head>
<body class="auth-page-body">
  <header class="app-header">
    <div class="header-left">
      <a class="sidebar-toggle-btn" href="login.php" title="Volver">
        <i class="fas fa-arrow-left"></i>
      </a>
      <div class="brand">
        <div class="brand-icon">
          <i class="fi fi-br-rocket-lunch"></i>
        </div>
        <div class="brand-text">
          <span class="brand-name">Azure Task Suite</span>
          <span class="brand-sub">Verificacion 2FA</span>
        </div>
      </div>
    </div>
    <div class="header-center">
      <div class="header-chips">
        <span class="chip chip-green"><i class="fas fa-lock"></i> Autenticacion 2FA</span>
      </div>
    </div>
    <div class="header-right">
      <a href="login.php" class="btn btn-outline">
        <i class="fas fa-redo"></i> Intentar otra cuenta
      </a>
    </div>
  </header>

  <main class="auth-main">
    <section class="auth-card">
      <div class="auth-aside">
        <span class="auth-kicker">Seguridad activada</span>
        <h1>Verifica tu identidad</h1>
        <p>Ingresa el codigo de 6 digitos de tu autenticador para continuar.</p>
        <ul class="auth-list">
          <li><i class="fas fa-shield-alt"></i> Proteccion de dos factores</li>
          <li><i class="fas fa-mobile-alt"></i> Usa Google Authenticator o Authy</li>
          <li><i class="fas fa-key"></i> Codigo unico por sesion</li>
        </ul>

        <div class="login-footer-animation">
          <div class="login-icon-circle-footer">
            <i class="fas fa-shield-alt"></i>
          </div>
          <div class="login-ripple-footer"></div>
          <div class="login-ripple-footer delay-1"></div>
        </div>
      </div>

      <div class="auth-form-panel">
        <h2>Codigo 2FA</h2>
        <p class="auth-muted">Abre tu aplicacion de autenticacion y escribe el codigo temporal.</p>

        <div id="verify2FAStatus" class="auth-status" aria-live="polite" hidden></div>

        <form id="verify2FAForm" class="auth-form">
          <div class="auth-otp-wrap">
            <label for="twoFactorCode" class="auth-otp-label">Codigo de 6 digitos</label>
            <input type="text" id="twoFactorCode" class="auth-otp-input" placeholder="000000" maxlength="6" pattern="[0-9]{6}" inputmode="numeric" required autofocus>
            <p class="auth-otp-help">Tip: el codigo cambia cada 30 segundos en tu app autenticadora.</p>
          </div>
          <div class="auth-actions">
            <button type="submit" class="btn btn-primary" id="verify2FASubmitBtn"><i class="fas fa-check-circle"></i> Verificar codigo</button>
            <a href="login.php" class="btn btn-outline">Cancelar</a>
          </div>
        </form>

        <div class="auth-switch">
          <span>Problemas para acceder?</span>
          <a href="login.php">Vuelve al login</a>
        </div>
      </div>
    </section>
  </main>

  <footer class="app-footer">
    <div class="footer-content">
      <div class="footer-section footer-primary"></div>
    </div>
  </footer>

  <script>
    const verify2FAForm = document.getElementById('verify2FAForm');
    const verify2FAStatus = document.getElementById('verify2FAStatus');
    const verify2FASubmitBtn = document.getElementById('verify2FASubmitBtn');
    const twoFactorCodeInput = document.getElementById('twoFactorCode');

    const setVerifyStatus = (type, message) => {
      if (!verify2FAStatus) return;
      verify2FAStatus.hidden = false;
      verify2FAStatus.className = `auth-status ${type}`;

      let icon = 'fa-circle-info';
      if (type === 'loading') icon = 'fa-circle-notch fa-spin';
      if (type === 'success') icon = 'fa-check-circle';
      if (type === 'error') icon = 'fa-triangle-exclamation';

      verify2FAStatus.innerHTML = `<i class="fas ${icon}"></i><span>${message}</span>`;
    };

    verify2FAForm.addEventListener('submit', async event => {
      event.preventDefault();

      const code = twoFactorCodeInput.value.trim();

      if (!/^[0-9]{6}$/.test(code)) {
        setVerifyStatus('error', 'El codigo debe tener 6 digitos numericos.');
        return;
      }

      try {
        verify2FASubmitBtn.disabled = true;
        setVerifyStatus('loading', 'Validando codigo de seguridad...');

        const response = await fetch('../api/auth.php?action=2fa-verify', {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ code })
        });

        let data = null;
        try {
          data = await response.json();
        } catch (_) {
          data = null;
        }

        if (!response.ok || !data || data.ok === false) {
          const message = data && data.message ? data.message : `Error HTTP ${response.status}`;
          setVerifyStatus('error', `${message} Verifica que el codigo sea correcto.`);
          verify2FASubmitBtn.disabled = false;
          return;
        }

        setVerifyStatus('success', 'Codigo validado. Redirigiendo al workspace...');
        window.location.href = 'index.php';
      } catch (error) {
        setVerifyStatus('error', `Error de conexion: ${error.message}`);
        verify2FASubmitBtn.disabled = false;
      }
    });

    twoFactorCodeInput.addEventListener('input', () => {
      twoFactorCodeInput.value = twoFactorCodeInput.value.replace(/\D/g, '').slice(0, 6);
    });
  </script>
</body>
</html>
