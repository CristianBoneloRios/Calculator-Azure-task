<?php

declare(strict_types=1);

$bootstrapError = null;

try {
  require_once __DIR__ . '/../api/app.php';
  ensureApplicationInstalled();

  if (currentUser() !== null) {
    header('Location: index.php');
    exit;
  }
} catch (Throwable $throwable) {
  $bootstrapError = 'No se pudo inicializar el workspace: ' . $throwable->getMessage();
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Registro | Azure Task Suite</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="https://cdn-uicons.flaticon.com/uicons-bold-rounded/css/uicons-bold-rounded.css" />
  <link rel="stylesheet" href="../styles.css">
</head>
<body class="auth-page-body">
  <header class="app-header">
    <div class="header-left">
      <a class="sidebar-toggle-btn" href="../index.php" title="Volver al inicio">
        <i class="fas fa-arrow-left"></i>
      </a>
      <div class="brand">
        <div class="brand-icon">
          <i class="fi fi-br-rocket-lunch"></i>
        </div>
        <div class="brand-text">
          <span class="brand-name">Azure Task Suite</span>
          <span class="brand-sub">Registro al Workspace</span>
        </div>
      </div>
    </div>
    <div class="header-center">
      <div class="header-chips">
        <span class="chip chip-blue"><i class="fas fa-user-plus"></i> Registro</span>
        <span class="chip chip-green"><i class="fas fa-lock"></i> Seguro</span>
      </div>
    </div>
    <div class="header-right">
      <a href="login.php" class="btn btn-outline">
        <i class="fas fa-right-to-bracket"></i> Ya tengo cuenta
      </a>
    </div>
  </header>

  <main class="auth-main">
    <section class="auth-card">
      <div class="auth-aside">
        <span class="auth-kicker">Nuevo acceso</span>
        <h1>Crea tu cuenta del Workspace</h1>
        <p>Activa tu espacio personal para organizar toda tu operacion diaria sin perder la estetica del sistema principal.</p>
        <ul class="auth-list">
          <li><i class="fas fa-user-pen"></i> Perfil personal editable</li>
          <li><i class="fas fa-list-check"></i> Tareas priorizadas</li>
          <li><i class="fas fa-bullseye"></i> Objetivos por avance</li>
          <li><i class="fas fa-calendar-days"></i> Agenda y calendario</li>
        </ul>
      </div>

      <div class="auth-form-panel">
        <h2>Crear cuenta</h2>
        <p class="auth-muted">Completa tus datos para continuar.</p>
        <div id="registerStatus" class="auth-status" aria-live="polite" hidden></div>

        <?php if ($bootstrapError !== null): ?>
          <div class="auth-alert" role="alert">
            <strong>La aplicacion no pudo iniciar correctamente en el servidor.</strong><br>
            <?php echo htmlspecialchars($bootstrapError, ENT_QUOTES, 'UTF-8'); ?>
          </div>
        <?php endif; ?>

        <form id="registerForm" class="auth-form">
          <div>
            <label for="registerFullName">Nombre completo</label>
            <input type="text" id="registerFullName" placeholder="Ej: Cristian Bonelo" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
          </div>
          <div>
            <label for="registerEmail">Correo</label>
            <input type="email" id="registerEmail" placeholder="tu@correo.com" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
          </div>
          <div class="auth-two-cols">
            <div>
              <label for="registerPassword">Contrasena</label>
              <input type="password" id="registerPassword" minlength="8" placeholder="Minimo 8 caracteres" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
            </div>
            <div>
              <label for="registerPasswordConfirm">Confirmar contrasena</label>
              <input type="password" id="registerPasswordConfirm" minlength="8" placeholder="Repite la contrasena" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
            </div>
          </div>
          <div class="auth-actions">
            <button type="submit" class="btn btn-primary" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?>><i class="fas fa-user-plus"></i> Crear cuenta</button>
            <a href="login.php" class="btn btn-outline">Ir al login</a>
          </div>
        </form>

        <div class="auth-switch">
          <span>Ya tienes cuenta?</span>
          <a href="login.php">Inicia sesion</a>
        </div>
      </div>
    </section>
  </main>

  <footer class="app-footer">
    <div class="footer-content">
      <div class="footer-section footer-primary">
        <div class="footer-left"><i class="fas fa-code"></i></div>
        <div class="footer-center">
          <span class="footer-powered">Powered by</span>
          <span class="footer-name">Cristian Jesus Bonelo Rios</span>
          <span class="footer-sep">|</span>
          <span class="footer-role">SOFTWARE QUALITY ANALYST</span>
          <span class="footer-sep">|</span>
          <span class="footer-dept">DEVELOPMENT &amp; INNOVATION</span>
        </div>
        <div class="footer-right"><span class="footer-version">Workspace</span></div>
      </div>
    </div>
  </footer>

  <script>
    const registerForm = document.getElementById('registerForm');
    const registerSubmitBtn = registerForm.querySelector('button[type="submit"]');
    const registerStatus = document.getElementById('registerStatus');

    const setRegisterStatus = (message, type = 'loading') => {
      registerStatus.hidden = false;
      registerStatus.className = `auth-status ${type}`;
      registerStatus.innerHTML = message;
    };

    const setRegisterLoadingState = isLoading => {
      const inputs = registerForm.querySelectorAll('input, button, a');
      inputs.forEach(element => {
        element.toggleAttribute('disabled', isLoading);
      });
      registerSubmitBtn.innerHTML = isLoading
        ? '<i class="fas fa-circle-notch fa-spin"></i> Creando cuenta...'
        : '<i class="fas fa-user-plus"></i> Crear cuenta';
    };

    registerForm.addEventListener('submit', async event => {
      event.preventDefault();

      if (<?php echo $bootstrapError !== null ? 'true' : 'false'; ?>) {
        return;
      }

      const fullName = document.getElementById('registerFullName').value;
      const email = document.getElementById('registerEmail').value;
      const password = document.getElementById('registerPassword').value;
      const confirmPassword = document.getElementById('registerPasswordConfirm').value;

      setRegisterStatus('<span class="azure-loading-spinner"><i class="fas fa-spinner"></i></span><span>Validando y creando tu cuenta...</span>', 'loading');
      setRegisterLoadingState(true);

      try {
        const response = await fetch('../api/auth.php?action=register', {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            full_name: fullName,
            email,
            password,
            confirm_password: confirmPassword
          })
        });

        let data = null;
        try {
          data = await response.json();
        } catch (_) {
          data = null;
        }

        if (!response.ok || !data || data.ok === false) {
          const serverMessage = data && data.message ? data.message : `Error HTTP ${response.status}.`;
          setRegisterStatus(`<i class="fas fa-circle-xmark"></i><span>${serverMessage}</span>`, 'error');
          setRegisterLoadingState(false);
          return;
        }

        setRegisterStatus('<i class="fas fa-circle-check"></i><span>Registro exitoso. Te vamos a redirigir al login para iniciar sesion.</span><div class="loading-progress-bar"><div class="loading-progress-fill"></div></div>', 'success');

        setTimeout(() => {
          window.location.href = data.redirect_to || 'login.php';
        }, 1800);
      } catch (error) {
        setRegisterStatus(`<i class="fas fa-circle-xmark"></i><span>No se pudo conectar al backend de autenticacion. ${error.message}</span>`, 'error');
        setRegisterLoadingState(false);
      }
    });
  </script>
</body>
</html>
