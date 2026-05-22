<?php

declare(strict_types=1);

$bootstrapError = null;

function loginEnvFallback(string $key, string $default): string
{
    $value = getenv($key);
    if ($value === false || $value === null || trim((string) $value) === '') {
        return $default;
    }

    return (string) $value;
}

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
  <title>Iniciar sesion | Azure Task Suite</title>
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
          <span class="brand-sub">Acceso al Workspace</span>
        </div>
      </div>
    </div>
    <div class="header-center">
      <div class="header-chips">
        <span class="chip chip-blue"><i class="fas fa-shield"></i> Seguro</span>
        <span class="chip chip-green"><i class="fas fa-user-check"></i> Sesion</span>
      </div>
    </div>
    <div class="header-right">
      <a href="register.php" class="btn btn-primary">
        <i class="fas fa-user-plus"></i> Registrarme
      </a>
    </div>
  </header>

  <main class="auth-main">
    <section class="auth-card">
      <div class="auth-aside">
        <span class="auth-kicker">Acceso protegido</span>
        <h1>Inicia sesion en tu Workspace</h1>
        <p>Gestiona calendario, tareas, metas y notas en la misma experiencia de Azure Task Suite.</p>
        <ul class="auth-list">
          <li><i class="fas fa-calendar-days"></i> Calendario integrado</li>
          <li><i class="fas fa-list-check"></i> Gestion de tareas</li>
          <li><i class="fas fa-bullseye"></i> Seguimiento de metas</li>
          <li><i class="fas fa-note-sticky"></i> Notas importantes</li>
        </ul>
      </div>

      <div class="auth-form-panel">
        <h2>Bienvenido de nuevo</h2>
        <p class="auth-muted">Ingresa con tu correo y contrasena.</p>

        <?php if ($bootstrapError !== null): ?>
          <div class="auth-alert" role="alert">
            <strong>La aplicacion no pudo iniciar correctamente en el servidor.</strong><br>
            <?php echo htmlspecialchars($bootstrapError, ENT_QUOTES, 'UTF-8'); ?>
          </div>
        <?php endif; ?>

        <form id="loginForm" class="auth-form">
          <div>
            <label for="loginEmail">Correo</label>
            <input type="email" id="loginEmail" value="<?php echo htmlspecialchars(function_exists('env') ? (string) env('APP_DEFAULT_ADMIN_EMAIL', 'admin@azuretask.local') : loginEnvFallback('APP_DEFAULT_ADMIN_EMAIL', 'admin@azuretask.local'), ENT_QUOTES, 'UTF-8'); ?>" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
          </div>
          <div>
            <label for="loginPassword">Contrasena</label>
            <input type="password" id="loginPassword" placeholder="Ingresa tu contrasena" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
          </div>
          <div class="auth-actions">
            <button type="submit" class="btn btn-primary" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?>><i class="fas fa-right-to-bracket"></i> Entrar</button>
            <a href="../index.php" class="btn btn-outline">Volver</a>
          </div>
        </form>

        <div class="auth-switch">
          <span>No tienes cuenta?</span>
          <a href="register.php">Registrate aqui</a>
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
    document.getElementById('loginForm').addEventListener('submit', async event => {
      event.preventDefault();

      if (<?php echo $bootstrapError !== null ? 'true' : 'false'; ?>) {
        return;
      }

      try {
        const response = await fetch('../api/auth.php?action=login', {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email: document.getElementById('loginEmail').value,
            password: document.getElementById('loginPassword').value
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
          alert(`${serverMessage}\n\nSi el error persiste, valida configuracion PHP y base de datos en Hostinger.`);
          return;
        }

        // Verificar si requiere 2FA
        if (data.requires_2fa) {
          window.location.href = 'verify-2fa.php';
        } else {
          window.location.href = 'index.php';
        }
      } catch (error) {
        alert(`No se pudo conectar al backend de autenticacion.\n\nDetalle: ${error.message}`);
      }
    });
  </script>
</body>
</html>
