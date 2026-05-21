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
  <title>Iniciar sesion | Workspace</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="assets/workspace.css">
</head>
<body>
  <main class="container py-5">
    <div class="row justify-content-center">
      <div class="col-lg-5">
        <div class="workspace-panel">
          <div class="mb-4">
            <span class="eyebrow">Acceso protegido</span>
            <h1 class="h2 mt-2">Inicia sesion en tu workspace</h1>
            <p class="workspace-muted mb-0">Desde aqui se gestionan perfil, notas, tareas, metas, calendario y futuras integraciones con Teams.</p>
          </div>

          <?php if ($bootstrapError !== null): ?>
            <div class="alert alert-warning" role="alert">
              <strong>El workspace no pudo iniciar en este servidor.</strong><br>
              <?php echo htmlspecialchars($bootstrapError, ENT_QUOTES, 'UTF-8'); ?>
            </div>
          <?php endif; ?>

          <form id="loginForm" class="workspace-form">
            <div>
              <label for="loginEmail" class="form-label">Correo</label>
              <input type="email" class="form-control" id="loginEmail" value="<?php echo htmlspecialchars(function_exists('env') ? (string) env('APP_DEFAULT_ADMIN_EMAIL', 'admin@azuretask.local') : loginEnvFallback('APP_DEFAULT_ADMIN_EMAIL', 'admin@azuretask.local'), ENT_QUOTES, 'UTF-8'); ?>" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
            </div>
            <div>
              <label for="loginPassword" class="form-label">Contrasena</label>
              <input type="password" class="form-control" id="loginPassword" placeholder="Ingresa la contrasena configurada en el entorno" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?> required>
            </div>
            <div class="workspace-actions">
              <button type="submit" class="btn btn-primary" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?>><i class="fas fa-right-to-bracket"></i> Entrar</button>
              <a href="../index.php" class="btn btn-outline-light">Volver</a>
            </div>
          </form>

          <div class="workspace-calendar-banner mt-4">
            <strong>Credenciales iniciales generadas desde el entorno</strong>
            <p class="mb-0 workspace-muted">Recomendado: entrar una vez y luego cambiar la contrasena desde Perfil.</p>
          </div>
        </div>
      </div>
    </div>
  </main>

  <script>
    document.getElementById('loginForm').addEventListener('submit', async event => {
      event.preventDefault();

      if (<?php echo $bootstrapError !== null ? 'true' : 'false'; ?>) {
        return;
      }

      const response = await fetch('../api/auth.php?action=login', {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: document.getElementById('loginEmail').value,
          password: document.getElementById('loginPassword').value
        })
      });

      const data = await response.json();
      if (!response.ok || data.ok === false) {
        alert(data.message || 'No fue posible iniciar sesion.');
        return;
      }

      window.location.href = 'index.php';
    });
  </script>
</body>
</html>