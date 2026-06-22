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
    <section class="auth-card auth-card-register">
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

        <div class="login-footer-animation">
          <div class="login-icon-circle-footer">
            <i class="fas fa-user-plus"></i>
          </div>
          <div class="login-ripple-footer"></div>
          <div class="login-ripple-footer delay-1"></div>
        </div>
      </div>

      <div class="auth-form-panel auth-form-panel-register">
        <h2>Crear cuenta</h2>
        <p class="auth-muted auth-muted-register">Completa tus datos para continuar.</p>
        <div id="registerStatus" class="auth-status" aria-live="polite" hidden></div>

        <?php if ($bootstrapError !== null): ?>
          <div class="auth-alert" role="alert">
            <strong>La aplicacion no pudo iniciar correctamente en el servidor.</strong><br>
            <?php echo htmlspecialchars($bootstrapError, ENT_QUOTES, 'UTF-8'); ?>
          </div>
        <?php endif; ?>

        <form id="registerForm" class="auth-form auth-form-register">
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
          <div class="norm-toggle-group auth-password-toggle">
            <label class="norm-toggle">
              <input type="checkbox" id="registerShowPassword" <?php echo $bootstrapError !== null ? 'disabled' : ''; ?>>
              <span class="norm-toggle-slider"></span>
            </label>
            <span>Mostrar contrasenas</span>
          </div>
          <div class="auth-password-meter" id="registerPasswordMeter">
            <div class="auth-password-meter-bar">
              <div class="auth-password-meter-fill" id="registerPasswordMeterFill"></div>
            </div>
            <span class="auth-password-meter-label" id="registerPasswordMeterLabel">Fuerza: muy debil</span>
            <ul class="auth-password-checklist" id="registerPasswordChecklist">
              <li class="auth-password-check" data-rule="letter"><i class="fas fa-circle"></i>Debe contener al menos una letra</li>
              <li class="auth-password-check" data-rule="upper"><i class="fas fa-circle"></i>Debe incluir una mayuscula</li>
              <li class="auth-password-check" data-rule="lower"><i class="fas fa-circle"></i>Debe incluir una minuscula</li>
              <li class="auth-password-check" data-rule="number"><i class="fas fa-circle"></i>Debe incluir un numero</li>
              <li class="auth-password-check" data-rule="special"><i class="fas fa-circle"></i>Debe incluir un caracter especial</li>
              <li class="auth-password-check" data-rule="length"><i class="fas fa-circle"></i>Debe tener minimo 8 caracteres</li>
            </ul>
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
        <div class="footer-left">
          <i class="fas fa-code"></i>
        </div>
        <div class="footer-center">
          <span class="footer-powered">Powered by</span>
          <span class="footer-name">Cristian Jesus Bonelo Rios</span>
          <span class="footer-sep">|</span>
          <span class="footer-role">SOFTWARE QUALITY ANALYST</span>
          <span class="footer-sep">|</span>
          <span class="footer-dept">DEVELOPMENT &amp; INNOVATION</span>
        </div>
        <div class="footer-right">
          <span class="footer-version">v1.0.0</span>
        </div>
      </div>

      <div class="footer-section footer-socials">
        <span class="footer-section-title"><i class="fas fa-share-alt"></i> Redes Sociales</span>
        <div class="footer-social-links">
          <a href="https://cristiandevbonelo.github.io/porfoliocristian/" target="_blank" rel="noopener noreferrer" aria-label="Portafolio">
            <i class="fas fa-globe"></i>
            <span>Portafolio</span>
          </a>
          <a href="#" aria-label="LinkedIn pendiente">
            <i class="fab fa-linkedin"></i>
            <span>LinkedIn (pendiente)</span>
          </a>
          <a href="https://github.com/CristianBoneloRios" target="_blank" rel="noopener noreferrer" aria-label="GitHub">
            <i class="fab fa-github"></i>
            <span>GitHub</span>
          </a>
        </div>
      </div>

      <div class="footer-section footer-photo-block">
        <div class="footer-photo-wrap" id="footerPhotoWrap">
          <img class="footer-photo-thumb" src="https://cristiandevbonelo.github.io/porfoliocristian/assets/1755273365367.jpg" alt="Foto de Cristian Bonelo" loading="lazy" />
        </div>
      </div>
    </div>
  </footer>

  <div class="photo-hover-backdrop" id="footerPhotoPreviewBackdrop" aria-hidden="true">
    <img class="photo-hover-preview" id="footerPhotoPreviewImg" src="https://cristiandevbonelo.github.io/porfoliocristian/assets/1755273365367.jpg" alt="Vista ampliada de foto" loading="lazy" />
  </div>

  <script>
    const registerForm = document.getElementById('registerForm');
    const registerSubmitBtn = registerForm.querySelector('button[type="submit"]');
    const registerStatus = document.getElementById('registerStatus');
    const registerPassword = document.getElementById('registerPassword');
    const registerPasswordConfirm = document.getElementById('registerPasswordConfirm');
    const registerShowPassword = document.getElementById('registerShowPassword');
    const registerPasswordMeterFill = document.getElementById('registerPasswordMeterFill');
    const registerPasswordMeterLabel = document.getElementById('registerPasswordMeterLabel');
    const registerPasswordChecklist = document.getElementById('registerPasswordChecklist');
    const footerPhotoWrap = document.getElementById('footerPhotoWrap');
    const footerPhotoPreviewBackdrop = document.getElementById('footerPhotoPreviewBackdrop');

    const passwordRules = {
      letter: value => /[A-Za-z]/.test(value),
      upper: value => /[A-Z]/.test(value),
      lower: value => /[a-z]/.test(value),
      number: value => /\d/.test(value),
      special: value => /[^A-Za-z0-9]/.test(value),
      length: value => value.length >= 8
    };

    const evaluatePassword = value => {
      const results = Object.fromEntries(
        Object.entries(passwordRules).map(([rule, tester]) => [rule, tester(value)])
      );
      const score = Object.values(results).filter(Boolean).length;
      const hasAdvancedLength = value.length >= 12;
      return { results, score, hasAdvancedLength };
    };

    const updateSubmitAvailability = isStrongOrExcellent => {
      if (<?php echo $bootstrapError !== null ? 'true' : 'false'; ?>) {
        registerSubmitBtn.disabled = true;
        return;
      }

      registerSubmitBtn.disabled = !isStrongOrExcellent;
      registerSubmitBtn.title = isStrongOrExcellent
        ? ''
        : 'La contrasena debe estar en nivel fuerte o excelente para habilitar el registro.';
    };

    const renderPasswordStrength = value => {
      const { results, score, hasAdvancedLength } = evaluatePassword(value);
      const maxScore = 7;
      const strengthScore = score + (hasAdvancedLength ? 1 : 0);
      const percent = Math.max(8, Math.min(100, Math.round((strengthScore / maxScore) * 100)));
      const checklistItems = registerPasswordChecklist.querySelectorAll('.auth-password-check');

      checklistItems.forEach(item => {
        const rule = item.dataset.rule;
        item.classList.toggle('met', Boolean(results[rule]));
      });

      let label = 'Fuerza: muy debil';
      let gradient = 'linear-gradient(90deg, #ef4444, #f59e0b)';
      let isStrongOrExcellent = false;

      if (score === 6 && hasAdvancedLength) {
        label = 'Fuerza: excelente';
        gradient = 'linear-gradient(90deg, #22c55e, #10b981)';
        isStrongOrExcellent = true;
      } else if (score === 6) {
        label = 'Fuerza: fuerte';
        gradient = 'linear-gradient(90deg, #84cc16, #22c55e)';
        isStrongOrExcellent = true;
      } else if (score >= 4) {
        label = 'Fuerza: media';
        gradient = 'linear-gradient(90deg, #f59e0b, #eab308)';
      } else if (score >= 2) {
        label = 'Fuerza: debil';
        gradient = 'linear-gradient(90deg, #f97316, #f59e0b)';
      }

      registerPasswordMeterFill.style.width = `${percent}%`;
      registerPasswordMeterFill.style.background = gradient;
      registerPasswordMeterLabel.textContent = label;
      updateSubmitAvailability(isStrongOrExcellent);

      return score === 6;
    };

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

      if (!isLoading) {
        renderPasswordStrength(registerPassword.value);
      }
    };

    registerShowPassword?.addEventListener('change', () => {
      const type = registerShowPassword.checked ? 'text' : 'password';
      registerPassword.type = type;
      registerPasswordConfirm.type = type;
    });

    registerPassword.addEventListener('input', () => {
      renderPasswordStrength(registerPassword.value);
    });

    if (footerPhotoWrap && footerPhotoPreviewBackdrop && window.matchMedia && window.matchMedia('(hover: hover)').matches) {
      let hideTimer = null;

      const showPreview = () => {
        if (hideTimer) {
          clearTimeout(hideTimer);
          hideTimer = null;
        }
        footerPhotoPreviewBackdrop.classList.add('active');
      };

      const hidePreview = () => {
        hideTimer = window.setTimeout(() => {
          footerPhotoPreviewBackdrop.classList.remove('active');
        }, 40);
      };

      footerPhotoWrap.addEventListener('mouseenter', showPreview);
      footerPhotoWrap.addEventListener('mouseleave', hidePreview);
    }

    renderPasswordStrength(registerPassword.value);

    registerForm.addEventListener('submit', async event => {
      event.preventDefault();

      if (<?php echo $bootstrapError !== null ? 'true' : 'false'; ?>) {
        return;
      }

      const fullName = document.getElementById('registerFullName').value;
      const email = document.getElementById('registerEmail').value;
      const password = registerPassword.value;
      const confirmPassword = registerPasswordConfirm.value;

      if (!renderPasswordStrength(password)) {
        setRegisterStatus('<i class="fas fa-circle-xmark"></i><span>La contrasena debe incluir letra, mayuscula, minuscula, numero y caracter especial (minimo 8 caracteres).</span>', 'error');
        return;
      }

      if (password !== confirmPassword) {
        setRegisterStatus('<i class="fas fa-circle-xmark"></i><span>Las contrasenas no coinciden.</span>', 'error');
        return;
      }

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
          const loginTarget = data.redirect_to || 'login.php';
          const separator = loginTarget.includes('?') ? '&' : '?';
          window.location.href = `${loginTarget}${separator}email=${encodeURIComponent(email.trim())}`;
        }, 1800);
      } catch (error) {
        setRegisterStatus(`<i class="fas fa-circle-xmark"></i><span>No se pudo conectar al backend de autenticacion. ${error.message}</span>`, 'error');
        setRegisterLoadingState(false);
      }
    });
  </script>
</body>
</html>
