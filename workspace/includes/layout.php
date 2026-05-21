<?php

declare(strict_types=1);

require_once dirname(__DIR__, 2) . '/api/app.php';

function workspaceLayoutStart(string $title, string $activePage, array $user): void
{
    $photo = $user['profile_photo_path'] ? '../' . ltrim((string) $user['profile_photo_path'], '/') : null;
    ?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title><?php echo htmlspecialchars($title, ENT_QUOTES, 'UTF-8'); ?></title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" />
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="https://cdn-uicons.flaticon.com/uicons-bold-rounded/css/uicons-bold-rounded.css" />
  <link rel="stylesheet" href="../styles.css">
  <link rel="stylesheet" href="assets/workspace.css">
</head>
<body data-page="<?php echo htmlspecialchars($activePage, ENT_QUOTES, 'UTF-8'); ?>" class="workspace-page-body">
  <header class="app-header">
    <div class="header-left">
      <button class="sidebar-toggle-btn" id="sidebarToggleBtn" title="Toggle Sidebar">
        <i class="fas fa-bars"></i>
      </button>
      <div class="brand">
        <div class="brand-icon">
          <i class="fi fi-br-rocket-lunch"></i>
        </div>
        <div class="brand-text">
          <span class="brand-name">Azure Task Suite</span>
          <span class="brand-sub">Workspace Activo</span>
        </div>
      </div>
    </div>

    <div class="header-center">
      <div class="header-chips">
        <span class="chip chip-blue"><i class="fas fa-layer-group"></i> Workspace</span>
        <span class="chip chip-green"><i class="fas fa-user-check"></i> Sesion Activa</span>
      </div>
    </div>

    <div class="header-right">
      <div class="header-badge">
        <i class="fas fa-user-circle"></i>
        <span><?php echo htmlspecialchars($user['full_name'], ENT_QUOTES, 'UTF-8'); ?></span>
      </div>
    </div>
  </header>

  <div class="sidebar-overlay" id="sidebarOverlay"></div>

  <div class="app-layout">
    <aside class="sidebar" id="sidebar">
      <div class="sidebar-inner">
        <nav class="sidebar-nav">
          <div class="nav-section">
            <span class="nav-section-label">Analizador Base</span>
            <a href="../index.php#upload-section" class="nav-item">
              <span class="nav-icon"><i class="fas fa-cloud-upload-alt"></i></span>
              <span class="nav-label">Cargar Archivos</span>
            </a>
            <a href="../index.php#results-section" class="nav-item">
              <span class="nav-icon"><i class="fas fa-table"></i></span>
              <span class="nav-label">Resultados</span>
            </a>
            <a href="../index.php#summary-section" class="nav-item">
              <span class="nav-icon"><i class="fas fa-chart-pie"></i></span>
              <span class="nav-label">Estadisticas</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Workspace</span>
            <?php echo workspaceNavLink('Inicio', 'index.php', 'index', $activePage, 'fa-chart-line'); ?>
            <?php echo workspaceNavLink('Perfil', 'profile.php', 'profile', $activePage, 'fa-user'); ?>
            <?php echo workspaceNavLink('Notas Importantes', 'notes.php', 'notes', $activePage, 'fa-note-sticky'); ?>
            <?php echo workspaceNavLink('Tareas del Dia', 'tasks.php', 'tasks', $activePage, 'fa-list-check'); ?>
            <?php echo workspaceNavLink('Metas', 'goals.php', 'goals', $activePage, 'fa-bullseye'); ?>
            <?php echo workspaceNavLink('Calendario', 'calendar.php', 'calendar', $activePage, 'fa-calendar-days'); ?>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Sesion</span>
            <div class="workspace-access-card is-active">
              <div class="workspace-access-header">
                <span class="workspace-status-dot"></span>
                <span class="workspace-status-text">Sesion activa</span>
              </div>
              <p><?php echo htmlspecialchars((string) ($user['last_login_at'] ?? 'Primer ingreso'), ENT_QUOTES, 'UTF-8'); ?></p>
              <a class="btn btn-danger btn-sm workspace-auth-btn" href="logout.php">
                <i class="fas fa-right-from-bracket"></i> Cerrar sesion
              </a>
            </div>
          </div>
        </nav>

        <div class="sidebar-footer-brand">
          <i class="fab fa-microsoft"></i>
          <span>Azure DevOps Compatible</span>
        </div>
      </div>
    </aside>

    <main class="main-content workspace-shared-main" id="mainContent">
      <section class="content-section workspace-top-content">
        <div class="section-header">
          <div class="section-title-group">
            <h2><i class="fas fa-layer-group"></i> <?php echo htmlspecialchars($title, ENT_QUOTES, 'UTF-8'); ?></h2>
            <p>Vista activa del workspace con el mismo header, sidebar y footer global.</p>
          </div>
          <div class="workspace-user-chip">
            <?php if ($photo): ?>
              <img src="<?php echo htmlspecialchars($photo, ENT_QUOTES, 'UTF-8'); ?>" alt="Foto de perfil">
            <?php else: ?>
              <span class="workspace-user-fallback"><i class="fas fa-user"></i></span>
            <?php endif; ?>
            <div>
              <strong><?php echo htmlspecialchars($user['full_name'], ENT_QUOTES, 'UTF-8'); ?></strong>
              <small><?php echo htmlspecialchars($user['email'], ENT_QUOTES, 'UTF-8'); ?></small>
            </div>
          </div>
        </div>

        <div class="workspace-content">
<?php
}

function workspaceLayoutEnd(): void
{
    ?>
        </div>
      </section>
    </main>
  </div>

  <footer class="app-footer">
    <div class="footer-content">
      <div class="footer-left d-flex align-items-center gap-2">
        <img src="https://media.licdn.com/dms/image/v2/D5603AQF_Y5pnaolD5g/profile-displayphoto-scale_400_400/B56ZiucgVJHcAg-/0/1755273365460?e=1781136000&v=beta&t=zxALSfRgKjZz0GeAv-HO7G68ASoWSSwIrKPZiK9TFJA" alt="Foto Cristian Bonelo" class="footer-avatar" width="36" height="36" style="border-radius:50%;object-fit:cover;border:2px solid var(--accent-blue);box-shadow:0 2px 8px #0003;" loading="lazy">
        <div class="d-none d-md-block">
          <span class="footer-powered">Desarrollado por</span>
          <a href="https://www.linkedin.com/in/cristiandevbonelo/" target="_blank" rel="noopener" class="footer-name">Cristian Jesus Bonelo Rios</a>
        </div>
      </div>
      <div class="footer-center d-flex flex-column flex-md-row align-items-center gap-1 gap-md-3">
        <span class="footer-role">SOFTWARE QUALITY ANALYST</span>
        <span class="footer-sep d-none d-md-inline">|</span>
        <span class="footer-dept">DEVELOPMENT &amp; INNOVATION</span>
        <span class="footer-sep d-none d-md-inline">|</span>
        <a href="https://cristiandevbonelo.github.io/porfoliocristian/" target="_blank" rel="noopener" class="footer-link"><i class="fas fa-globe"></i> Portafolio</a>
        <span class="footer-sep d-none d-md-inline">|</span>
        <a href="mailto:cristiandevbonelo@gmail.com" class="footer-link" title="Contáctame"><i class="fas fa-envelope"></i> Contacto</a>
        <span class="footer-sep d-none d-md-inline">|</span>
        <a href="https://www.linkedin.com/in/cristiandevbonelo/" target="_blank" rel="noopener" class="footer-link" title="LinkedIn"><i class="fab fa-linkedin"></i></a>
        <a href="https://github.com/cristiandevbonelo" target="_blank" rel="noopener" class="footer-link" title="GitHub"><i class="fab fa-github"></i></a>
      </div>
      <div class="footer-right d-flex flex-column align-items-end gap-1">
        <span class="footer-version">v1.0.0</span>
        <span class="footer-date" id="footerDate"></span>
        <span class="footer-status"><i class="fas fa-circle" style="color:var(--accent-green);font-size:8px;"></i> Online</span>
      </div>
    </div>
    <div class="footer-bottom text-center mt-1" style="font-size:10px;color:var(--text-3);">
      &copy; <?php echo date('Y'); ?> Cristian Bonelo. Todos los derechos reservados.
    </div>
    <script>
      // Fecha y hora en vivo en el footer
      function updateFooterDate() {
        const el = document.getElementById('footerDate');
        if (!el) return;
        const now = new Date();
        const opts = { year: 'numeric', month: 'short', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
        el.textContent = now.toLocaleString('es-CO', opts) + ' (GMT' + (now.getTimezoneOffset()/-60) + ')';
      }
      setInterval(updateFooterDate, 1000);
      updateFooterDate();
    </script>
  </footer>

  <div class="toast-container position-fixed bottom-0 end-0 p-3" id="workspaceToastContainer"></div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
  <script src="assets/workspace.js"></script>
  <script>
    (function () {
      const toggle = document.getElementById('sidebarToggleBtn');
      const sidebar = document.getElementById('sidebar');
      const overlay = document.getElementById('sidebarOverlay');

      if (!toggle || !sidebar || !overlay) {
        return;
      }

      function closeSidebar() {
        sidebar.classList.remove('mobile-open');
        overlay.classList.remove('visible');
      }

      toggle.addEventListener('click', function () {
        const isMobile = window.innerWidth <= 768;
        if (isMobile) {
          sidebar.classList.toggle('mobile-open');
          overlay.classList.toggle('visible');
        } else {
          sidebar.classList.toggle('collapsed');
        }
      });

      overlay.addEventListener('click', closeSidebar);
    }());
  </script>
</body>
</html>
<?php
}

function workspaceNavLink(string $label, string $href, string $pageKey, string $activePage, string $icon): string
{
  $activeClass = $pageKey === $activePage ? 'active' : '';
    return sprintf(
    '<a class="nav-item %s" href="%s"><span class="nav-icon"><i class="fas %s"></i></span><span class="nav-label">%s</span></a>',
        $activeClass,
        htmlspecialchars($href, ENT_QUOTES, 'UTF-8'),
        htmlspecialchars($icon, ENT_QUOTES, 'UTF-8'),
        htmlspecialchars($label, ENT_QUOTES, 'UTF-8')
    );
}