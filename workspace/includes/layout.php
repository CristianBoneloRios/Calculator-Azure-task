<?php

declare(strict_types=1);

require_once dirname(__DIR__, 2) . '/api/app.php';

function workspaceLayoutStart(string $title, string $activePage, array $user): void
{
    $photo = $user['profile_photo_path'] ? '../' . ltrim((string) $user['profile_photo_path'], '/') : null;
  $developerProfile = getDeveloperIdentityProfile();
  $developerPhoto = !empty($developerProfile['photo_path']) ? '../' . ltrim((string) $developerProfile['photo_path'], '/') : null;
    ?>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title><?php echo htmlspecialchars($title, ENT_QUOTES, 'UTF-8'); ?></title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
  <link rel="stylesheet" href="https://cdn-uicons.flaticon.com/uicons-bold-rounded/css/uicons-bold-rounded.css" />
  <link rel="stylesheet" href="../styles.css">
  <link rel="stylesheet" href="assets/workspace.css">
  <?php if ($activePage === 'profile'): ?>
    <link rel="stylesheet" href="assets/profile.css">
  <?php endif; ?>
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
            <?php echo workspaceNavLink('Generacion documentos', 'generacion_documentos.php', 'generacion_documentos', $activePage, 'fa-file-lines'); ?>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Sesion</span>
            <div class="about-me-btn" style="cursor:default;margin-bottom:8px;">
              <?php if ($developerPhoto): ?>
                <div class="about-me-btn-avatar">
                  <img src="<?php echo htmlspecialchars($developerPhoto, ENT_QUOTES, 'UTF-8'); ?>" alt="Foto del desarrollador" />
                </div>
              <?php else: ?>
                <div class="about-me-btn-avatar">
                  <i class="fas fa-user"></i>
                </div>
              <?php endif; ?>
              <div class="about-me-btn-info">
                <span class="about-me-btn-name"><?php echo htmlspecialchars((string) ($developerProfile['display_name'] ?? 'Desarrollador'), ENT_QUOTES, 'UTF-8'); ?></span>
                <span class="about-me-btn-role" style="color:var(--accent-green);">&#9679; <?php echo htmlspecialchars((string) ($developerProfile['role_label'] ?? 'Developer'), ENT_QUOTES, 'UTF-8'); ?></span>
              </div>
            </div>
            <a class="nav-item" href="logout.php" style="color:var(--accent-red,#ef4444);">
              <span class="nav-icon"><i class="fas fa-right-from-bracket"></i></span>
              <span class="nav-label">Cerrar sesion</span>
            </a>
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

  <div id="workspaceToastContainer"></div>
  <script src="assets/workspace.js"></script>
  <script>
    (function () {
      // Sidebar toggle
      const toggle = document.getElementById('sidebarToggleBtn');
      const sidebar = document.getElementById('sidebar');
      const overlay = document.getElementById('sidebarOverlay');

      if (toggle && sidebar && overlay) {
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
      }

      // Footer photo hover (same as app.js initFooterPhotoHoverPreview)
      if (window.matchMedia('(hover: hover)').matches) {
        const wrap    = document.getElementById('footerPhotoWrap');
        const backdrop = document.getElementById('footerPhotoPreviewBackdrop');
        if (wrap && backdrop) {
          wrap.addEventListener('mouseenter', function () { backdrop.classList.add('active'); });
          wrap.addEventListener('mouseleave', function () { backdrop.classList.remove('active'); });
        }
      }
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