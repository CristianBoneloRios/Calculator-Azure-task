<?php

declare(strict_types=1);

require_once __DIR__ . '/db.php';

function ensureApplicationInstalled(): void
{
    static $isReady = false;

    if ($isReady) {
        return;
    }

    $pdo = db();
    if (!tableExists($pdo, 'users')) {
        installSchema($pdo, __DIR__ . '/schema.sql');
    }

    seedDefaultAdmin($pdo);
    seedPublicProfile($pdo);
    $isReady = true;
}

function tableExists(PDO $pdo, string $tableName): bool
{
    $sql = 'SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = :table_name';
    $stmt = $pdo->prepare($sql);
    $stmt->execute(['table_name' => $tableName]);

    return (int) $stmt->fetchColumn() > 0;
}

function installSchema(PDO $pdo, string $schemaPath): void
{
    $sql = file_get_contents($schemaPath);
    if (!is_string($sql) || trim($sql) === '') {
        throw new RuntimeException('No se pudo leer el esquema SQL.');
    }

    $sql = preg_replace('/^\s*--.*$/m', '', $sql);
    $statements = preg_split('/;\s*(?:\r?\n|$)/', (string) $sql);

    if (!is_array($statements)) {
        return;
    }

    foreach ($statements as $statement) {
        $statement = trim($statement);
        if ($statement === '') {
            continue;
        }
        $pdo->exec($statement);
    }
}

function seedDefaultAdmin(PDO $pdo): void
{
    $count = (int) $pdo->query('SELECT COUNT(*) FROM users')->fetchColumn();
    if ($count > 0) {
        return;
    }

    $stmt = $pdo->prepare('INSERT INTO users (full_name, email, password_hash, role, last_login_at, last_seen_at) VALUES (:full_name, :email, :password_hash, :role, NULL, NULL)');
    $stmt->execute([
        'full_name' => env('APP_DEFAULT_ADMIN_NAME', 'Administrador'),
        'email' => env('APP_DEFAULT_ADMIN_EMAIL', 'admin@localhost'),
        'password_hash' => password_hash(env('APP_DEFAULT_ADMIN_PASSWORD', 'ChangeMe123!'), PASSWORD_DEFAULT),
        'role' => 'admin',
    ]);
}

function seedPublicProfile(PDO $pdo): void
{
    $slug = env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo');
    $stmt = $pdo->prepare('SELECT id FROM public_profiles WHERE slug = :slug LIMIT 1');
    $stmt->execute(['slug' => $slug]);
    $existing = $stmt->fetch();

    if ($existing) {
        return;
    }

    $adminId = (int) $pdo->query('SELECT id FROM users ORDER BY id ASC LIMIT 1')->fetchColumn();
    $insert = $pdo->prepare('INSERT INTO public_profiles (slug, display_name, role_title, company_name, bio, updated_by_user_id) VALUES (:slug, :display_name, :role_title, :company_name, :bio, :updated_by_user_id)');
    $insert->execute([
        'slug' => $slug,
        'display_name' => env('PUBLIC_PROFILE_NAME', 'Cristian Jesus Bonelo Rios'),
        'role_title' => env('PUBLIC_PROFILE_ROLE', 'Software Quality Analyst'),
        'company_name' => env('PUBLIC_PROFILE_COMPANY', 'Olimpia IT'),
        'bio' => 'Perfil publico editable desde el workspace para centralizar la foto y datos visibles en la portada del sistema.',
        'updated_by_user_id' => $adminId > 0 ? $adminId : null,
    ]);
}

function authenticateUser(string $email, string $password): ?array
{
    $stmt = db()->prepare('SELECT * FROM users WHERE email = :email LIMIT 1');
    $stmt->execute(['email' => strtolower(trim($email))]);
    $user = $stmt->fetch();

    if (!$user || !password_verify($password, (string) $user['password_hash'])) {
        return null;
    }

    return $user;
}

function sanitizeUser(array $user): array
{
    return [
        'id' => (int) $user['id'],
        'full_name' => (string) $user['full_name'],
        'email' => (string) $user['email'],
        'role' => (string) $user['role'],
        'profile_photo_path' => $user['profile_photo_path'] ? (string) $user['profile_photo_path'] : null,
        'last_login_at' => $user['last_login_at'] ? (string) $user['last_login_at'] : null,
        'last_seen_at' => $user['last_seen_at'] ? (string) $user['last_seen_at'] : null,
    ];
}

function startAuthenticatedSession(array $user): array
{
    session_regenerate_id(true);

    $token = bin2hex(random_bytes(32));
    $now = date('Y-m-d H:i:s');

    $_SESSION['user_id'] = (int) $user['id'];
    $_SESSION['workspace_session_token'] = $token;

    $updateUser = db()->prepare('UPDATE users SET last_login_at = :now, last_seen_at = :now WHERE id = :id');
    $updateUser->execute([
        'now' => $now,
        'id' => (int) $user['id'],
    ]);

    $insertSession = db()->prepare('INSERT INTO user_sessions (user_id, session_token, login_at, last_seen_at, ip_address, user_agent, is_active) VALUES (:user_id, :session_token, :login_at, :last_seen_at, :ip_address, :user_agent, 1)');
    $insertSession->execute([
        'user_id' => (int) $user['id'],
        'session_token' => $token,
        'login_at' => $now,
        'last_seen_at' => $now,
        'ip_address' => $_SERVER['REMOTE_ADDR'] ?? null,
        'user_agent' => substr((string) ($_SERVER['HTTP_USER_AGENT'] ?? ''), 0, 255),
    ]);

    return currentUser() ?? sanitizeUser($user);
}

function currentUser(): ?array
{
    $userId = (int) ($_SESSION['user_id'] ?? 0);
    if ($userId <= 0) {
        return null;
    }

    $stmt = db()->prepare('SELECT * FROM users WHERE id = :id LIMIT 1');
    $stmt->execute(['id' => $userId]);
    $user = $stmt->fetch();

    if (!$user) {
        return null;
    }

    $now = date('Y-m-d H:i:s');
    db()->prepare('UPDATE users SET last_seen_at = :now WHERE id = :id')->execute([
        'now' => $now,
        'id' => $userId,
    ]);

    $token = (string) ($_SESSION['workspace_session_token'] ?? '');
    if ($token !== '') {
        db()->prepare('UPDATE user_sessions SET last_seen_at = :now, is_active = 1 WHERE session_token = :session_token')->execute([
            'now' => $now,
            'session_token' => $token,
        ]);
    }

    $user['last_seen_at'] = $now;
    return sanitizeUser($user);
}

function requireApiAuth(): array
{
    $user = currentUser();
    if ($user === null) {
        jsonResponse([
            'ok' => false,
            'message' => 'Sesion requerida.'
        ], 401);
    }

    return $user;
}

function requirePageAuth(): array
{
    $user = currentUser();
    if ($user !== null) {
        return $user;
    }

    header('Location: login.php');
    exit;
}

function logoutCurrentUser(): void
{
    $token = (string) ($_SESSION['workspace_session_token'] ?? '');
    if ($token !== '') {
        db()->prepare('UPDATE user_sessions SET logout_at = :logout_at, is_active = 0 WHERE session_token = :session_token')->execute([
            'logout_at' => date('Y-m-d H:i:s'),
            'session_token' => $token,
        ]);
    }

    $_SESSION = [];
    if (ini_get('session.use_cookies')) {
        $params = session_get_cookie_params();
        setcookie(session_name(), '', time() - 42000, $params['path'], $params['domain'], (bool) $params['secure'], (bool) $params['httponly']);
    }

    session_destroy();
}

function getPublicProfile(string $slug): ?array
{
    $stmt = db()->prepare('SELECT * FROM public_profiles WHERE slug = :slug LIMIT 1');
    $stmt->execute(['slug' => $slug]);
    $profile = $stmt->fetch();

    if (!$profile) {
        return null;
    }

    return [
        'slug' => (string) $profile['slug'],
        'display_name' => (string) $profile['display_name'],
        'role_title' => (string) $profile['role_title'],
        'company_name' => $profile['company_name'] ? (string) $profile['company_name'] : null,
        'bio' => $profile['bio'] ? (string) $profile['bio'] : null,
        'photo_url' => $profile['photo_path'] ? buildAssetUrl((string) $profile['photo_path']) : null,
        'photo_path' => $profile['photo_path'] ? (string) $profile['photo_path'] : null,
    ];
}

function buildAssetUrl(string $relativePath): string
{
    $baseUrl = rtrim((string) env('APP_URL', ''), '/');
    $normalizedPath = ltrim($relativePath, '/');

    if ($baseUrl !== '') {
      return $baseUrl . '/' . $normalizedPath;
    }

    return $normalizedPath;
}

function saveProfilePhoto(array $user, array $file, bool $alsoUpdatePublicProfile = false): string
{
    if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
        throw new RuntimeException('No se pudo subir la imagen.');
    }

    $tmpPath = (string) ($file['tmp_name'] ?? '');
    if ($tmpPath === '' || !is_uploaded_file($tmpPath)) {
        throw new RuntimeException('Archivo temporal invalido.');
    }

    $mimeType = mime_content_type($tmpPath) ?: 'application/octet-stream';
    if (!str_starts_with($mimeType, 'image/')) {
        throw new RuntimeException('Solo se permiten imagenes.');
    }

    $extension = pathinfo((string) ($file['name'] ?? 'photo.png'), PATHINFO_EXTENSION);
    $extension = $extension !== '' ? strtolower($extension) : 'png';

    $relativeDir = 'uploads/profile-photos';
    $absoluteDir = APP_ROOT . '/' . $relativeDir;
    if (!is_dir($absoluteDir) && !mkdir($absoluteDir, 0775, true) && !is_dir($absoluteDir)) {
        throw new RuntimeException('No se pudo crear la carpeta de fotos.');
    }

    $filename = sprintf('user-%d-%s.%s', (int) $user['id'], date('YmdHis'), $extension);
    $relativePath = $relativeDir . '/' . $filename;
    $absolutePath = APP_ROOT . '/' . $relativePath;

    if (!move_uploaded_file($tmpPath, $absolutePath)) {
        throw new RuntimeException('No se pudo guardar la foto en el servidor.');
    }

    db()->prepare('UPDATE users SET profile_photo_path = :profile_photo_path WHERE id = :id')->execute([
        'profile_photo_path' => $relativePath,
        'id' => (int) $user['id'],
    ]);

    db()->prepare('INSERT INTO profile_photo_changes (user_id, changed_by_user_id, file_path, original_name, mime_type, file_size) VALUES (:user_id, :changed_by_user_id, :file_path, :original_name, :mime_type, :file_size)')->execute([
        'user_id' => (int) $user['id'],
        'changed_by_user_id' => (int) $user['id'],
        'file_path' => $relativePath,
        'original_name' => substr((string) ($file['name'] ?? 'photo'), 0, 255),
        'mime_type' => $mimeType,
        'file_size' => isset($file['size']) ? (int) $file['size'] : null,
    ]);

    if ($alsoUpdatePublicProfile) {
        $slug = env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo');
        db()->prepare('UPDATE public_profiles SET photo_path = :photo_path, updated_by_user_id = :updated_by_user_id WHERE slug = :slug')->execute([
            'photo_path' => $relativePath,
            'updated_by_user_id' => (int) $user['id'],
            'slug' => $slug,
        ]);
    }

    return $relativePath;
}