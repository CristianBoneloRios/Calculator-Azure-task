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
    if (shouldInstallSchema($pdo)) {
        installSchema($pdo, __DIR__ . '/schema.sql');
    }

    runColumnMigrations($pdo);
    ensureDocumentGenerationSchema($pdo);
    ensureDeveloperIdentitySchema($pdo);
    ensureNotesCollaborationSchema($pdo);
    seedDefaultAdmin($pdo);
    seedPublicProfile($pdo);
    $isReady = true;
}

function ensureDocumentGenerationSchema(PDO $pdo): void
{
    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS document_generation_jobs (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id BIGINT UNSIGNED NOT NULL,
            user_email VARCHAR(255) NOT NULL,
            input_file_name VARCHAR(255) DEFAULT NULL,
            input_file_type VARCHAR(50) DEFAULT NULL,
            input_file_size BIGINT DEFAULT NULL,
            input_file_path TEXT DEFAULT NULL,
            input_file_hash VARCHAR(255) DEFAULT NULL,
            input_base64 LONGTEXT DEFAULT NULL,
            generation_type VARCHAR(50) DEFAULT NULL,
            description TEXT DEFAULT NULL,
            pa_request_id VARCHAR(255) DEFAULT NULL,
            pa_status VARCHAR(50) DEFAULT NULL,
            pa_response TEXT DEFAULT NULL,
            output_file_name VARCHAR(255) DEFAULT NULL,
            output_file_type VARCHAR(50) DEFAULT NULL,
            output_file_path TEXT DEFAULT NULL,
            output_file_size BIGINT DEFAULT NULL,
            output_file_url TEXT DEFAULT NULL,
            ai_summary TEXT DEFAULT NULL,
            ai_confidence DECIMAL(5,2) DEFAULT NULL,
            status ENUM("pending","processing","completed","error") NOT NULL DEFAULT "pending",
            error_message TEXT DEFAULT NULL,
            ip_address VARCHAR(45) DEFAULT NULL,
            user_agent TEXT DEFAULT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            completed_at TIMESTAMP NULL DEFAULT NULL,
            INDEX idx_user (user_id),
            INDEX idx_status (status),
            INDEX idx_created_at (created_at)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );

    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS user_document_security (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id BIGINT UNSIGNED NOT NULL UNIQUE,
            access_key_hash VARCHAR(255) NOT NULL,
            is_enabled TINYINT(1) NOT NULL DEFAULT 1,
            last_verified_at TIMESTAMP NULL DEFAULT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            INDEX idx_user_doc_security_user (user_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );
}

/**
 * Safely add any columns that may be missing from tables created before these
 * columns were added to schema.sql.  Uses information_schema so it is safe to
 * run on every request – the query is fast and the ALTER is only executed when
 * the column is actually absent.
 */
function runColumnMigrations(PDO $pdo): void
{
    $migrations = [
        // users – 2FA columns
        ['users', 'two_factor_secret',  'VARCHAR(255) DEFAULT NULL   AFTER password_hash'],
        ['users', 'two_factor_enabled', 'TINYINT(1)  NOT NULL DEFAULT 0 AFTER two_factor_secret'],
        ['users', 'profile_photo_path', 'VARCHAR(255) DEFAULT NULL   AFTER role'],
        ['users', 'last_login_at',      'DATETIME DEFAULT NULL        AFTER profile_photo_path'],
        ['users', 'last_seen_at',       'DATETIME DEFAULT NULL        AFTER last_login_at'],
        // integrations – outbound URL stored in metadata (JSON), no new column needed
        // calendar_events – no missing columns expected
    ];

    foreach ($migrations as [$table, $column, $definition]) {
        if (!columnExists($pdo, $table, $column)) {
            try {
                $pdo->exec("ALTER TABLE `{$table}` ADD COLUMN `{$column}` {$definition}");
            } catch (Throwable $e) {
                error_log("Migration failed [{$table}.{$column}]: " . $e->getMessage());
            }
        }
    }
}

function columnExists(PDO $pdo, string $table, string $column): bool
{
    try {
        $stmt = $pdo->prepare(
            'SELECT COUNT(*) FROM information_schema.columns
             WHERE table_schema = DATABASE()
               AND table_name   = :table
               AND column_name  = :column'
        );
        $stmt->execute(['table' => $table, 'column' => $column]);
        return (int) $stmt->fetchColumn() > 0;
    } catch (Throwable $throwable) {
        try {
            $stmt = $pdo->query('SHOW COLUMNS FROM `' . str_replace('`', '``', $table) . '`');
            $columns = $stmt ? $stmt->fetchAll(PDO::FETCH_ASSOC) : [];
            foreach ($columns as $columnMeta) {
                if (isset($columnMeta['Field']) && (string) $columnMeta['Field'] === $column) {
                    return true;
                }
            }
            return false;
        } catch (Throwable $fallbackThrowable) {
            throw $throwable;
        }
    }
}

function ensureTwoFactorSchemaReady(PDO $pdo): void
{
    $requiredColumns = [
        ['users', 'two_factor_secret'],
        ['users', 'two_factor_enabled'],
    ];

    $missingColumns = [];
    foreach ($requiredColumns as [$table, $column]) {
        if (!columnExists($pdo, $table, $column)) {
            $missingColumns[] = $table . '.' . $column;
        }
    }

    if ($missingColumns === []) {
        ensureTwoFactorRecoveryTable($pdo);
        return;
    }

    // Retry migrations once in case schema changed after a deployment.
    runColumnMigrations($pdo);

    $stillMissing = [];
    foreach ($requiredColumns as [$table, $column]) {
        if (!columnExists($pdo, $table, $column)) {
            $stillMissing[] = $table . '.' . $column;
        }
    }

    if ($stillMissing !== []) {
        throw new RuntimeException(
            'La base de datos no tiene las columnas requeridas para 2FA: ' . implode(', ', $stillMissing) . '. Ejecuta la migracion del esquema (api/schema.sql) en Hostinger y vuelve a intentar.'
        );
    }

    ensureTwoFactorRecoveryTable($pdo);
}

function ensureTwoFactorRecoveryTable(PDO $pdo): void
{
    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS two_factor_recovery_codes (
            id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id INT UNSIGNED NOT NULL,
            code_hash VARCHAR(255) NOT NULL,
            used_at DATETIME DEFAULT NULL,
            is_valid TINYINT(1) NOT NULL DEFAULT 1,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            KEY idx_user_id (user_id),
            KEY idx_used_at (used_at),
            CONSTRAINT fk_two_factor_recovery_codes_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );
}

function shouldInstallSchema(PDO $pdo): bool
{
    $requiredTables = [
        'users',
        'user_sessions',
        'public_profiles',
        'profile_photo_changes',
        'notes',
        'daily_tasks',
        'goals',
        'calendar_sources',
        'calendar_events',
        'integrations',
        'app_settings',
    ];

    foreach ($requiredTables as $tableName) {
        if (!tableExists($pdo, $tableName)) {
            return true;
        }
    }

    return false;
}

function tableExists(PDO $pdo, string $tableName): bool
{
    try {
        $sql = 'SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = :table_name';
        $stmt = $pdo->prepare($sql);
        $stmt->execute(['table_name' => $tableName]);

        return (int) $stmt->fetchColumn() > 0;
    } catch (Throwable $throwable) {
        try {
            $stmt = $pdo->prepare('SHOW TABLES LIKE :table_name');
            $stmt->execute(['table_name' => $tableName]);
            return (bool) $stmt->fetchColumn();
        } catch (Throwable $fallbackThrowable) {
            throw $throwable;
        }
    }
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

function registerUserAccount(string $fullName, string $email, string $password): array
{
    $cleanName = trim($fullName);
    $cleanEmail = strtolower(trim($email));

    if ($cleanName === '' || mb_strlen($cleanName) < 3) {
        throw new InvalidArgumentException('El nombre debe tener al menos 3 caracteres.');
    }

    if (!filter_var($cleanEmail, FILTER_VALIDATE_EMAIL)) {
        throw new InvalidArgumentException('Correo invalido.');
    }

    if (strlen($password) < 8) {
        throw new InvalidArgumentException('La contrasena debe tener al menos 8 caracteres.');
    }

    $hasLetter = preg_match('/[A-Za-z]/', $password) === 1;
    $hasUpper = preg_match('/[A-Z]/', $password) === 1;
    $hasLower = preg_match('/[a-z]/', $password) === 1;
    $hasNumber = preg_match('/\d/', $password) === 1;
    $hasSpecial = preg_match('/[^A-Za-z0-9]/', $password) === 1;

    if (!$hasLetter || !$hasUpper || !$hasLower || !$hasNumber || !$hasSpecial) {
        throw new InvalidArgumentException('La contrasena debe incluir letra, mayuscula, minuscula, numero y caracter especial.');
    }

    $stmt = db()->prepare('INSERT INTO users (full_name, email, password_hash, role, last_login_at, last_seen_at) VALUES (:full_name, :email, :password_hash, :role, NULL, NULL)');

    try {
        $stmt->execute([
            'full_name' => $cleanName,
            'email' => $cleanEmail,
            'password_hash' => password_hash($password, PASSWORD_DEFAULT),
            'role' => 'member',
        ]);
    } catch (PDOException $exception) {
        $sqlState = (string) $exception->getCode();
        if ($sqlState === '23000') {
            throw new RuntimeException('Ya existe una cuenta con ese correo.');
        }
        throw $exception;
    }

    $userId = (int) db()->lastInsertId();
    $userStmt = db()->prepare('SELECT * FROM users WHERE id = :id LIMIT 1');
    $userStmt->execute(['id' => $userId]);
    $user = $userStmt->fetch();

    if (!$user) {
        throw new RuntimeException('No se pudo recuperar el usuario registrado.');
    }

    return $user;
}

function generateTwoFactorSecret(): string
{
    return encodeBase32(random_bytes(20));
}

function verifyTwoFactorCode(string $secret, string $code): bool
{
    if (strlen($code) !== 6 || !ctype_digit($code)) {
        return false;
    }

    $secretKey = decodeBase32($secret);
    if ($secretKey === null || $secretKey === '') {
        return false;
    }

    // TOTP: Time-based One-Time Password
    $time = (int) floor(time() / 30);
    $hmacs = [];

    for ($i = -1; $i <= 1; $i++) {
        $timestamp = pack('N2', 0, $time + $i);
        $hmac = hash_hmac('sha1', $timestamp, $secretKey, true);
        $offset = ord($hmac[19]) & 0xf;
        $code_int = (ord($hmac[$offset]) & 0x7f) << 24 |
                    (ord($hmac[$offset + 1]) & 0xff) << 16 |
                    (ord($hmac[$offset + 2]) & 0xff) << 8 |
                    (ord($hmac[$offset + 3]) & 0xff);
        $hmacs[] = str_pad($code_int % 1000000, 6, '0', STR_PAD_LEFT);
    }

    return in_array($code, $hmacs, true);
}

function generateTwoFactorRecoveryCodes(int $count = 8): array
{
    $codes = [];
    $alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

    while (count($codes) < $count) {
        $raw = '';
        for ($i = 0; $i < 8; $i++) {
            $raw .= $alphabet[random_int(0, strlen($alphabet) - 1)];
        }

        $formatted = substr($raw, 0, 4) . '-' . substr($raw, 4, 4);
        $codes[$formatted] = true;
    }

    return array_keys($codes);
}

function storeTwoFactorRecoveryCodes(PDO $pdo, int $userId, array $plainCodes): void
{
    $pdo->prepare('DELETE FROM two_factor_recovery_codes WHERE user_id = :user_id')
        ->execute(['user_id' => $userId]);

    $insert = $pdo->prepare(
        'INSERT INTO two_factor_recovery_codes (user_id, code_hash, used_at, is_valid)
         VALUES (:user_id, :code_hash, NULL, 1)'
    );

    foreach ($plainCodes as $plainCode) {
        $insert->execute([
            'user_id' => $userId,
            'code_hash' => password_hash(normalizeTwoFactorRecoveryCode((string) $plainCode), PASSWORD_DEFAULT),
        ]);
    }
}

function consumeTwoFactorRecoveryCode(PDO $pdo, int $userId, string $candidateCode): bool
{
    $normalizedCandidate = normalizeTwoFactorRecoveryCode($candidateCode);
    if ($normalizedCandidate === '') {
        return false;
    }

    $stmt = $pdo->prepare(
        'SELECT id, code_hash
         FROM two_factor_recovery_codes
         WHERE user_id = :user_id
           AND is_valid = 1
           AND used_at IS NULL
         ORDER BY id ASC'
    );
    $stmt->execute(['user_id' => $userId]);
    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

    if (!is_array($rows) || $rows === []) {
        return false;
    }

    foreach ($rows as $row) {
        $hash = (string) ($row['code_hash'] ?? '');
        if ($hash === '' || !password_verify($normalizedCandidate, $hash)) {
            continue;
        }

        $update = $pdo->prepare(
            'UPDATE two_factor_recovery_codes
             SET used_at = :used_at, is_valid = 0
             WHERE id = :id AND is_valid = 1 AND used_at IS NULL'
        );
        $update->execute([
            'used_at' => date('Y-m-d H:i:s'),
            'id' => (int) $row['id'],
        ]);

        return $update->rowCount() > 0;
    }

    return false;
}

function normalizeTwoFactorRecoveryCode(string $value): string
{
    return strtoupper(preg_replace('/[^A-Z0-9]/i', '', $value) ?? '');
}

function encodeBase32(string $binary): string
{
    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    $bits = '';
    $length = strlen($binary);

    for ($index = 0; $index < $length; $index++) {
        $bits .= str_pad(decbin(ord($binary[$index])), 8, '0', STR_PAD_LEFT);
    }

    $output = '';
    $bitLength = strlen($bits);
    for ($offset = 0; $offset < $bitLength; $offset += 5) {
        $chunk = substr($bits, $offset, 5);
        if ($chunk === '') {
            continue;
        }

        $chunk = str_pad($chunk, 5, '0', STR_PAD_RIGHT);
        $output .= $alphabet[bindec($chunk)];
    }

    return $output;
}

function decodeBase32(string $encoded): ?string
{
    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    $clean = strtoupper(preg_replace('/[^A-Z2-7]/', '', $encoded) ?? '');

    if ($clean === '') {
        return null;
    }

    $bits = '';
    $length = strlen($clean);
    for ($index = 0; $index < $length; $index++) {
        $position = strpos($alphabet, $clean[$index]);
        if ($position === false) {
            return null;
        }

        $bits .= str_pad(decbin($position), 5, '0', STR_PAD_LEFT);
    }

    $output = '';
    $bitLength = strlen($bits);
    for ($offset = 0; $offset + 8 <= $bitLength; $offset += 8) {
        $output .= chr(bindec(substr($bits, $offset, 8)));
    }

    return $output;
}

function sanitizeUser(array $user): array
{
    return [
        'id' => (int) $user['id'],
        'full_name' => (string) $user['full_name'],
        'email' => (string) $user['email'],
        'role' => (string) $user['role'],
        'two_factor_enabled' => !empty($user['two_factor_enabled']),
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

    try {
        $updateUser = db()->prepare('UPDATE users SET last_login_at = :now, last_seen_at = :now WHERE id = :id');
        $updateUser->execute([
            'now' => $now,
            'id' => (int) $user['id'],
        ]);
    } catch (Throwable $throwable) {
        error_log('No se pudo actualizar last_login_at/last_seen_at: ' . $throwable->getMessage());
    }

    try {
        $insertSession = db()->prepare('INSERT INTO user_sessions (user_id, session_token, login_at, last_seen_at, ip_address, user_agent, is_active) VALUES (:user_id, :session_token, :login_at, :last_seen_at, :ip_address, :user_agent, 1)');
        $insertSession->execute([
            'user_id' => (int) $user['id'],
            'session_token' => $token,
            'login_at' => $now,
            'last_seen_at' => $now,
            'ip_address' => $_SERVER['REMOTE_ADDR'] ?? null,
            'user_agent' => substr((string) ($_SERVER['HTTP_USER_AGENT'] ?? ''), 0, 255),
        ]);
    } catch (Throwable $throwable) {
        // This telemetry record should not block successful authentication.
        error_log('No se pudo registrar user_session: ' . $throwable->getMessage());
    }

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

    try {
        db()->prepare('UPDATE users SET last_seen_at = :now WHERE id = :id')->execute([
            'now' => $now,
            'id' => $userId,
        ]);
    } catch (Throwable $throwable) {
        error_log('No se pudo actualizar last_seen_at: ' . $throwable->getMessage());
    }

    $token = (string) ($_SESSION['workspace_session_token'] ?? '');
    if ($token !== '') {
        try {
            db()->prepare('UPDATE user_sessions SET last_seen_at = :now, is_active = 1 WHERE session_token = :session_token')->execute([
                'now' => $now,
                'session_token' => $token,
            ]);
        } catch (Throwable $throwable) {
            error_log('No se pudo actualizar user_session activa: ' . $throwable->getMessage());
        }
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

function buildAppUrl(string $path = ''): string
{
    $baseUrl = rtrim((string) env('APP_URL', ''), '/');

    if ($baseUrl === '') {
        $https = !empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off';
        $scheme = $https ? 'https' : 'http';
        $host = (string) ($_SERVER['HTTP_HOST'] ?? '');
        $scriptName = (string) ($_SERVER['SCRIPT_NAME'] ?? '');
        $basePath = dirname(dirname($scriptName));

        if ($basePath === '\\' || $basePath === '/') {
            $basePath = '';
        }

        if ($host !== '') {
            $baseUrl = $scheme . '://' . $host . rtrim(str_replace('\\', '/', $basePath), '/');
        }
    }

    $normalizedPath = '/' . ltrim($path, '/');
    if ($path === '') {
        $normalizedPath = '';
    }

    if ($baseUrl !== '') {
        return $baseUrl . $normalizedPath;
    }

    return $normalizedPath === '' ? '/' : $normalizedPath;
}

function decodeIntegrationMetadata(?string $metadata): array
{
    if (!is_string($metadata) || trim($metadata) === '') {
        return [];
    }

    $decoded = json_decode($metadata, true);
    return is_array($decoded) ? $decoded : [];
}

function encodeIntegrationMetadata(array $metadata): string
{
    return (string) json_encode($metadata, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
}

function ensureCalendarSource(int $userId, string $provider, ?string $externalAccountEmail = null): array
{
    $pdo = db();
    $stmt = $pdo->prepare('SELECT * FROM calendar_sources WHERE user_id = :user_id AND provider = :provider LIMIT 1');
    $stmt->execute([
        'user_id' => $userId,
        'provider' => $provider,
    ]);
    $source = $stmt->fetch();

    if ($source) {
        if ($externalAccountEmail !== null && $externalAccountEmail !== '' && $externalAccountEmail !== $source['external_account_email']) {
            $update = $pdo->prepare('UPDATE calendar_sources SET external_account_email = :external_account_email, updated_at = CURRENT_TIMESTAMP WHERE id = :id');
            $update->execute([
                'external_account_email' => $externalAccountEmail,
                'id' => (int) $source['id'],
            ]);
            $source['external_account_email'] = $externalAccountEmail;
        }

        return $source;
    }

    $insert = $pdo->prepare('INSERT INTO calendar_sources (user_id, provider, external_account_email, sync_enabled, sync_status, last_synced_at) VALUES (:user_id, :provider, :external_account_email, 1, :sync_status, NULL)');
    $insert->execute([
        'user_id' => $userId,
        'provider' => $provider,
        'external_account_email' => $externalAccountEmail !== '' ? $externalAccountEmail : null,
        'sync_status' => 'configured',
    ]);

    return [
        'id' => (int) $pdo->lastInsertId(),
        'user_id' => $userId,
        'provider' => $provider,
        'external_account_email' => $externalAccountEmail !== '' ? $externalAccountEmail : null,
        'sync_enabled' => 1,
        'sync_status' => 'configured',
        'last_synced_at' => null,
    ];
}

function getPowerAutomateIntegration(int $userId): ?array
{
    $stmt = db()->prepare('SELECT * FROM integrations WHERE user_id = :user_id AND provider = :provider LIMIT 1');
    $stmt->execute([
        'user_id' => $userId,
        'provider' => 'power_automate_calendar',
    ]);

    $integration = $stmt->fetch();
    return $integration ?: null;
}

function createOrRotatePowerAutomateSecret(int $userId, ?string $externalAccountEmail = null): array
{
    $pdo = db();
    $source = ensureCalendarSource($userId, 'power_automate', $externalAccountEmail);
    $integration = getPowerAutomateIntegration($userId);

    if (!$integration) {
        $insert = $pdo->prepare('INSERT INTO integrations (user_id, provider, status, metadata) VALUES (:user_id, :provider, :status, :metadata)');
        $insert->execute([
            'user_id' => $userId,
            'provider' => 'power_automate_calendar',
            'status' => 'configured',
            'metadata' => encodeIntegrationMetadata(['source_id' => (int) $source['id']]),
        ]);

        $integration = getPowerAutomateIntegration($userId);
    }

    if (!$integration) {
        throw new RuntimeException('No se pudo inicializar la integracion de Power Automate.');
    }

    $plainSecret = bin2hex(random_bytes(24));
    $metadata = decodeIntegrationMetadata($integration['metadata'] ?? null);
    $metadata['source_id'] = (int) $source['id'];
    $metadata['webhook_secret_hash'] = password_hash($plainSecret, PASSWORD_DEFAULT);
    $metadata['rotated_at'] = date(DATE_ATOM);
    if ($externalAccountEmail !== null && $externalAccountEmail !== '') {
        $metadata['external_account_email'] = strtolower(trim($externalAccountEmail));
    }

    $update = $pdo->prepare('UPDATE integrations SET status = :status, metadata = :metadata, updated_at = CURRENT_TIMESTAMP WHERE id = :id');
    $update->execute([
        'status' => 'configured',
        'metadata' => encodeIntegrationMetadata($metadata),
        'id' => (int) $integration['id'],
    ]);

    return [
        'token' => sprintf('pa_%d.%s', (int) $integration['id'], $plainSecret),
        'webhook_url' => buildAppUrl('/api/power_automate.php'),
        'header_name' => 'X-Power-Automate-Key',
        'external_account_email' => $metadata['external_account_email'] ?? $source['external_account_email'] ?? null,
    ];
}

function getPowerAutomateSetup(int $userId): array
{
    $integration = getPowerAutomateIntegration($userId);
    $source = ensureCalendarSource($userId, 'power_automate');
    $metadata = decodeIntegrationMetadata($integration['metadata'] ?? null);

    return [
        'configured'            => !empty($metadata['webhook_secret_hash']),
        'webhook_url'           => buildAppUrl('/api/power_automate.php'),
        'header_name'           => 'X-Power-Automate-Key',
        'token_preview'         => $integration ? 'pa_' . (int) $integration['id'] . '.***' : null,
        'external_account_email'=> $metadata['external_account_email'] ?? $source['external_account_email'] ?? null,
        'last_synced_at'        => $source['last_synced_at'] ?? null,
        'sync_status'           => $source['sync_status'] ?? 'pending',
        'last_payload_at'       => $metadata['last_payload_at'] ?? null,
        'outbound_webhook_url'  => $metadata['outbound_webhook_url'] ?? null,
    ];
}

function savePowerAutomateOutboundUrl(int $userId, ?string $url): void
{
    $pdo         = db();
    $integration = getPowerAutomateIntegration($userId);

    if (!$integration) {
        $source = ensureCalendarSource($userId, 'power_automate');
        $insert = $pdo->prepare('INSERT INTO integrations (user_id, provider, status, metadata) VALUES (:user_id, :provider, :status, :metadata)');
        $insert->execute([
            'user_id'  => $userId,
            'provider' => 'power_automate_calendar',
            'status'   => 'configured',
            'metadata' => encodeIntegrationMetadata(['source_id' => (int) $source['id']]),
        ]);
        $integration = getPowerAutomateIntegration($userId);
    }

    if (!$integration) {
        throw new RuntimeException('No se pudo inicializar la integracion de Power Automate.');
    }

    $metadata = decodeIntegrationMetadata($integration['metadata'] ?? null);
    if ($url !== null && $url !== '') {
        $metadata['outbound_webhook_url'] = $url;
    } else {
        unset($metadata['outbound_webhook_url']);
    }

    $pdo->prepare('UPDATE integrations SET metadata = :metadata, updated_at = CURRENT_TIMESTAMP WHERE id = :id')
        ->execute(['metadata' => encodeIntegrationMetadata($metadata), 'id' => (int) $integration['id']]);
}

function findPowerAutomateIntegrationByToken(string $token): ?array
{
    if (!preg_match('/^pa_(\d+)\.([a-f0-9]{48})$/i', trim($token), $matches)) {
        return null;
    }

    $integrationId = (int) $matches[1];
    $secret = $matches[2];

    $stmt = db()->prepare('SELECT * FROM integrations WHERE id = :id AND provider = :provider LIMIT 1');
    $stmt->execute([
        'id' => $integrationId,
        'provider' => 'power_automate_calendar',
    ]);
    $integration = $stmt->fetch();

    if (!$integration) {
        return null;
    }

    $metadata = decodeIntegrationMetadata($integration['metadata'] ?? null);
    $hash = (string) ($metadata['webhook_secret_hash'] ?? '');
    if ($hash === '' || !password_verify($secret, $hash)) {
        return null;
    }

    $integration['decoded_metadata'] = $metadata;
    return $integration;
}

function normalizeDateTimeValue(string $value): string
{
    try {
        $dateTime = new DateTimeImmutable($value);
    } catch (Exception $exception) {
        throw new InvalidArgumentException('Fecha invalida recibida desde Power Automate.');
    }

    return $dateTime->setTimezone(new DateTimeZone(date_default_timezone_get()))->format('Y-m-d H:i:s');
}

function detectTeamsSession(array $session): bool
{
    if (isset($session['is_teams_session'])) {
        return (bool) $session['is_teams_session'];
    }

    $haystack = strtolower(trim(implode(' ', array_filter([
        (string) ($session['location'] ?? ''),
        (string) ($session['meeting_url'] ?? ''),
        (string) ($session['title'] ?? ''),
        (string) ($session['description'] ?? ''),
        (string) ($session['online_meeting_provider'] ?? ''),
    ]))));

    return $haystack !== '' && (
        strpos($haystack, 'teams.microsoft.com') !== false ||
        strpos($haystack, 'microsoft teams') !== false ||
        strpos($haystack, 'teamsforbusiness') !== false
    );
}

function syncPowerAutomateDailySessions(array $integration, array $payload): array
{
    $pdo = db();
    $metadata = $integration['decoded_metadata'] ?? decodeIntegrationMetadata($integration['metadata'] ?? null);
    $sourceId = (int) ($metadata['source_id'] ?? 0);
    if ($sourceId <= 0) {
        throw new RuntimeException('La integracion no tiene una fuente de calendario asociada.');
    }

    $date = (string) ($payload['date'] ?? '');
    $sessions = $payload['sessions'] ?? null;
    if (!preg_match('/^\d{4}-\d{2}-\d{2}$/', $date) || !is_array($sessions)) {
        throw new InvalidArgumentException('El payload debe incluir date y sessions.');
    }

    $sourceEmail = strtolower(trim((string) ($payload['source_email'] ?? '')));
    $expectedEmail = strtolower(trim((string) ($metadata['external_account_email'] ?? '')));
    if ($expectedEmail !== '' && $sourceEmail !== '' && $expectedEmail !== $sourceEmail) {
        throw new RuntimeException('La cuenta origen no coincide con la configurada para esta integracion.');
    }

    $replaceDay = array_key_exists('replace_day', $payload) ? (bool) $payload['replace_day'] : true;
    $dayStart = $date . ' 00:00:00';
    $dayEnd = $date . ' 23:59:59';

    $pdo->beginTransaction();

    try {
        if ($replaceDay) {
            $delete = $pdo->prepare('DELETE FROM calendar_events WHERE user_id = :user_id AND source_id = :source_id AND source_type IN ("power_automate", "power_automate_teams") AND start_at BETWEEN :day_start AND :day_end');
            $delete->execute([
                'user_id' => (int) $integration['user_id'],
                'source_id' => $sourceId,
                'day_start' => $dayStart,
                'day_end' => $dayEnd,
            ]);
        }

        $insert = $pdo->prepare('INSERT INTO calendar_events (user_id, source_id, external_event_id, title, description, start_at, end_at, location, meeting_url, source_type) VALUES (:user_id, :source_id, :external_event_id, :title, :description, :start_at, :end_at, :location, :meeting_url, :source_type)');

        $storedSessions = 0;
        $teamsSessions = 0;
        $teamsMinutes = 0;

        foreach ($sessions as $session) {
            if (!is_array($session)) {
                continue;
            }

            $title = trim((string) ($session['title'] ?? ''));
            $startValue = trim((string) ($session['start_at'] ?? $session['start'] ?? ''));
            $endValue = trim((string) ($session['end_at'] ?? $session['end'] ?? ''));

            if ($title === '' || $startValue === '' || $endValue === '') {
                continue;
            }

            $startAt = normalizeDateTimeValue($startValue);
            $endAt = normalizeDateTimeValue($endValue);
            if ($endAt <= $startAt) {
                continue;
            }

            $isTeamsSession = detectTeamsSession($session);
            $sourceType = $isTeamsSession ? 'power_automate_teams' : 'power_automate';

            $insert->execute([
                'user_id' => (int) $integration['user_id'],
                'source_id' => $sourceId,
                'external_event_id' => trim((string) ($session['external_event_id'] ?? $session['id'] ?? '')) ?: null,
                'title' => $title,
                'description' => trim((string) ($session['description'] ?? '')) ?: null,
                'start_at' => $startAt,
                'end_at' => $endAt,
                'location' => trim((string) ($session['location'] ?? '')) ?: null,
                'meeting_url' => trim((string) ($session['meeting_url'] ?? $session['join_url'] ?? '')) ?: null,
                'source_type' => $sourceType,
            ]);

            $storedSessions++;

            if ($isTeamsSession) {
                $teamsSessions++;
                $teamsMinutes += max(0, (int) round((strtotime($endAt) - strtotime($startAt)) / 60));
            }
        }

        $syncTimestamp = date('Y-m-d H:i:s');
        $updateSource = $pdo->prepare('UPDATE calendar_sources SET sync_enabled = 1, sync_status = :sync_status, last_synced_at = :last_synced_at, external_account_email = :external_account_email WHERE id = :id');
        $updateSource->execute([
            'sync_status' => 'connected',
            'last_synced_at' => $syncTimestamp,
            'external_account_email' => $sourceEmail !== '' ? $sourceEmail : ($expectedEmail !== '' ? $expectedEmail : null),
            'id' => $sourceId,
        ]);

        $metadata['last_payload_at'] = date(DATE_ATOM);
        $metadata['last_sync_date'] = $date;
        $metadata['last_teams_minutes'] = $teamsMinutes;
        if ($sourceEmail !== '') {
            $metadata['external_account_email'] = $sourceEmail;
        }

        $updateIntegration = $pdo->prepare('UPDATE integrations SET status = :status, metadata = :metadata, updated_at = CURRENT_TIMESTAMP WHERE id = :id');
        $updateIntegration->execute([
            'status' => 'connected',
            'metadata' => encodeIntegrationMetadata($metadata),
            'id' => (int) $integration['id'],
        ]);

        $pdo->commit();

        return [
            'stored_sessions' => $storedSessions,
            'teams_sessions' => $teamsSessions,
            'teams_minutes' => $teamsMinutes,
            'date' => $date,
        ];
    } catch (Throwable $throwable) {
        $pdo->rollBack();
        throw $throwable;
    }
}

function formatMinutesAsHoursLabel(int $minutes): string
{
    $hours = intdiv(max(0, $minutes), 60);
    $remainingMinutes = max(0, $minutes) % 60;

    if ($hours === 0) {
        return $remainingMinutes . ' min';
    }

    if ($remainingMinutes === 0) {
        return $hours . ' h';
    }

    return $hours . ' h ' . $remainingMinutes . ' min';
}

function saveProfilePhoto(array $user, array $file, bool $alsoUpdatePublicProfile = false): string
{
    if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
        throw new RuntimeException('No se pudo subir la imagen.');
    }

    $fileSize = isset($file['size']) ? (int) $file['size'] : 0;
    if ($fileSize <= 0) {
        throw new RuntimeException('La imagen recibida esta vacia.');
    }

    if ($fileSize > 5 * 1024 * 1024) {
        throw new RuntimeException('La imagen supera el limite de 5MB.');
    }

    $tmpPath = (string) ($file['tmp_name'] ?? '');
    if ($tmpPath === '' || !is_uploaded_file($tmpPath)) {
        throw new RuntimeException('Archivo temporal invalido.');
    }

    $mimeType = mime_content_type($tmpPath) ?: 'application/octet-stream';
    $allowedMimeTypes = [
        'image/jpeg' => 'jpg',
        'image/png' => 'png',
        'image/webp' => 'webp',
    ];

    if (!isset($allowedMimeTypes[$mimeType])) {
        throw new RuntimeException('Solo se permiten imagenes PNG, JPG o WebP.');
    }

    if (@getimagesize($tmpPath) === false) {
        throw new RuntimeException('El archivo recibido no es una imagen valida.');
    }

    $extension = $allowedMimeTypes[$mimeType];

    $relativeDir = 'uploads/profile-photos';
    $absoluteDir = APP_ROOT . '/' . $relativeDir;
    if (!is_dir($absoluteDir) && !mkdir($absoluteDir, 0775, true) && !is_dir($absoluteDir)) {
        throw new RuntimeException('No se pudo crear la carpeta de fotos.');
    }

    $filename = sprintf('user-%d-%s-%s.%s', (int) $user['id'], date('YmdHis'), bin2hex(random_bytes(4)), $extension);
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
        'file_size' => $fileSize,
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

function ensureDeveloperIdentitySchema(PDO $pdo): void
{
    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS developer_identity_profile (
            id TINYINT UNSIGNED NOT NULL PRIMARY KEY,
            display_name VARCHAR(150) NOT NULL,
            role_label VARCHAR(150) NOT NULL,
            photo_path VARCHAR(255) DEFAULT NULL,
            owner_user_id INT UNSIGNED NOT NULL,
            owner_email VARCHAR(190) NOT NULL,
            updated_by_user_id INT UNSIGNED DEFAULT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            CONSTRAINT fk_dev_identity_owner FOREIGN KEY (owner_user_id) REFERENCES users(id) ON DELETE RESTRICT,
            CONSTRAINT fk_dev_identity_updated_by FOREIGN KEY (updated_by_user_id) REFERENCES users(id) ON DELETE SET NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );

    $existing = $pdo->query('SELECT id FROM developer_identity_profile WHERE id = 1 LIMIT 1')->fetchColumn();
    if ($existing !== false) {
        return;
    }

    $owner = resolveDeveloperIdentityOwnerUser($pdo);
    if ($owner === null) {
        return;
    }

    $insert = $pdo->prepare(
        'INSERT INTO developer_identity_profile (id, display_name, role_label, photo_path, owner_user_id, owner_email, updated_by_user_id)
         VALUES (1, :display_name, :role_label, :photo_path, :owner_user_id, :owner_email, :updated_by_user_id)'
    );
    $insert->execute([
        'display_name' => env('DEVELOPER_IDENTITY_NAME', (string) ($owner['full_name'] ?? 'Cristian Jesus Bonelo Rios')),
        'role_label' => env('DEVELOPER_IDENTITY_ROLE', 'Software Quality Analyst'),
        'photo_path' => $owner['profile_photo_path'] ?: null,
        'owner_user_id' => (int) $owner['id'],
        'owner_email' => (string) $owner['email'],
        'updated_by_user_id' => (int) $owner['id'],
    ]);
}

function resolveDeveloperIdentityOwnerUser(PDO $pdo): ?array
{
    $ownerEmail = strtolower(trim((string) env('DEVELOPER_IDENTITY_OWNER_EMAIL', 'cristianbonelorios@hotmail.com')));

    $byEmail = $pdo->prepare('SELECT * FROM users WHERE email = :email AND role = :role LIMIT 1');
    $byEmail->execute(['email' => $ownerEmail, 'role' => 'admin']);
    $owner = $byEmail->fetch(PDO::FETCH_ASSOC);
    if (is_array($owner)) {
        return $owner;
    }

    $firstAdmin = $pdo->query("SELECT * FROM users WHERE role = 'admin' ORDER BY id ASC LIMIT 1")->fetch(PDO::FETCH_ASSOC);
    return is_array($firstAdmin) ? $firstAdmin : null;
}

function getDeveloperIdentityProfile(): array
{
    $pdo = db();
    ensureDeveloperIdentitySchema($pdo);

    $stmt = $pdo->query('SELECT * FROM developer_identity_profile WHERE id = 1 LIMIT 1');
    $profile = $stmt->fetch(PDO::FETCH_ASSOC);
    if (!is_array($profile)) {
        throw new RuntimeException('No se pudo cargar el perfil del desarrollador.');
    }

    return [
        'id' => 1,
        'display_name' => (string) $profile['display_name'],
        'role_label' => (string) $profile['role_label'],
        'photo_path' => $profile['photo_path'] ? (string) $profile['photo_path'] : null,
        'photo_url' => $profile['photo_path'] ? buildAssetUrl((string) $profile['photo_path']) : null,
        'owner_user_id' => (int) $profile['owner_user_id'],
        'owner_email' => (string) $profile['owner_email'],
        'updated_by_user_id' => $profile['updated_by_user_id'] ? (int) $profile['updated_by_user_id'] : null,
        'created_at' => (string) $profile['created_at'],
        'updated_at' => (string) $profile['updated_at'],
    ];
}

function userCanManageDeveloperIdentity(array $user, ?array $profile = null): bool
{
    if ((string) ($user['role'] ?? '') !== 'admin') {
        return false;
    }

    $resolvedProfile = $profile ?? getDeveloperIdentityProfile();
    return (int) ($resolvedProfile['owner_user_id'] ?? 0) === (int) ($user['id'] ?? 0);
}

function saveDeveloperIdentityPhoto(array $actorUser, array $file): string
{
    if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
        throw new RuntimeException('No se pudo subir la imagen del desarrollador.');
    }

    $fileSize = isset($file['size']) ? (int) $file['size'] : 0;
    if ($fileSize <= 0) {
        throw new RuntimeException('La imagen recibida esta vacia.');
    }

    if ($fileSize > 5 * 1024 * 1024) {
        throw new RuntimeException('La imagen supera el limite de 5MB.');
    }

    $tmpPath = (string) ($file['tmp_name'] ?? '');
    if ($tmpPath === '' || !is_uploaded_file($tmpPath)) {
        throw new RuntimeException('Archivo temporal invalido.');
    }

    $mimeType = mime_content_type($tmpPath) ?: 'application/octet-stream';
    $allowedMimeTypes = [
        'image/jpeg' => 'jpg',
        'image/png' => 'png',
        'image/webp' => 'webp',
    ];

    if (!isset($allowedMimeTypes[$mimeType])) {
        throw new RuntimeException('Solo se permiten imagenes PNG, JPG o WebP.');
    }

    if (@getimagesize($tmpPath) === false) {
        throw new RuntimeException('El archivo recibido no es una imagen valida.');
    }

    $extension = $allowedMimeTypes[$mimeType];
    $relativeDir = 'uploads/developer-photos';
    $absoluteDir = APP_ROOT . '/' . $relativeDir;
    if (!is_dir($absoluteDir) && !mkdir($absoluteDir, 0775, true) && !is_dir($absoluteDir)) {
        throw new RuntimeException('No se pudo crear la carpeta de foto del desarrollador.');
    }

    $filename = sprintf('developer-owner-%d-%s-%s.%s', (int) $actorUser['id'], date('YmdHis'), bin2hex(random_bytes(4)), $extension);
    $relativePath = $relativeDir . '/' . $filename;
    $absolutePath = APP_ROOT . '/' . $relativePath;

    if (!move_uploaded_file($tmpPath, $absolutePath)) {
        throw new RuntimeException('No se pudo guardar la foto del desarrollador en el servidor.');
    }

    db()->prepare('UPDATE developer_identity_profile SET photo_path = :photo_path, updated_by_user_id = :updated_by_user_id WHERE id = 1')
        ->execute([
            'photo_path' => $relativePath,
            'updated_by_user_id' => (int) $actorUser['id'],
        ]);

    return $relativePath;
}

function ensureNotesCollaborationSchema(PDO $pdo): void
{
    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS note_comments (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            note_id INT UNSIGNED NOT NULL,
            user_id INT UNSIGNED NOT NULL,
            parent_comment_id BIGINT UNSIGNED DEFAULT NULL,
            content TEXT NOT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            KEY idx_note_comments_note (note_id),
            KEY idx_note_comments_user (user_id),
            KEY idx_note_comments_parent (parent_comment_id),
            CONSTRAINT fk_note_comments_note FOREIGN KEY (note_id) REFERENCES notes(id) ON DELETE CASCADE,
            CONSTRAINT fk_note_comments_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
            CONSTRAINT fk_note_comments_parent FOREIGN KEY (parent_comment_id) REFERENCES note_comments(id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );

    $pdo->exec(
        'CREATE TABLE IF NOT EXISTS note_shares (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            note_id INT UNSIGNED NOT NULL,
            owner_user_id INT UNSIGNED NOT NULL,
            invited_email VARCHAR(190) NOT NULL,
            invited_user_id INT UNSIGNED DEFAULT NULL,
            is_active TINYINT(1) NOT NULL DEFAULT 1,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY uniq_note_share_email (note_id, invited_email),
            KEY idx_note_shares_note (note_id),
            KEY idx_note_shares_owner (owner_user_id),
            KEY idx_note_shares_invited_user (invited_user_id),
            CONSTRAINT fk_note_shares_note FOREIGN KEY (note_id) REFERENCES notes(id) ON DELETE CASCADE,
            CONSTRAINT fk_note_shares_owner FOREIGN KEY (owner_user_id) REFERENCES users(id) ON DELETE CASCADE,
            CONSTRAINT fk_note_shares_invited_user FOREIGN KEY (invited_user_id) REFERENCES users(id) ON DELETE SET NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );
}