<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

try {
    require_once __DIR__ . '/app.php';
    require_once __DIR__ . '/webauthn_helper.php';
    ensureApplicationInstalled();
    ensureBiometricSchema(db());

    $input = jsonInput();
    $action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? ''));

    // ── Morpho E3/E4 ──────────────────────────────────────────────────────────
    if ($action === 'capture-fingerprint') {
        $user = requirePageAuth();
        captureFingerprintAction($user);
    } elseif ($action === 'save-fingerprint') {
        $user = requirePageAuth();
        saveFingerprintAction($user, $input);
    } elseif ($action === 'verify-fingerprint') {
        verifyFingerprintAction($input);

    // ── WebAuthn (lector del PC) ──────────────────────────────────────────────
    } elseif ($action === 'webauthn-register-options') {
        $user = requirePageAuth();
        webAuthnRegisterOptionsAction($user);
    } elseif ($action === 'webauthn-register-complete') {
        $user = requirePageAuth();
        webAuthnRegisterCompleteAction($user, $input);
    } elseif ($action === 'webauthn-auth-options') {
        webAuthnAuthOptionsAction($input);
    } elseif ($action === 'webauthn-auth-complete') {
        webAuthnAuthCompleteAction($input);

    // ── Gestión común ─────────────────────────────────────────────────────────
    } elseif ($action === 'check-biometric-status') {
        $user = requirePageAuth();
        checkBiometricStatusAction($user);
    } elseif ($action === 'disable-biometric') {
        $user = requirePageAuth();
        disableBiometricAction($user, $input);
    } else {
        jsonResponse(['ok' => false, 'message' => 'Accion biometrica no valida.'], 400);
    }
} catch (Throwable $throwable) {
    jsonResponse([
        'ok'      => false,
        'message' => 'Error en API biometrica: ' . $throwable->getMessage(),
    ], 500);
}

/**
 * Asegura que las tablas de biometría existan (creación lazy).
 */
function ensureBiometricSchema(PDO $pdo): void
{
    $pdo->exec('
        CREATE TABLE IF NOT EXISTS user_webauthn_credentials (
            id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id INT UNSIGNED NOT NULL,
            credential_id VARCHAR(1024) NOT NULL UNIQUE,
            public_key_pem TEXT NOT NULL,
            sign_count INT UNSIGNED NOT NULL DEFAULT 0,
            device_label VARCHAR(100) DEFAULT "platform",
            is_enabled TINYINT(1) NOT NULL DEFAULT 1,
            last_used_at DATETIME DEFAULT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            KEY idx_webauthn_user (user_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    ');
    $pdo->exec('
        CREATE TABLE IF NOT EXISTS user_biometric_fingerprints (
            id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id INT UNSIGNED NOT NULL UNIQUE,
            template_data LONGBLOB NOT NULL,
            fingerprint_hash VARCHAR(255) NOT NULL,
            device_type VARCHAR(100) NOT NULL DEFAULT "morpho_e3_e4",
            device_serial VARCHAR(255) DEFAULT NULL,
            is_enabled TINYINT(1) NOT NULL DEFAULT 1,
            last_verified_at DATETIME DEFAULT NULL,
            verification_count INT UNSIGNED NOT NULL DEFAULT 0,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    ');
    $pdo->exec('
        CREATE TABLE IF NOT EXISTS biometric_login_attempts (
            id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
            user_id INT UNSIGNED DEFAULT NULL,
            device_type VARCHAR(100) NOT NULL DEFAULT "morpho_e3_e4",
            attempt_type ENUM("capture","verify","failed_verify","success") NOT NULL,
            status ENUM("success","failure","timeout","device_error") NOT NULL,
            error_message TEXT DEFAULT NULL,
            ip_address VARCHAR(45) DEFAULT NULL,
            user_agent VARCHAR(255) DEFAULT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    ');
}

/**
 * Devuelve el estado unificado de Morpho + WebAuthn para el usuario.
 */
function checkBiometricStatusAction(array $user): void
{
    $pdo = db();

    // Morpho E3/E4
    $stmt = $pdo->prepare('
        SELECT is_enabled, device_type, last_verified_at, verification_count
        FROM user_biometric_fingerprints
        WHERE user_id = :uid
        LIMIT 1
    ');
    $stmt->execute(['uid' => $user['id']]);
    $morpho = $stmt->fetch();

    // WebAuthn (lector del PC)
    $stmt2 = $pdo->prepare('
        SELECT COUNT(*) AS cnt, MAX(last_used_at) AS last_used
        FROM user_webauthn_credentials
        WHERE user_id = :uid AND is_enabled = 1
    ');
    $stmt2->execute(['uid' => $user['id']]);
    $wa = $stmt2->fetch();

    jsonResponse([
        'ok'          => true,
        'morpho'      => $morpho ? [
            'enrolled'           => true,
            'enabled'            => (bool) $morpho['is_enabled'],
            'device_type'        => $morpho['device_type'],
            'last_verified'      => $morpho['last_verified_at'],
            'verification_count' => (int) $morpho['verification_count'],
        ] : ['enrolled' => false],
        'webauthn'    => [
            'enrolled'         => ((int) $wa['cnt']) > 0,
            'credential_count' => (int) $wa['cnt'],
            'last_used'        => $wa['last_used'],
        ],
        'has_biometric' => (bool) ($morpho && $morpho['is_enabled']) || ((int) $wa['cnt']) > 0,
    ]);
}

/**
 * Captura una huella digital (primera verificación)
 */
function captureFingerprintAction(array $user): void
{
    // Esta es una acción que señala que se debe capturar
    // El cliente (JavaScript) se encargará de comunicarse con el captor
    jsonResponse([
        'ok' => true,
        'message' => 'Listo para capturar huella digital. Acerca tu dedo al captor.',
        'device_type' => 'morpho_e3_e4',
        'user_id' => $user['id']
    ]);
}

/**
 * Guarda la huella digital capturada
 */
function saveFingerprintAction(array $user, array $input): void
{
    $pdo = db();
    
    // Validar que se reciba el template y hash
    $templateData = (string) ($input['template_data'] ?? '');
    $fingerprintHash = (string) ($input['fingerprint_hash'] ?? '');
    $deviceSerial = (string) ($input['device_serial'] ?? '');

    if (empty($templateData) || empty($fingerprintHash)) {
        jsonResponse([
            'ok' => false,
            'message' => 'Template de huella digital requerido.'
        ], 422);
    }

    // Verificar que el template tenga un tamaño razonable
    if (strlen($templateData) < 100 || strlen($templateData) > 5000000) {
        jsonResponse([
            'ok' => false,
            'message' => 'Tamaño de template invalido.'
        ], 422);
    }

    try {
        // Verificar si ya existe huella digital registrada
        $stmt = $pdo->prepare('
            SELECT id FROM user_biometric_fingerprints 
            WHERE user_id = :user_id 
            LIMIT 1
        ');
        $stmt->execute(['user_id' => $user['id']]);
        $existing = $stmt->fetch();

        if ($existing) {
            // Actualizar la huella existente
            $stmt = $pdo->prepare('
                UPDATE user_biometric_fingerprints 
                SET template_data = :template_data,
                    fingerprint_hash = :fingerprint_hash,
                    device_serial = :device_serial,
                    is_enabled = 1,
                    updated_at = NOW() 
                WHERE user_id = :user_id
            ');
            $result = $stmt->execute([
                'template_data' => $templateData,
                'fingerprint_hash' => $fingerprintHash,
                'device_serial' => $deviceSerial,
                'user_id' => $user['id']
            ]);
        } else {
            // Insertar nueva huella digital
            $stmt = $pdo->prepare('
                INSERT INTO user_biometric_fingerprints 
                (user_id, template_data, fingerprint_hash, device_type, device_serial, is_enabled) 
                VALUES 
                (:user_id, :template_data, :fingerprint_hash, :device_type, :device_serial, 1)
            ');
            $result = $stmt->execute([
                'user_id' => $user['id'],
                'template_data' => $templateData,
                'fingerprint_hash' => $fingerprintHash,
                'device_type' => 'morpho_e3_e4',
                'device_serial' => $deviceSerial
            ]);
        }

        if (!$result) {
            jsonResponse([
                'ok' => false,
                'message' => 'No se pudo guardar la huella digital.'
            ], 500);
        }

        // Registrar el intento en el historial
        logBiometricAttempt($pdo, $user['id'], 'capture', 'success');

        jsonResponse([
            'ok' => true,
            'message' => 'Huella digital guardada exitosamente.'
        ]);
    } catch (PDOException $e) {
        jsonResponse([
            'ok' => false,
            'message' => 'Error en base de datos: ' . $e->getMessage()
        ], 500);
    }
}

/**
 * Verifica la huella digital durante el login
 */
function verifyFingerprintAction(array $input): void
{
    $pdo = db();
    
    $email = (string) ($input['email'] ?? '');
    $fingerprintHash = (string) ($input['fingerprint_hash'] ?? '');

    if (empty($email) || empty($fingerprintHash)) {
        logBiometricAttempt($pdo, null, 'verify', 'failure', 'Datos invalidos');
        jsonResponse([
            'ok' => false,
            'message' => 'Email y huella digital requeridos.'
        ], 422);
    }

    try {
        // Buscar al usuario
        $stmt = $pdo->prepare('SELECT id, email FROM users WHERE email = :email LIMIT 1');
        $stmt->execute(['email' => $email]);
        $user = $stmt->fetch();

        if (!$user) {
            logBiometricAttempt($pdo, null, 'verify', 'failure', 'Usuario no encontrado');
            jsonResponse([
                'ok' => false,
                'message' => 'Usuario no encontrado.'
            ], 404);
        }

        // Buscar la huella digital del usuario
        $stmt = $pdo->prepare('
            SELECT template_data, fingerprint_hash, is_enabled 
            FROM user_biometric_fingerprints 
            WHERE user_id = :user_id AND is_enabled = 1
            LIMIT 1
        ');
        $stmt->execute(['user_id' => $user['id']]);
        $biometric = $stmt->fetch();

        if (!$biometric) {
            logBiometricAttempt($pdo, $user['id'], 'verify', 'failure', 'No hay huella registrada');
            jsonResponse([
                'ok' => false,
                'message' => 'Este usuario no tiene huella digital registrada.'
            ], 404);
        }

        // Comparar hashes (en una implementación real, usarías comparación más sofisticada)
        $isValid = hash_equals($biometric['fingerprint_hash'], $fingerprintHash);

        if ($isValid) {
            // Actualizar información de última verificación
            $stmt = $pdo->prepare('
                UPDATE user_biometric_fingerprints 
                SET last_verified_at = NOW(),
                    verification_count = verification_count + 1
                WHERE user_id = :user_id
            ');
            $stmt->execute(['user_id' => $user['id']]);

            // Registrar intento exitoso
            logBiometricAttempt($pdo, $user['id'], 'verify', 'success');

            // Iniciar sesión del usuario
            $stmt = $pdo->prepare('SELECT * FROM users WHERE id = :id LIMIT 1');
            $stmt->execute(['id' => $user['id']]);
            $userToAuthenticate = $stmt->fetch();

            $authenticatedUser = startAuthenticatedSession($userToAuthenticate);

            jsonResponse([
                'ok' => true,
                'message' => 'Autenticacion biometrica exitosa.',
                'user' => $authenticatedUser
            ]);
        } else {
            logBiometricAttempt($pdo, $user['id'], 'verify', 'failure', 'Huella no coincide');
            jsonResponse([
                'ok' => false,
                'message' => 'La huella digital no coincide.'
            ], 403);
        }
    } catch (PDOException $e) {
        logBiometricAttempt($pdo, null, 'verify', 'failure', 'Error DB: ' . $e->getMessage());
        jsonResponse([
            'ok' => false,
            'message' => 'Error en base de datos.'
        ], 500);
    }
}

/**
 * Deshabilita la autenticación biométrica del usuario.
 * $input['type'] puede ser 'morpho', 'webauthn' o vacío (ambos).
 */
function disableBiometricAction(array $user, array $input): void
{
    $pdo  = db();
    $type = (string) ($input['type'] ?? 'all');

    if ($type === 'morpho' || $type === 'all') {
        $pdo->prepare('UPDATE user_biometric_fingerprints SET is_enabled = 0, updated_at = NOW() WHERE user_id = :uid')
            ->execute(['uid' => $user['id']]);
    }
    if ($type === 'webauthn' || $type === 'all') {
        $pdo->prepare('UPDATE user_webauthn_credentials SET is_enabled = 0, updated_at = NOW() WHERE user_id = :uid')
            ->execute(['uid' => $user['id']]);
    }

    jsonResponse(['ok' => true, 'message' => 'Autenticacion biometrica deshabilitada.']);
}

// ─── WebAuthn handlers ────────────────────────────────────────────────────────

/**
 * Retorna las opciones de registro para WebAuthn (reto + parámetros).
 */
function webAuthnRegisterOptionsAction(array $user): void
{
    $options = webauthn_registration_options($user);
    jsonResponse(['ok' => true, 'options' => $options]);
}

/**
 * Verifica la respuesta del navegador y guarda la credencial WebAuthn.
 */
function webAuthnRegisterCompleteAction(array $user, array $input): void
{
    $pdo = db();

    try {
        $credential = webauthn_verify_registration($input['response'] ?? $input);

        // Upsert: una credencial por usuario (se puede ampliar a múltiples)
        $pdo->prepare('
            DELETE FROM user_webauthn_credentials WHERE user_id = :uid
        ')->execute(['uid' => $user['id']]);

        $pdo->prepare('
            INSERT INTO user_webauthn_credentials
                (user_id, credential_id, public_key_pem, sign_count, device_label, is_enabled)
            VALUES
                (:uid, :cred_id, :pk_pem, :sign_count, "platform", 1)
        ')->execute([
            'uid'        => $user['id'],
            'cred_id'    => $credential['credentialId'],
            'pk_pem'     => $credential['publicKeyPem'],
            'sign_count' => $credential['signCount'],
        ]);

        logBiometricAttempt($pdo, $user['id'], 'capture', 'success');
        jsonResponse(['ok' => true, 'message' => 'Lector de huella del PC registrado exitosamente.']);
    } catch (Throwable $e) {
        logBiometricAttempt($pdo, $user['id'], 'capture', 'failure', $e->getMessage());
        jsonResponse(['ok' => false, 'message' => 'No se pudo registrar: ' . $e->getMessage()], 422);
    }
}

/**
 * Retorna el reto de autenticación para WebAuthn.
 * Requiere el email del usuario para devolver las credentialIds correctas.
 */
function webAuthnAuthOptionsAction(array $input): void
{
    $pdo   = db();
    $email = trim((string) ($input['email'] ?? ''));

    if ($email === '') {
        jsonResponse(['ok' => false, 'message' => 'Email requerido.'], 422);
    }

    $stmt = $pdo->prepare('SELECT id FROM users WHERE email = :email LIMIT 1');
    $stmt->execute(['email' => $email]);
    $user = $stmt->fetch();

    if (!$user) {
        // No revelar si el usuario existe
        jsonResponse(['ok' => false, 'message' => 'No se encontraron credenciales biometricas.'], 404);
    }

    $stmt2 = $pdo->prepare('
        SELECT credential_id FROM user_webauthn_credentials
        WHERE user_id = :uid AND is_enabled = 1
    ');
    $stmt2->execute(['uid' => $user['id']]);
    $creds = $stmt2->fetchAll();

    if (empty($creds)) {
        jsonResponse(['ok' => false, 'message' => 'Este usuario no tiene lector de huella del PC registrado.'], 404);
    }

    $options = webauthn_auth_options($creds);
    jsonResponse(['ok' => true, 'options' => $options, 'user_id' => $user['id']]);
}

/**
 * Verifica la aserción WebAuthn y crea la sesión del usuario.
 */
function webAuthnAuthCompleteAction(array $input): void
{
    $pdo    = db();
    $userId = (int) ($input['user_id'] ?? 0);
    $credId = (string) ($input['credential_id'] ?? '');

    if ($userId <= 0 || $credId === '') {
        jsonResponse(['ok' => false, 'message' => 'Datos incompletos.'], 422);
    }

    // Cargar la credencial almacenada
    $stmt = $pdo->prepare('
        SELECT id, credential_id, public_key_pem, sign_count
        FROM user_webauthn_credentials
        WHERE user_id = :uid AND credential_id = :cred_id AND is_enabled = 1
        LIMIT 1
    ');
    $stmt->execute(['uid' => $userId, 'cred_id' => $credId]);
    $storedCred = $stmt->fetch();

    if (!$storedCred) {
        logBiometricAttempt($pdo, $userId, 'verify', 'failure', 'Credencial no encontrada');
        jsonResponse(['ok' => false, 'message' => 'Credencial no encontrada.'], 404);
    }

    try {
        $newSignCount = webauthn_verify_assertion($input['response'] ?? $input, $storedCred);

        // Actualizar signCount y last_used_at
        $pdo->prepare('
            UPDATE user_webauthn_credentials
            SET sign_count = :cnt, last_used_at = NOW(), updated_at = NOW()
            WHERE id = :id
        ')->execute(['cnt' => $newSignCount, 'id' => $storedCred['id']]);

        // Iniciar sesión
        $stmt2 = $pdo->prepare('SELECT * FROM users WHERE id = :id LIMIT 1');
        $stmt2->execute(['id' => $userId]);
        $userRow = $stmt2->fetch();

        if (!$userRow) {
            jsonResponse(['ok' => false, 'message' => 'Usuario no encontrado.'], 404);
        }

        logBiometricAttempt($pdo, $userId, 'verify', 'success');
        $authenticatedUser = startAuthenticatedSession($userRow);

        jsonResponse([
            'ok'      => true,
            'message' => 'Autenticacion biometrica exitosa.',
            'user'    => $authenticatedUser,
        ]);
    } catch (Throwable $e) {
        logBiometricAttempt($pdo, $userId, 'failed_verify', 'failure', $e->getMessage());
        jsonResponse(['ok' => false, 'message' => 'Verificacion fallida: ' . $e->getMessage()], 403);
    }
}

/**
 * Registra intentos de autenticación biométrica para auditoría
 */
function logBiometricAttempt(
    PDO $pdo,
    ?int $userId,
    string $attemptType,
    string $status,
    string $errorMessage = ''
): void {
    try {
        $stmt = $pdo->prepare('
            INSERT INTO biometric_login_attempts 
            (user_id, device_type, attempt_type, status, error_message, ip_address, user_agent)
            VALUES
            (:user_id, :device_type, :attempt_type, :status, :error_message, :ip_address, :user_agent)
        ');
        $stmt->execute([
            'user_id' => $userId,
            'device_type' => 'morpho_e3_e4',
            'attempt_type' => $attemptType,
            'status' => $status,
            'error_message' => $errorMessage ?: null,
            'ip_address' => getClientIp(),
            'user_agent' => substr($_SERVER['HTTP_USER_AGENT'] ?? '', 0, 255)
        ]);
    } catch (PDOException $e) {
        // Ignorar errores de logging
    }
}

function getClientIp(): string
{
    $ip = '';
    if (array_key_exists('HTTP_CF_CONNECTING_IP', $_SERVER)) {
        $ip = $_SERVER['HTTP_CF_CONNECTING_IP'];
    } elseif (array_key_exists('HTTP_X_FORWARDED_FOR', $_SERVER)) {
        $ip = trim(explode(',', $_SERVER['HTTP_X_FORWARDED_FOR'])[0]);
    } elseif (array_key_exists('REMOTE_ADDR', $_SERVER)) {
        $ip = $_SERVER['REMOTE_ADDR'];
    }
    return substr($ip, 0, 45);
}
