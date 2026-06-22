<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    $user = requireApiAuth();
    $pdo = db();

    $input = jsonInput();
    $action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? 'history'));

    switch ($action) {
        case 'security_status':
            $security = getDocumentSecurity($pdo, (int) $user['id']);
            jsonResponse([
                'ok' => true,
                'configured' => $security !== null,
                'enabled' => $security !== null && (int) $security['is_enabled'] === 1,
                'verified' => isDocumentAccessVerified((int) $user['id']),
                'last_verified_at' => $security['last_verified_at'] ?? null,
            ]);
            break;

        case 'set_access_key':
            $accessKey = trim((string) ($input['access_key'] ?? $_POST['access_key'] ?? ''));
            if (mb_strlen($accessKey) < 6) {
                jsonResponse(['ok' => false, 'message' => 'La clave secundaria debe tener al menos 6 caracteres.'], 422);
            }

            upsertDocumentSecurity($pdo, (int) $user['id'], $accessKey);
            jsonResponse(['ok' => true, 'message' => 'Clave secundaria guardada correctamente.']);
            break;

        case 'remove_access_key':
            $stmt = $pdo->prepare('DELETE FROM user_document_security WHERE user_id = :user_id');
            $stmt->execute(['user_id' => (int) $user['id']]);
            clearDocumentAccessVerification();
            jsonResponse(['ok' => true, 'message' => 'Clave secundaria eliminada.']);
            break;

        case 'verify_access_key':
            $accessKey = trim((string) ($input['access_key'] ?? $_POST['access_key'] ?? ''));
            $security = getDocumentSecurity($pdo, (int) $user['id']);
            if ($security === null || (int) $security['is_enabled'] !== 1) {
                jsonResponse(['ok' => true, 'message' => 'No hay clave secundaria configurada.', 'verified' => true]);
            }

            if ($accessKey === '' || !password_verify($accessKey, (string) $security['access_key_hash'])) {
                jsonResponse(['ok' => false, 'message' => 'Clave secundaria incorrecta.'], 401);
            }

            markDocumentAccessVerified($pdo, (int) $user['id']);
            jsonResponse(['ok' => true, 'message' => 'Acceso verificado.', 'verified' => true]);
            break;

        case 'get_config':
            jsonResponse([
                'ok'             => true,
                'webhook_url'    => getDocumentWebhookUrl($pdo, (int) $user['id']),
                'callback_url'   => getDocumentCallbackUrl(),
                'callback_token' => getOrCreateCallbackToken($pdo),
            ]);
            break;

        case 'set_webhook_url':
            $url = trim((string) ($input['webhook_url'] ?? $_POST['webhook_url'] ?? ''));
            if ($url !== '' && !filter_var($url, FILTER_VALIDATE_URL)) {
                jsonResponse(['ok' => false, 'message' => 'La URL de Power Automate no es valida.'], 422);
            }

            setDocumentWebhookUrl($pdo, (int) $user['id'], $url !== '' ? $url : null);
            jsonResponse(['ok' => true, 'message' => 'URL de Power Automate guardada.']);
            break;

        case 'download':
            enforceDocumentAccess($pdo, (int) $user['id']);

            $jobId = (int) ($_GET['id'] ?? $input['id'] ?? 0);
            if ($jobId <= 0) {
                jsonResponse(['ok' => false, 'message' => 'ID invalido.'], 422);
            }

            $job = fetchDocumentJob($pdo, (int) $user['id'], $jobId);
            if ($job === null) {
                jsonResponse(['ok' => false, 'message' => 'Documento no encontrado.'], 404);
            }

            if ((string) ($job['status'] ?? '') !== 'completed' || empty($job['output_file_url'])) {
                jsonResponse(['ok' => false, 'message' => 'El documento aun no esta listo para descarga.'], 422);
            }

            jsonResponse([
                'ok' => true,
                'download_url' => (string) $job['output_file_url'],
                'file_name' => (string) ($job['output_file_name'] ?? ''),
            ]);
            break;

        case 'create':
        case 'generate':
            enforceDocumentAccess($pdo, (int) $user['id']);

            $generationType = trim((string) ($_POST['generation_type'] ?? $input['generation_type'] ?? 'manual'));
            $description = trim((string) ($_POST['description'] ?? $input['description'] ?? ''));
            $allowedGenerationTypes = ['manual', 'guia', 'informe'];

            if (!in_array($generationType, $allowedGenerationTypes, true)) {
                jsonResponse(['ok' => false, 'message' => 'Tipo de generacion invalido.'], 422);
            }

            $uploadedFiles = normalizeUploadedFiles($_FILES['files'] ?? null);
            if ($uploadedFiles === []) {
                jsonResponse(['ok' => false, 'message' => 'Debes adjuntar al menos un archivo.'], 422);
            }

            $webhookUrl = getDocumentWebhookUrl($pdo, (int) $user['id']);
            if ($webhookUrl === null || $webhookUrl === '') {
                jsonResponse(['ok' => false, 'message' => 'Configura primero la URL de Power Automate para documentos.'], 422);
            }

            $results = [];
            foreach ($uploadedFiles as $file) {
                $results[] = createDocumentGenerationJob(
                    $pdo,
                    $user,
                    $file,
                    $generationType,
                    $description,
                    $webhookUrl
                );
            }

            jsonResponse([
                'ok' => true,
                'message' => 'Proceso de generacion ejecutado.',
                'jobs' => $results,
            ]);
            break;

        case 'delete_job':
            enforceDocumentAccess($pdo, (int) $user['id']);

            $jobId = (int) ($_GET['id'] ?? $input['id'] ?? 0);
            if ($jobId <= 0) {
                jsonResponse(['ok' => false, 'message' => 'ID de trabajo invalido.'], 422);
            }

            // Verificar que el trabajo pertenece al usuario
            $job = fetchDocumentJob($pdo, (int) $user['id'], $jobId);
            if ($job === null) {
                jsonResponse(['ok' => false, 'message' => 'Trabajo no encontrado.'], 404);
            }

            // Eliminar el trabajo
            $stmt = $pdo->prepare('DELETE FROM document_generation_jobs WHERE id = :id AND user_id = :user_id');
            $stmt->execute([
                'id' => $jobId,
                'user_id' => (int) $user['id'],
            ]);

            jsonResponse([
                'ok' => true,
                'message' => 'Trabajo eliminado correctamente.',
            ]);
            break;

        case 'history':
        default:
            enforceDocumentAccess($pdo, (int) $user['id']);

            $stmt = $pdo->prepare(
                'SELECT id, input_file_name, generation_type, status, output_file_name, output_file_url, error_message, created_at, updated_at, completed_at
                 FROM document_generation_jobs
                 WHERE user_id = :user_id
                 ORDER BY created_at DESC
                 LIMIT 120'
            );
            $stmt->execute(['user_id' => (int) $user['id']]);

            jsonResponse([
                'ok' => true,
                'jobs' => $stmt->fetchAll() ?: [],
            ]);
            break;
    }
} catch (Throwable $throwable) {
    error_log('generacion_documentos.php error: ' . $throwable->getMessage() . ' in ' . $throwable->getFile() . ':' . $throwable->getLine());

    jsonResponse([
        'ok' => false,
        'message' => applicationErrorMessage($throwable),
    ], 500);
}

function normalizeUploadedFiles($raw): array
{
    if (!is_array($raw) || !isset($raw['name'])) {
        return [];
    }

    if (!is_array($raw['name'])) {
        if (($raw['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
            return [];
        }

        return [$raw];
    }

    $normalized = [];
    $count = count($raw['name']);
    for ($i = 0; $i < $count; $i++) {
        $error = $raw['error'][$i] ?? UPLOAD_ERR_NO_FILE;
        if ($error !== UPLOAD_ERR_OK) {
            continue;
        }

        $normalized[] = [
            'name' => $raw['name'][$i] ?? 'archivo',
            'type' => $raw['type'][$i] ?? '',
            'tmp_name' => $raw['tmp_name'][$i] ?? '',
            'error' => $error,
            'size' => $raw['size'][$i] ?? 0,
        ];
    }

    return $normalized;
}

function getDocumentCallbackUrl(): string
{
    return buildAppUrl('/api/document_callback.php');
}

function getOrCreateCallbackToken(PDO $pdo): string
{
    $key  = 'document_generation_callback_secret';
    $stmt = $pdo->prepare('SELECT setting_value FROM app_settings WHERE setting_key = :key LIMIT 1');
    $stmt->execute(['key' => $key]);
    $existing = (string) ($stmt->fetchColumn() ?: '');

    if ($existing !== '') {
        return $existing;
    }

    $token = bin2hex(random_bytes(32));
    $pdo->prepare(
        'INSERT INTO app_settings (setting_key, setting_value)
         VALUES (:key, :value)
         ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value), updated_at = CURRENT_TIMESTAMP'
    )->execute(['key' => $key, 'value' => $token]);

    return $token;
}

function getDocumentWebhookSettingKey(int $userId): string
{
    return 'document_generation_pa_url_user_' . $userId;
}

function getDocumentWebhookUrl(PDO $pdo, int $userId): ?string
{
    $stmt = $pdo->prepare('SELECT setting_value FROM app_settings WHERE setting_key = :setting_key LIMIT 1');
    $stmt->execute(['setting_key' => getDocumentWebhookSettingKey($userId)]);
    $value = $stmt->fetchColumn();

    if (!is_string($value) || trim($value) === '') {
        return null;
    }

    return trim($value);
}

function setDocumentWebhookUrl(PDO $pdo, int $userId, ?string $url): void
{
    $settingKey = getDocumentWebhookSettingKey($userId);

    if ($url === null || trim($url) === '') {
        $delete = $pdo->prepare('DELETE FROM app_settings WHERE setting_key = :setting_key');
        $delete->execute(['setting_key' => $settingKey]);
        return;
    }

    $stmt = $pdo->prepare(
        'INSERT INTO app_settings (setting_key, setting_value)
         VALUES (:setting_key, :setting_value)
         ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value), updated_at = CURRENT_TIMESTAMP'
    );
    $stmt->execute([
        'setting_key' => $settingKey,
        'setting_value' => $url,
    ]);
}

function getDocumentSecurity(PDO $pdo, int $userId): ?array
{
    $stmt = $pdo->prepare('SELECT * FROM user_document_security WHERE user_id = :user_id LIMIT 1');
    $stmt->execute(['user_id' => $userId]);
    $row = $stmt->fetch();

    return is_array($row) ? $row : null;
}

function upsertDocumentSecurity(PDO $pdo, int $userId, string $accessKey): void
{
    $hash = password_hash($accessKey, PASSWORD_BCRYPT);

    $stmt = $pdo->prepare(
        'INSERT INTO user_document_security (user_id, access_key_hash, is_enabled)
         VALUES (:user_id, :access_key_hash, 1)
         ON DUPLICATE KEY UPDATE access_key_hash = VALUES(access_key_hash), is_enabled = 1, updated_at = CURRENT_TIMESTAMP'
    );
    $stmt->execute([
        'user_id' => $userId,
        'access_key_hash' => $hash,
    ]);
}

function markDocumentAccessVerified(PDO $pdo, int $userId): void
{
    $_SESSION['doc_access_user_id'] = $userId;
    $_SESSION['doc_access_verified_until'] = time() + (15 * 60);

    $stmt = $pdo->prepare('UPDATE user_document_security SET last_verified_at = CURRENT_TIMESTAMP WHERE user_id = :user_id');
    $stmt->execute(['user_id' => $userId]);
}

function clearDocumentAccessVerification(): void
{
    unset($_SESSION['doc_access_user_id'], $_SESSION['doc_access_verified_until']);
}

function isDocumentAccessVerified(int $userId): bool
{
    $verifiedUserId = (int) ($_SESSION['doc_access_user_id'] ?? 0);
    $verifiedUntil = (int) ($_SESSION['doc_access_verified_until'] ?? 0);

    return $verifiedUserId === $userId && $verifiedUntil >= time();
}

function enforceDocumentAccess(PDO $pdo, int $userId): void
{
    $security = getDocumentSecurity($pdo, $userId);
    if ($security === null || (int) $security['is_enabled'] !== 1) {
        return;
    }

    if (!isDocumentAccessVerified($userId)) {
        jsonResponse([
            'ok' => false,
            'message' => 'Debes verificar la clave secundaria para acceder a los documentos.',
            'requires_document_key' => true,
        ], 403);
    }
}

function detectInputFileType(string $extension): ?string
{
    $ext = strtolower($extension);

    // Devuelve la extensión real, no una categoría genérica
    if (in_array($ext, ['pdf', 'docx', 'txt', 'jpg', 'jpeg', 'png', 'mp3', 'wav'], true)) {
        // Normaliza 'jpeg' a 'jpg'
        return $ext === 'jpeg' ? 'jpg' : $ext;
    }

    return null;
}

function createDocumentGenerationJob(PDO $pdo, array $user, array $file, string $generationType, string $description, string $webhookUrl): array
{
    $originalName = substr((string) ($file['name'] ?? 'archivo'), 0, 255);
    $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
    $inputType = detectInputFileType($extension);

    if ($inputType === null) {
        throw new RuntimeException('Tipo de archivo no soportado: ' . $originalName);
    }

    $tmpPath = (string) ($file['tmp_name'] ?? '');
    $size = (int) ($file['size'] ?? 0);

    if ($tmpPath === '' || !is_uploaded_file($tmpPath)) {
        throw new RuntimeException('Archivo temporal invalido para ' . $originalName);
    }

    $rawContent = file_get_contents($tmpPath);
    if (!is_string($rawContent) || $rawContent === '') {
        throw new RuntimeException('No se pudo leer el contenido del archivo ' . $originalName);
    }

    $base64Content = base64_encode($rawContent);
    $hash = hash('sha256', $rawContent);

    $inputDirRelative = 'uploads/document-generation/input';
    $inputDirAbsolute = APP_ROOT . '/' . $inputDirRelative;
    if (!is_dir($inputDirAbsolute) && !mkdir($inputDirAbsolute, 0775, true) && !is_dir($inputDirAbsolute)) {
        throw new RuntimeException('No se pudo crear la carpeta de documentos de entrada.');
    }

    $safeName = preg_replace('/[^A-Za-z0-9._-]/', '_', $originalName) ?: 'archivo.' . $extension;
    $storedName = date('YmdHis') . '-' . bin2hex(random_bytes(4)) . '-' . $safeName;
    $storedRelative = $inputDirRelative . '/' . $storedName;
    $storedAbsolute = APP_ROOT . '/' . $storedRelative;

    if (!move_uploaded_file($tmpPath, $storedAbsolute)) {
        throw new RuntimeException('No se pudo guardar el archivo de entrada ' . $originalName);
    }

    $insert = $pdo->prepare(
        'INSERT INTO document_generation_jobs (
            user_id, user_email, input_file_name, input_file_type, input_file_size,
            input_file_path, input_file_hash, input_base64, generation_type, description,
            status, ip_address, user_agent
         ) VALUES (
            :user_id, :user_email, :input_file_name, :input_file_type, :input_file_size,
            :input_file_path, :input_file_hash, :input_base64, :generation_type, :description,
            :status, :ip_address, :user_agent
         )'
    );

    $insert->execute([
        'user_id' => (int) $user['id'],
        'user_email' => (string) ($user['email'] ?? ''),
        'input_file_name' => $originalName,
        'input_file_type' => $inputType,
        'input_file_size' => $size,
        'input_file_path' => $storedRelative,
        'input_file_hash' => $hash,
        'input_base64' => $base64Content,
        'generation_type' => $generationType,
        'description' => $description !== '' ? $description : null,
        'status' => 'pending',
        'ip_address' => $_SERVER['REMOTE_ADDR'] ?? null,
        'user_agent' => substr((string) ($_SERVER['HTTP_USER_AGENT'] ?? ''), 0, 500),
    ]);

    $jobId = (int) $pdo->lastInsertId();

    $pdo->prepare('UPDATE document_generation_jobs SET status = :status WHERE id = :id')->execute([
        'status' => 'processing',
        'id' => $jobId,
    ]);

    $callbackToken = getOrCreateCallbackToken($pdo);

    $payload = json_encode([
        'action'          => 'generate_manual',
        'job_id'          => $jobId,
        'file_name'       => $originalName,
        'file_type'       => $inputType,
        'file_content'    => $base64Content,
        'generation_type' => $generationType,
        'description'     => $description,
        'callback_url'    => getDocumentCallbackUrl(),
        'callback_token'  => $callbackToken,
        'user' => [
            'id'    => (int) $user['id'],
            'email' => (string) ($user['email'] ?? ''),
        ],
    ], JSON_THROW_ON_ERROR);

    $outbound = sendOutboundJsonRequest($webhookUrl, $payload, 45);

    $responseJson = [];
    if (!empty($outbound['body'])) {
        $decoded = json_decode((string) $outbound['body'], true);
        if (is_array($decoded)) {
            $responseJson = $decoded;
        }
    }

    $statusFromPA = strtolower((string) ($responseJson['status'] ?? ''));
    $isCompleted = $outbound['ok'] && $statusFromPA === 'ok';

    $outputUrl = (string) ($responseJson['file_url'] ?? '');
    $outputType = (string) ($responseJson['file_type'] ?? '');
    $outputName = $outputUrl !== '' ? basename(parse_url($outputUrl, PHP_URL_PATH) ?: '') : null;
    $paRequestId = (string) ($responseJson['request_id'] ?? $responseJson['id'] ?? '');

    $update = $pdo->prepare(
        'UPDATE document_generation_jobs
         SET pa_request_id = :pa_request_id,
             pa_status = :pa_status,
             pa_response = :pa_response,
             output_file_name = :output_file_name,
             output_file_type = :output_file_type,
             output_file_url = :output_file_url,
             status = :status,
             error_message = :error_message,
             completed_at = :completed_at,
             updated_at = CURRENT_TIMESTAMP
         WHERE id = :id'
    );

    $paStatus = $isCompleted ? 'ok' : ((string) ($outbound['status_line'] ?? 'error'));
    $errorMessage = $isCompleted
        ? null
        : buildPowerAutomateErrorMessage('Power Automate no pudo completar la generacion.', $outbound);

    $update->execute([
        'pa_request_id' => $paRequestId !== '' ? $paRequestId : null,
        'pa_status' => $paStatus,
        'pa_response' => !empty($outbound['body']) ? mb_substr((string) $outbound['body'], 0, 60000) : null,
        'output_file_name' => $outputName ?: null,
        'output_file_type' => $outputType !== '' ? $outputType : null,
        'output_file_url' => $outputUrl !== '' ? $outputUrl : null,
        'status' => $isCompleted ? 'completed' : 'error',
        'error_message' => $errorMessage,
        'completed_at' => $isCompleted ? date('Y-m-d H:i:s') : null,
        'id' => $jobId,
    ]);

    return [
        'id' => $jobId,
        'file_name' => $originalName,
        'status' => $isCompleted ? 'completed' : 'error',
        'output_file_url' => $outputUrl !== '' ? $outputUrl : null,
        'error_message' => $errorMessage,
    ];
}

function fetchDocumentJob(PDO $pdo, int $userId, int $jobId): ?array
{
    $stmt = $pdo->prepare('SELECT * FROM document_generation_jobs WHERE id = :id AND user_id = :user_id LIMIT 1');
    $stmt->execute([
        'id' => $jobId,
        'user_id' => $userId,
    ]);

    $row = $stmt->fetch();
    return is_array($row) ? $row : null;
}

/**
 * Envía la petición HTTP POST y cierra la conexión inmediatamente (fire-and-forget).
 * Usa un timeout de 2 s para que PHP no quede bloqueado esperando la respuesta de
 * Power Automate, evitando el error 504 en Power Apps.
 */
function fireAndForgetWebhookRequest(string $url, string $jsonPayload): void
{
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        if ($ch !== false) {
            curl_setopt_array($ch, [
                CURLOPT_RETURNTRANSFER => false,
                CURLOPT_POST           => true,
                CURLOPT_HTTPHEADER     => [
                    'Content-Type: application/json',
                    'Expect:',
                ],
                CURLOPT_POSTFIELDS     => $jsonPayload,
                // 2s puede cortar cargas base64 grandes y dejar el flujo sin datos.
                // Con 15s damos margen para enviar el body completo sin volver al modo sincrono largo.
                CURLOPT_TIMEOUT        => 15,
                CURLOPT_CONNECTTIMEOUT => 5,
            ]);
            @curl_exec($ch);
            curl_close($ch);
        }
        return;
    }

    // Fallback stream cuando cURL no está disponible.
    $ctx = stream_context_create([
        'http' => [
            'method'  => 'POST',
            'header'  => "Content-Type: application/json\r\n",
            'content' => $jsonPayload,
            'timeout' => 15,
        ],
    ]);
    @file_get_contents($url, false, $ctx);
}

function sendOutboundJsonRequest(string $url, string $jsonPayload, int $timeoutSeconds = 15): array
{
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        if ($ch !== false) {
            curl_setopt_array($ch, [
                CURLOPT_RETURNTRANSFER => true,
                CURLOPT_POST => true,
                CURLOPT_HTTPHEADER => ['Content-Type: application/json'],
                CURLOPT_POSTFIELDS => $jsonPayload,
                CURLOPT_TIMEOUT => $timeoutSeconds,
            ]);

            $body = curl_exec($ch);
            $curlError = curl_error($ch);
            $httpCode = (int) curl_getinfo($ch, CURLINFO_HTTP_CODE);
            curl_close($ch);

            return [
                'ok' => $httpCode >= 200 && $httpCode < 300,
                'status_line' => $httpCode > 0 ? ('HTTP ' . $httpCode) : '',
                'body' => is_string($body) ? trim($body) : '',
                'transport_error' => trim($curlError),
            ];
        }
    }

    $ctx = stream_context_create([
        'http' => [
            'method' => 'POST',
            'header' => "Content-Type: application/json\r\n",
            'content' => $jsonPayload,
            'timeout' => $timeoutSeconds,
            'ignore_errors' => true,
        ],
    ]);

    $body = @file_get_contents($url, false, $ctx);
    $headers = $http_response_header ?? [];

    return [
        'ok' => is_string($headers[0] ?? '') && preg_match('/HTTP\/\S+ 2/', (string) $headers[0]) === 1,
        'status_line' => (string) ($headers[0] ?? ''),
        'body' => is_string($body) ? trim($body) : '',
        'transport_error' => is_string($body) ? '' : 'Sin respuesta HTTP desde la URL configurada.',
    ];
}

function buildPowerAutomateErrorMessage(string $baseMessage, array $result): string
{
    $details = [];

    if (!empty($result['status_line'])) {
        $details[] = (string) $result['status_line'];
    }

    if (!empty($result['transport_error'])) {
        $details[] = (string) $result['transport_error'];
    }

    if (!empty($result['body'])) {
        $details[] = mb_substr((string) $result['body'], 0, 250);
    }

    return $details === [] ? $baseMessage : ($baseMessage . ' Detalle: ' . implode(' | ', $details));
}

