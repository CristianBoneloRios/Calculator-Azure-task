<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
        http_response_code(405);
        echo json_encode(['ok' => false, 'message' => 'Metodo no permitido.']);
        exit;
    }

    // Autenticacion con token secreto enviado por Power Automate.
    $receivedToken = trim((string) ($_SERVER['HTTP_X_CALLBACK_TOKEN'] ?? ''));
    if ($receivedToken === '') {
        http_response_code(401);
        echo json_encode(['ok' => false, 'message' => 'Token de callback requerido.']);
        exit;
    }

    $pdo = db();

    // Recuperar el secreto almacenado en app_settings (global, no por usuario).
    $stmt = $pdo->prepare(
        'SELECT setting_value FROM app_settings WHERE setting_key = :key LIMIT 1'
    );
    $stmt->execute(['key' => 'document_generation_callback_secret']);
    $storedToken = (string) ($stmt->fetchColumn() ?: '');

    if ($storedToken === '' || !hash_equals($storedToken, $receivedToken)) {
        http_response_code(401);
        echo json_encode(['ok' => false, 'message' => 'Token de callback invalido.']);
        exit;
    }

    // Leer cuerpo JSON enviado por Power Automate.
    $raw = (string) (file_get_contents('php://input') ?: '');
    $input = [];
    if ($raw !== '') {
        $decoded = json_decode($raw, true);
        if (is_array($decoded)) {
            $input = $decoded;
        }
    }

    $jobId     = (int) ($input['job_id'] ?? 0);
    $status    = strtolower(trim((string) ($input['status'] ?? '')));
    $fileUrl   = trim((string) ($input['file_url'] ?? ''));
    $fileType  = trim((string) ($input['file_type'] ?? ''));
    $paReqId   = trim((string) ($input['request_id'] ?? $input['id'] ?? ''));
    $errorMsg  = trim((string) ($input['error_message'] ?? ''));

    if ($jobId <= 0) {
        http_response_code(422);
        echo json_encode(['ok' => false, 'message' => 'job_id invalido o faltante.']);
        exit;
    }

    if (!in_array($status, ['completed', 'error'], true)) {
        http_response_code(422);
        echo json_encode(['ok' => false, 'message' => 'status debe ser "completed" o "error".']);
        exit;
    }

    // Verificar que el job existe.
    $checkStmt = $pdo->prepare(
        'SELECT id FROM document_generation_jobs WHERE id = :id LIMIT 1'
    );
    $checkStmt->execute(['id' => $jobId]);
    if ($checkStmt->fetchColumn() === false) {
        http_response_code(404);
        echo json_encode(['ok' => false, 'message' => 'Job no encontrado.']);
        exit;
    }

    $isCompleted = $status === 'completed';
    $outputName  = ($isCompleted && $fileUrl !== '')
        ? (basename(parse_url($fileUrl, PHP_URL_PATH) ?: '') ?: null)
        : null;

    $update = $pdo->prepare(
        'UPDATE document_generation_jobs
         SET pa_request_id  = COALESCE(:pa_request_id, pa_request_id),
             pa_status       = :pa_status,
             output_file_name = :output_file_name,
             output_file_type = :output_file_type,
             output_file_url  = :output_file_url,
             status           = :status,
             error_message    = :error_message,
             completed_at     = :completed_at,
             updated_at       = CURRENT_TIMESTAMP
         WHERE id = :id'
    );

    $update->execute([
        'pa_request_id'   => $paReqId !== '' ? $paReqId : null,
        'pa_status'       => $isCompleted ? 'ok' : 'error',
        'output_file_name'=> $outputName,
        'output_file_type'=> $fileType !== '' ? $fileType : null,
        'output_file_url' => $fileUrl !== '' ? $fileUrl : null,
        'status'          => $isCompleted ? 'completed' : 'error',
        'error_message'   => !$isCompleted && $errorMsg !== '' ? $errorMsg : null,
        'completed_at'    => $isCompleted ? date('Y-m-d H:i:s') : null,
        'id'              => $jobId,
    ]);

    http_response_code(200);
    echo json_encode(['ok' => true, 'message' => 'Job actualizado.', 'job_id' => $jobId]);

} catch (Throwable $throwable) {
    error_log('document_callback.php error: ' . $throwable->getMessage());
    http_response_code(500);
    echo json_encode(['ok' => false, 'message' => 'Error interno del servidor.']);
}
