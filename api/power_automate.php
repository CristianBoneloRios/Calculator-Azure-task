<?php

declare(strict_types=1);

require_once __DIR__ . '/app.php';

header('Content-Type: application/json; charset=utf-8');

try {
    ensureApplicationInstalled();

    if (($_SERVER['REQUEST_METHOD'] ?? 'GET') !== 'POST') {
        jsonResponse([
            'ok' => false,
            'message' => 'Metodo no permitido.',
        ], 405);
    }

    $token = trim((string) ($_SERVER['HTTP_X_POWER_AUTOMATE_KEY'] ?? ''));
    if ($token === '') {
        $authorization = trim((string) ($_SERVER['HTTP_AUTHORIZATION'] ?? ''));
        if (stripos($authorization, 'Bearer ') === 0) {
            $token = trim(substr($authorization, 7));
        }
    }

    if ($token === '') {
        jsonResponse([
            'ok' => false,
            'message' => 'Falta el encabezado de autenticacion de Power Automate.',
        ], 401);
    }

    $integration = findPowerAutomateIntegrationByToken($token);
    if ($integration === null) {
        jsonResponse([
            'ok' => false,
            'message' => 'Token de integracion invalido.',
        ], 401);
    }

    $payload = jsonInput();
    $result = syncPowerAutomateDailySessions($integration, $payload);

    jsonResponse([
        'ok' => true,
        'message' => 'Sincronizacion procesada.',
        'result' => $result,
    ]);
} catch (InvalidArgumentException $exception) {
    jsonResponse([
        'ok' => false,
        'message' => $exception->getMessage(),
    ], 422);
} catch (Throwable $throwable) {
    jsonResponse([
        'ok' => false,
        'message' => applicationErrorMessage($throwable),
    ], 500);
}