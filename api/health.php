<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

function healthJsonResponse(array $payload, int $status = 200): void
{
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

try {
    require_once __DIR__ . '/app.php';
} catch (Throwable $throwable) {
    healthJsonResponse([
        'ok' => false,
        'checks' => [
            [
                'name' => 'app_bootstrap',
                'ok' => false,
                'details' => $throwable->getMessage(),
            ],
        ],
        'timestamp' => date('c'),
    ], 500);
}

$checks = [];

$checks[] = [
    'name' => 'php_version',
    'ok' => true,
    'details' => PHP_VERSION,
];

$checks[] = [
    'name' => 'pdo_mysql_extension',
    'ok' => extension_loaded('pdo_mysql'),
    'details' => extension_loaded('pdo_mysql') ? 'loaded' : 'missing',
];

$envPath = APP_ROOT . '/.env';
$checks[] = [
    'name' => 'env_file',
    'ok' => is_file($envPath),
    'details' => is_file($envPath) ? 'present' : 'missing',
];

try {
    ensureApplicationInstalled();
    $db = db();
    $stmt = $db->query('SELECT 1');
    $dbCheck = $stmt !== false;
    $checks[] = [
        'name' => 'database_connection',
        'ok' => $dbCheck,
        'details' => $dbCheck ? 'ok' : 'query_failed',
    ];
} catch (Throwable $throwable) {
    $checks[] = [
        'name' => 'database_connection',
        'ok' => false,
        'details' => $throwable->getMessage(),
    ];
}

$allOk = true;
foreach ($checks as $check) {
    if (!$check['ok']) {
        $allOk = false;
        break;
    }
}

jsonResponse([
    'ok' => $allOk,
    'checks' => $checks,
    'timestamp' => date('c'),
]);
