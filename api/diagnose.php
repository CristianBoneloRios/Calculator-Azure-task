<?php
/**
 * Diagnóstico Ultra-Simple
 * No depende de includes complejos, solo reporta datos raw
 */

header('Content-Type: application/json; charset=utf-8');
http_response_code(200);

$diagnostics = [];

// 1. PHP Version
$diagnostics['php_version'] = PHP_VERSION;

// 2. System Info
$diagnostics['system'] = php_uname();

// 3. Current directory
$apiDir = __DIR__;
$publicHtmlDir = dirname($apiDir);
$appRoot = $publicHtmlDir;

$diagnostics['paths'] = [
    'api_dir' => $apiDir,
    'public_html_dir' => $publicHtmlDir,
    'app_root' => $appRoot,
];

// 4. Check if .env exists (en public_html o arriba)
$envPaths = [
    $publicHtmlDir . '/.env',
    dirname($publicHtmlDir) . '/.env',
];
$envPath = null;
foreach ($envPaths as $path) {
    if (file_exists($path)) {
        $envPath = $path;
        break;
    }
}

$diagnostics['env_file'] = [
    'checked_paths' => $envPaths,
    'exists' => $envPath !== null,
    'found_at' => $envPath,
    'readable' => $envPath && is_readable($envPath),
];

// 5. Try to read .env content
if ($envPath && is_readable($envPath)) {
    $envContent = file_get_contents($envPath);
    $envLines = array_filter(
        array_map('trim', explode("\n", $envContent)),
        fn($line) => $line && !str_starts_with($line, '#')
    );
    $diagnostics['env_vars_count'] = count($envLines);
    
    // Extract only non-sensitive vars
    $diagnostics['env_sample'] = [];
    foreach ($envLines as $line) {
        if (strpos($line, '=') !== false) {
            [$key, $val] = explode('=', $line, 2);
            $key = trim($key);
            if (!in_array($key, ['DB_PASSWORD'])) {
                $diagnostics['env_sample'][$key] = trim($val, '\'"');
            }
        }
    }
}

// 6. Check PHP extensions
$extensions = ['pdo', 'pdo_mysql', 'curl', 'json'];
$diagnostics['extensions'] = [];
foreach ($extensions as $ext) {
    $diagnostics['extensions'][$ext] = extension_loaded($ext);
}

// 7. Try to connect to DB
$diagnostics['db_connection'] = ['status' => 'not_attempted'];
try {
    if ($envPath && file_exists($envPath)) {
        // Parse .env manually
        $lines = file($envPath, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
        $env = [];
        foreach ($lines as $line) {
            if ($line && !str_starts_with(trim($line), '#')) {
                $parts = explode('=', $line, 2);
                if (count($parts) === 2) {
                    $key = trim($parts[0]);
                    $val = trim($parts[1], '\'"');
                    $env[$key] = $val;
                }
            }
        }
        
        $host = $env['DB_HOST'] ?? 'localhost';
        $port = $env['DB_PORT'] ?? '3306';
        $database = $env['DB_DATABASE'] ?? '';
        $username = $env['DB_USERNAME'] ?? '';
        $password = $env['DB_PASSWORD'] ?? '';
        $charset = $env['DB_CHARSET'] ?? 'utf8mb4';
        
        $dsn = "mysql:host=$host;port=$port;dbname=$database;charset=$charset";
        
        try {
            $pdo = new PDO($dsn, $username, $password, [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            ]);
            $diagnostics['db_connection'] = [
                'status' => 'success',
                'host' => $host,
                'database' => $database,
                'version' => $pdo->getAttribute(PDO::ATTR_SERVER_VERSION),
            ];
        } catch (PDOException $e) {
            $diagnostics['db_connection'] = [
                'status' => 'failed',
                'error' => $e->getMessage(),
                'host' => $host,
                'database' => $database,
                'username' => $username,
            ];
        }
    }
} catch (Throwable $e) {
    $diagnostics['db_connection'] = [
        'status' => 'error',
        'error' => $e->getMessage(),
    ];
}

// 8. Check important files
$files = [
    'api/bootstrap.php',
    'api/db.php',
    'api/app.php',
    'api/health.php',
    'api/auth.php',
    'index.html',
    'index.php',
    '.htaccess',
];
$diagnostics['files'] = [];
foreach ($files as $file) {
    $path = $publicHtmlDir . '/' . $file;
    $diagnostics['files'][$file] = [
        'path' => $path,
        'exists' => file_exists($path),
        'readable' => is_readable($path),
        'size' => file_exists($path) ? filesize($path) : 0,
    ];
}

// 9. Server variables
$diagnostics['server_info'] = [
    'request_method' => $_SERVER['REQUEST_METHOD'] ?? 'N/A',
    'script_name' => $_SERVER['SCRIPT_NAME'] ?? 'N/A',
    'remote_addr' => $_SERVER['REMOTE_ADDR'] ?? 'N/A',
    'server_port' => $_SERVER['SERVER_PORT'] ?? 'N/A',
    'https' => isset($_SERVER['HTTPS']) ? 'yes' : 'no',
];

// 10. Try to load bootstrap and report errors
$diagnostics['bootstrap'] = ['status' => 'not_attempted'];
try {
    // Silence is golden - we're just checking if it works
    ob_start();
    require_once $apiDir . '/bootstrap.php';
    ob_end_clean();
    $diagnostics['bootstrap'] = ['status' => 'success'];
} catch (Throwable $e) {
    $diagnostics['bootstrap'] = [
        'status' => 'error',
        'error' => $e->getMessage(),
        'file' => $e->getFile(),
        'line' => $e->getLine(),
    ];
}

echo json_encode($diagnostics, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
