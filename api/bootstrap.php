<?php

declare(strict_types=1);

if (!function_exists('str_starts_with')) {
    function str_starts_with(string $haystack, string $needle): bool
    {
        return $needle === '' || strpos($haystack, $needle) === 0;
    }
}

if (!function_exists('str_ends_with')) {
    function str_ends_with(string $haystack, string $needle): bool
    {
        if ($needle === '') {
            return true;
        }

        return substr($haystack, -strlen($needle)) === $needle;
    }
}

const APP_ROOT = __DIR__ . '/..';

$configPath = __DIR__ . '/config.php';
if (is_file($configPath)) {
    require_once $configPath;
}

loadEnv(APP_ROOT . '/.env');
date_default_timezone_set(env('APP_TIMEZONE', 'UTC'));

if (session_status() !== PHP_SESSION_ACTIVE) {
    $sessionName = env('APP_SESSION_NAME', 'app_session');
    if ($sessionName !== '') {
        session_name($sessionName);
    }
    session_start();
}

function loadEnv(string $path): void
{
    if (!is_file($path) || !is_readable($path)) {
        return;
    }

    $lines = file($path, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
    if ($lines === false) {
        return;
    }

    foreach ($lines as $line) {
        $trimmed = trim($line);
        if ($trimmed === '' || str_starts_with($trimmed, '#')) {
            continue;
        }

        $parts = explode('=', $trimmed, 2);
        if (count($parts) !== 2) {
            continue;
        }

        $name = trim($parts[0]);
        $value = trim($parts[1]);

        if ($name === '') {
            continue;
        }

        if ((str_starts_with($value, '"') && str_ends_with($value, '"')) || (str_starts_with($value, "'") && str_ends_with($value, "'"))) {
            $value = substr($value, 1, -1);
        }

        $_ENV[$name] = $value;
        $_SERVER[$name] = $value;
        putenv($name . '=' . $value);
    }
}

function env(string $key, ?string $default = null): ?string
{
    $value = $_ENV[$key] ?? $_SERVER[$key] ?? getenv($key);

    if (($value === false || $value === null || $value === '') && defined($key)) {
        $value = constant($key);
    }

    if ($value === false || $value === null || $value === '') {
        return $default;
    }

    return (string) $value;
}

if (!function_exists('jsonResponse')) {
    function jsonResponse(array $payload, int $status = 200): void
    {
        http_response_code($status);
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
        exit;
    }
}

if (!function_exists('jsonInput')) {
    function jsonInput(): array
    {
        $raw = file_get_contents('php://input');
        if (!is_string($raw) || trim($raw) === '') {
            return [];
        }

        $data = json_decode($raw, true);
        return is_array($data) ? $data : [];
    }
}

function applicationErrorMessage(Throwable $throwable): string
{
    $message = $throwable->getMessage();

    $looksLikeDatabaseError =
        stripos($message, 'base de datos') !== false ||
        stripos($message, 'db_') !== false ||
        stripos($message, 'pdo') !== false ||
        stripos($message, 'sqlstate') !== false ||
        stripos($message, 'mysql') !== false ||
        stripos($message, 'unknown column') !== false ||
        stripos($message, 'table') !== false ||
        stripos($message, 'column') !== false ||
        stripos($message, 'foreign key') !== false ||
        stripos($message, 'information_schema') !== false;

    if ($looksLikeDatabaseError) {
        return 'No se pudo inicializar la aplicacion porque la conexion a la base de datos no esta lista. En Hostinger verifica que subiste el archivo .env, que DB_HOST sea el hostname real de MySQL y que PDO MySQL este habilitado.';
    }

    return 'La aplicacion no pudo iniciar correctamente en el servidor. Revisa la configuracion de PHP y de la base de datos en Hostinger.';
}
