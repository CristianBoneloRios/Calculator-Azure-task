<?php

declare(strict_types=1);

require_once __DIR__ . '/bootstrap.php';

function db(): PDO
{
    static $pdo = null;

    if ($pdo instanceof PDO) {
        return $pdo;
    }

    $driver = env('DB_CONNECTION', 'mysql');
    $host = env('DB_HOST', 'localhost');
    $port = env('DB_PORT', '3306');
    $database = env('DB_DATABASE', '');
    $charset = env('DB_CHARSET', 'utf8mb4');
    $username = env('DB_USERNAME', '');
    $password = env('DB_PASSWORD', '');

    if ($database === '' || $username === '') {
        throw new RuntimeException('Configuracion de base de datos incompleta. Revisa DB_DATABASE y DB_USERNAME en .env o api/config.php.');
    }

    $dsn = sprintf('%s:host=%s;port=%s;dbname=%s;charset=%s', $driver, $host, $port, $database, $charset);

    try {
        $pdo = new PDO($dsn, $username, $password, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES => false,
        ]);
    } catch (Throwable $throwable) {
        throw new RuntimeException(
            sprintf(
                'No se pudo conectar a la base de datos. Verifica DB_HOST (%s), DB_PORT (%s), DB_DATABASE (%s), DB_USERNAME (%s) y que el servidor permita conexiones MySQL.',
                $host,
                $port,
                $database,
                $username
            ),
            0,
            $throwable
        );
    }

    return $pdo;
}