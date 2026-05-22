<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

function publicProfileJsonResponse(array $payload, int $status = 200): void
{
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    $slug = (string) ($_GET['slug'] ?? env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo'));
    $profile = getPublicProfile($slug);

    if ($profile === null) {
        jsonResponse([
            'ok' => false,
            'message' => 'Perfil publico no encontrado.'
        ], 404);
    }

    jsonResponse([
        'ok' => true,
        'profile' => $profile,
    ]);
} catch (Throwable $throwable) {
    $message = function_exists('applicationErrorMessage')
        ? applicationErrorMessage($throwable)
        : ('No se pudo iniciar public_profile.php: ' . $throwable->getMessage());

    publicProfileJsonResponse([
        'ok' => false,
        'message' => $message,
    ], 500);
}