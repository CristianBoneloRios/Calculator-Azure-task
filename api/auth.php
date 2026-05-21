<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

function authJsonResponse(array $payload, int $status = 200): void
{
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    $input = jsonInput();
    $action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? 'session'));

    if ($action === 'session') {
        $user = currentUser();
        jsonResponse([
            'ok' => true,
            'authenticated' => $user !== null,
            'user' => $user,
        ]);
    }

    if ($action === 'logout') {
        logoutCurrentUser();
        jsonResponse([
            'ok' => true,
            'message' => 'Sesion cerrada correctamente.'
        ]);
    }

    if ($action !== 'login') {
        jsonResponse([
            'ok' => false,
            'message' => 'Accion no soportada.'
        ], 400);
    }

    $email = (string) ($input['email'] ?? $_POST['email'] ?? '');
    $password = (string) ($input['password'] ?? $_POST['password'] ?? '');

    if (trim($email) === '' || trim($password) === '') {
        jsonResponse([
            'ok' => false,
            'message' => 'Correo y contrasena son obligatorios.'
        ], 422);
    }

    $user = authenticateUser($email, $password);
    if ($user === null) {
        jsonResponse([
            'ok' => false,
            'message' => 'Credenciales invalidas.'
        ], 401);
    }

    $authenticatedUser = startAuthenticatedSession($user);

    jsonResponse([
        'ok' => true,
        'authenticated' => true,
        'user' => $authenticatedUser,
    ]);
} catch (Throwable $throwable) {
    $message = function_exists('applicationErrorMessage')
        ? applicationErrorMessage($throwable)
        : ('No se pudo iniciar auth.php: ' . $throwable->getMessage());

    authJsonResponse([
        'ok' => false,
        'message' => $message,
    ], 500);
}