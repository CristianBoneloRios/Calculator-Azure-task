<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    $input = jsonInput();
    $action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? 'session'));

    if ($action === 'generate-2fa') {
        $secret = generateTwoFactorSecret();
        $email = (string) ($input['email'] ?? '');

        if (empty($email)) {
            jsonResponse([
                'ok' => false,
                'message' => 'Email requerido para generar 2FA.'
            ], 422);
        }

        // Generar URL para Google Authenticator
        $label = urlencode("Azure Task Suite ({$email})");
        $issuer = urlencode('Azure Task Suite');
        $otpauth_url = "otpauth://totp/{$label}?secret={$secret}&issuer={$issuer}";

        // Almacenar temporalmente el secret en la sesión
        $_SESSION['_2fa_secret_temp'] = $secret;
        $_SESSION['_2fa_email_temp'] = $email;

        jsonResponse([
            'ok' => true,
            'secret' => $secret,
            'otpauth_url' => $otpauth_url,
            'message' => 'Secret 2FA generado. Escanea el código QR con tu autenticador.'
        ]);
    }

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

    if ($action !== 'login' && $action !== 'register') {
        jsonResponse([
            'ok' => false,
            'message' => 'Accion no soportada.'
        ], 400);
    }

    if ($action === 'register') {
        $fullName = (string) ($input['full_name'] ?? $_POST['full_name'] ?? '');
        $email = (string) ($input['email'] ?? $_POST['email'] ?? '');
        $password = (string) ($input['password'] ?? $_POST['password'] ?? '');
        $confirmPassword = (string) ($input['confirm_password'] ?? $_POST['confirm_password'] ?? '');

        if (trim($fullName) === '' || trim($email) === '' || trim($password) === '') {
            jsonResponse([
                'ok' => false,
                'message' => 'Nombre, correo y contrasena son obligatorios.'
            ], 422);
        }

        if ($password !== $confirmPassword) {
            jsonResponse([
                'ok' => false,
                'message' => 'Las contrasenas no coinciden.'
            ], 422);
        }

        try {
            $user = registerUserAccount($fullName, $email, $password);
        } catch (InvalidArgumentException|RuntimeException $registerError) {
            jsonResponse([
                'ok' => false,
                'message' => $registerError->getMessage(),
            ], 422);
        }

        jsonResponse([
            'ok' => true,
            'authenticated' => false,
            'user' => sanitizeUser($user),
            'redirect_to' => 'login.php',
            'message' => 'Cuenta creada correctamente. Ahora puedes iniciar sesion.'
        ]);
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

    // Verificar si el usuario tiene 2FA habilitado
    if ($user['two_factor_enabled'] && !empty($user['two_factor_secret'])) {
        $_SESSION['_2fa_user_id'] = $user['id'];
        $_SESSION['_2fa_secret'] = $user['two_factor_secret'];

        jsonResponse([
            'ok' => true,
            'requires_2fa' => true,
            'message' => 'Verifica tu codigo 2FA.'
        ]);
    }

    // Login normal sin 2FA
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

    jsonResponse([
        'ok' => false,
        'message' => $message,
    ], 500);
}