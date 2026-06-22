<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    $origin = (string) ($_SERVER['HTTP_ORIGIN'] ?? '*');
    header('Access-Control-Allow-Origin: ' . $origin);
    header('Vary: Origin');
    header('Access-Control-Allow-Headers: Content-Type, Authorization, X-API-Key');
    header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
    http_response_code(204);
    exit;
}

try {
    require_once __DIR__ . '/app.php';
    ensureApplicationInstalled();

    allowCopilotCors();
    requireCopilotApiKey();

    $input = jsonInput();
    $action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? 'capabilities'));

    switch ($action) {
        case 'capabilities':
            jsonResponse([
                'ok' => true,
                'service' => 'Azure Task Suite Copilot Agent API',
                'actions' => [
                    'capabilities',
                    'validate_system',
                    'get_user_snapshot',
                    'search_user_content',
                ],
                'version' => '1.0.0',
            ]);
            break;

        case 'validate_system':
            jsonResponse([
                'ok' => true,
                'timestamp' => date('c'),
                'database' => validateDatabase(),
                'tables' => validateRequiredTables(),
                'stats' => systemStats(),
            ]);
            break;

        case 'get_user_snapshot':
            $email = strtolower(trim((string) ($input['email'] ?? $_GET['email'] ?? '')));
            if ($email === '' || !filter_var($email, FILTER_VALIDATE_EMAIL)) {
                jsonResponse([
                    'ok' => false,
                    'message' => 'Debes enviar un email valido.',
                ], 422);
            }

            $limit = max(1, min(50, (int) ($input['limit'] ?? $_GET['limit'] ?? 10)));
            $user = findUserByEmail($email);
            if ($user === null) {
                jsonResponse([
                    'ok' => false,
                    'message' => 'Usuario no encontrado.',
                ], 404);
            }

            $userId = (int) $user['id'];
            jsonResponse([
                'ok' => true,
                'user' => sanitizeUser($user),
                'summary' => userSummary($userId),
                'recent' => [
                    'tasks' => fetchAllRows(
                        'SELECT id, task_date, title, status, priority, due_time, created_at FROM daily_tasks WHERE user_id = :user_id ORDER BY task_date DESC, id DESC LIMIT ' . $limit,
                        ['user_id' => $userId]
                    ),
                    'goals' => fetchAllRows(
                        'SELECT id, title, status, progress_percent, target_date, created_at FROM goals WHERE user_id = :user_id ORDER BY id DESC LIMIT ' . $limit,
                        ['user_id' => $userId]
                    ),
                    'notes' => fetchAllRows(
                        'SELECT id, title, color, is_pinned, updated_at FROM notes WHERE user_id = :user_id ORDER BY updated_at DESC LIMIT ' . $limit,
                        ['user_id' => $userId]
                    ),
                    'events' => fetchAllRows(
                        'SELECT id, title, start_at, end_at, source_type, location FROM calendar_events WHERE user_id = :user_id ORDER BY start_at DESC LIMIT ' . $limit,
                        ['user_id' => $userId]
                    ),
                ],
            ]);
            break;

        case 'search_user_content':
            $email = strtolower(trim((string) ($input['email'] ?? $_GET['email'] ?? '')));
            $query = trim((string) ($input['query'] ?? $_GET['query'] ?? ''));
            $entity = strtolower(trim((string) ($input['entity'] ?? $_GET['entity'] ?? 'all')));
            $limit = max(1, min(25, (int) ($input['limit'] ?? $_GET['limit'] ?? 10)));

            if ($email === '' || !filter_var($email, FILTER_VALIDATE_EMAIL)) {
                jsonResponse(['ok' => false, 'message' => 'Debes enviar un email valido.'], 422);
            }
            if ($query === '') {
                jsonResponse(['ok' => false, 'message' => 'Debes enviar un texto de busqueda.'], 422);
            }

            $user = findUserByEmail($email);
            if ($user === null) {
                jsonResponse(['ok' => false, 'message' => 'Usuario no encontrado.'], 404);
            }

            $userId = (int) $user['id'];
            $like = '%' . $query . '%';
            $allowedEntities = ['all', 'tasks', 'goals', 'notes', 'events'];
            if (!in_array($entity, $allowedEntities, true)) {
                jsonResponse(['ok' => false, 'message' => 'Entidad no valida.'], 422);
            }

            $results = [];

            if ($entity === 'all' || $entity === 'tasks') {
                $results['tasks'] = fetchAllRows(
                    'SELECT id, task_date, title, description, status, priority FROM daily_tasks WHERE user_id = :user_id AND (title LIKE :q OR description LIKE :q) ORDER BY task_date DESC, id DESC LIMIT ' . $limit,
                    ['user_id' => $userId, 'q' => $like]
                );
            }

            if ($entity === 'all' || $entity === 'goals') {
                $results['goals'] = fetchAllRows(
                    'SELECT id, title, description, status, progress_percent, target_date FROM goals WHERE user_id = :user_id AND (title LIKE :q OR description LIKE :q) ORDER BY id DESC LIMIT ' . $limit,
                    ['user_id' => $userId, 'q' => $like]
                );
            }

            if ($entity === 'all' || $entity === 'notes') {
                $results['notes'] = fetchAllRows(
                    'SELECT id, title, content, color, is_pinned, updated_at FROM notes WHERE user_id = :user_id AND (title LIKE :q OR content LIKE :q) ORDER BY updated_at DESC LIMIT ' . $limit,
                    ['user_id' => $userId, 'q' => $like]
                );
            }

            if ($entity === 'all' || $entity === 'events') {
                $results['events'] = fetchAllRows(
                    'SELECT id, title, description, start_at, end_at, source_type, location FROM calendar_events WHERE user_id = :user_id AND (title LIKE :q OR description LIKE :q OR location LIKE :q) ORDER BY start_at DESC LIMIT ' . $limit,
                    ['user_id' => $userId, 'q' => $like]
                );
            }

            jsonResponse([
                'ok' => true,
                'user' => [
                    'id' => (int) $user['id'],
                    'full_name' => (string) $user['full_name'],
                    'email' => (string) $user['email'],
                ],
                'query' => $query,
                'entity' => $entity,
                'results' => $results,
            ]);
            break;

        default:
            jsonResponse([
                'ok' => false,
                'message' => 'Accion no soportada.',
            ], 400);
    }
} catch (Throwable $throwable) {
    jsonResponse([
        'ok' => false,
        'message' => applicationErrorMessage($throwable),
    ], 500);
}

function allowCopilotCors(): void
{
    $origin = (string) ($_SERVER['HTTP_ORIGIN'] ?? '*');
    $allowedOrigin = (string) env('COPILOT_ALLOWED_ORIGIN', '*');

    if ($allowedOrigin === '*' || $allowedOrigin === $origin) {
        header('Access-Control-Allow-Origin: ' . ($allowedOrigin === '*' ? '*' : $origin));
    }

    header('Vary: Origin');
}

function requireCopilotApiKey(): void
{
    $expected = (string) (env('COPILOT_AGENT_API_KEY', '') ?: (defined('COPILOT_AGENT_API_KEY') ? (string) COPILOT_AGENT_API_KEY : ''));
    if ($expected === '') {
        jsonResponse([
            'ok' => false,
            'message' => 'COPILOT_AGENT_API_KEY no configurada en el servidor.',
        ], 503);
    }

    $received = extractApiKeyFromRequest();
    if ($received === '' || !hash_equals($expected, $received)) {
        jsonResponse([
            'ok' => false,
            'message' => 'API key invalida.',
        ], 401);
    }
}

function extractApiKeyFromRequest(): string
{
    $headerApiKey = (string) ($_SERVER['HTTP_X_API_KEY'] ?? '');
    if ($headerApiKey !== '') {
        return trim($headerApiKey);
    }

    $auth = (string) ($_SERVER['HTTP_AUTHORIZATION'] ?? '');
    if ($auth !== '' && stripos($auth, 'Bearer ') === 0) {
        return trim(substr($auth, 7));
    }

    return '';
}

function validateDatabase(): array
{
    try {
        $value = fetchScalarValue('SELECT 1');
        return [
            'connected' => ((int) $value) === 1,
            'message' => 'Conexion a base de datos operativa.',
        ];
    } catch (Throwable $throwable) {
        return [
            'connected' => false,
            'message' => $throwable->getMessage(),
        ];
    }
}

function validateRequiredTables(): array
{
    $requiredTables = [
        'users',
        'user_sessions',
        'notes',
        'daily_tasks',
        'goals',
        'calendar_events',
        'public_profiles',
    ];

    $output = [];
    foreach ($requiredTables as $tableName) {
        $output[$tableName] = tableExists(db(), $tableName);
    }

    return $output;
}

function systemStats(): array
{
    return [
        'users' => (int) fetchScalarValue('SELECT COUNT(*) FROM users'),
        'users_2fa_enabled' => (int) fetchScalarValue('SELECT COUNT(*) FROM users WHERE two_factor_enabled = 1'),
        'notes' => (int) fetchScalarValue('SELECT COUNT(*) FROM notes'),
        'tasks' => (int) fetchScalarValue('SELECT COUNT(*) FROM daily_tasks'),
        'goals' => (int) fetchScalarValue('SELECT COUNT(*) FROM goals'),
        'events' => (int) fetchScalarValue('SELECT COUNT(*) FROM calendar_events'),
    ];
}

function userSummary(int $userId): array
{
    return [
        'tasks_total' => (int) fetchScalarValue('SELECT COUNT(*) FROM daily_tasks WHERE user_id = :user_id', ['user_id' => $userId]),
        'tasks_done' => (int) fetchScalarValue('SELECT COUNT(*) FROM daily_tasks WHERE user_id = :user_id AND status = :status', ['user_id' => $userId, 'status' => 'done']),
        'goals_total' => (int) fetchScalarValue('SELECT COUNT(*) FROM goals WHERE user_id = :user_id', ['user_id' => $userId]),
        'notes_total' => (int) fetchScalarValue('SELECT COUNT(*) FROM notes WHERE user_id = :user_id', ['user_id' => $userId]),
        'events_total' => (int) fetchScalarValue('SELECT COUNT(*) FROM calendar_events WHERE user_id = :user_id', ['user_id' => $userId]),
    ];
}

function findUserByEmail(string $email): ?array
{
    $stmt = db()->prepare('SELECT * FROM users WHERE email = :email LIMIT 1');
    $stmt->execute(['email' => $email]);
    $row = $stmt->fetch();

    return is_array($row) ? $row : null;
}

function fetchScalarValue(string $sql, array $params = [])
{
    $stmt = db()->prepare($sql);
    $stmt->execute($params);

    return $stmt->fetchColumn();
}

function fetchAllRows(string $sql, array $params = []): array
{
    $stmt = db()->prepare($sql);
    $stmt->execute($params);
    $rows = $stmt->fetchAll();

    return is_array($rows) ? $rows : [];
}
