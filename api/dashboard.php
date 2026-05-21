<?php

declare(strict_types=1);

require_once __DIR__ . '/app.php';

try {
ensureApplicationInstalled();
$user = requireApiAuth();
$pdo = db();

$input = jsonInput();
$action = (string) ($_GET['action'] ?? $_POST['action'] ?? ($input['action'] ?? 'summary'));

switch ($action) {
    case 'summary':
        $today = date('Y-m-d');
        $summary = [
            'notes' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM notes WHERE user_id = :user_id', ['user_id' => $user['id']]),
            'tasks_today' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM daily_tasks WHERE user_id = :user_id AND task_date = :task_date', ['user_id' => $user['id'], 'task_date' => $today]),
            'goals_active' => (int) fetchScalar($pdo, "SELECT COUNT(*) FROM goals WHERE user_id = :user_id AND status = 'active'", ['user_id' => $user['id']]),
            'events_upcoming' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM calendar_events WHERE user_id = :user_id AND start_at >= :start_at', ['user_id' => $user['id'], 'start_at' => date('Y-m-d H:i:s')]),
        ];

        $tasks = fetchAll($pdo, 'SELECT * FROM daily_tasks WHERE user_id = :user_id AND task_date = :task_date ORDER BY status = "done" ASC, due_time IS NULL ASC, due_time ASC, id DESC LIMIT 6', ['user_id' => $user['id'], 'task_date' => $today]);
        $goals = fetchAll($pdo, 'SELECT * FROM goals WHERE user_id = :user_id ORDER BY status = "active" DESC, target_date IS NULL ASC, target_date ASC LIMIT 4', ['user_id' => $user['id']]);
        $events = fetchAll($pdo, 'SELECT * FROM calendar_events WHERE user_id = :user_id AND end_at >= :now ORDER BY start_at ASC LIMIT 6', ['user_id' => $user['id'], 'now' => date('Y-m-d H:i:s')]);
        $notes = fetchAll($pdo, 'SELECT * FROM notes WHERE user_id = :user_id ORDER BY is_pinned DESC, updated_at DESC LIMIT 4', ['user_id' => $user['id']]);

        jsonResponse([
            'ok' => true,
            'summary' => $summary,
            'tasks' => $tasks,
            'goals' => $goals,
            'events' => $events,
            'notes' => $notes,
        ]);
        break;

    case 'notes_list':
        jsonResponse([
            'ok' => true,
            'notes' => fetchAll($pdo, 'SELECT * FROM notes WHERE user_id = :user_id ORDER BY is_pinned DESC, updated_at DESC', ['user_id' => $user['id']]),
        ]);
        break;

    case 'note_save':
        $noteId = (int) ($input['id'] ?? 0);
        $title = trim((string) ($input['title'] ?? ''));
        $content = trim((string) ($input['content'] ?? ''));
        $color = trim((string) ($input['color'] ?? 'blue'));
        $isPinned = !empty($input['is_pinned']) ? 1 : 0;

        if ($title === '' || $content === '') {
            jsonResponse(['ok' => false, 'message' => 'Titulo y contenido son obligatorios.'], 422);
        }

        if ($noteId > 0) {
            executeStatement($pdo, 'UPDATE notes SET title = :title, content = :content, color = :color, is_pinned = :is_pinned WHERE id = :id AND user_id = :user_id', [
                'title' => $title,
                'content' => $content,
                'color' => $color,
                'is_pinned' => $isPinned,
                'id' => $noteId,
                'user_id' => $user['id'],
            ]);
        } else {
            executeStatement($pdo, 'INSERT INTO notes (user_id, title, content, color, is_pinned) VALUES (:user_id, :title, :content, :color, :is_pinned)', [
                'user_id' => $user['id'],
                'title' => $title,
                'content' => $content,
                'color' => $color,
                'is_pinned' => $isPinned,
            ]);
        }

        jsonResponse(['ok' => true, 'message' => 'Nota guardada.']);
        break;

    case 'note_delete':
        executeStatement($pdo, 'DELETE FROM notes WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Nota eliminada.']);
        break;

    case 'tasks_list':
        jsonResponse([
            'ok' => true,
            'tasks' => fetchAll($pdo, 'SELECT * FROM daily_tasks WHERE user_id = :user_id ORDER BY task_date ASC, status = "done" ASC, due_time IS NULL ASC, due_time ASC, id DESC', ['user_id' => $user['id']]),
        ]);
        break;

    case 'task_save':
        $taskId = (int) ($input['id'] ?? 0);
        $title = trim((string) ($input['title'] ?? ''));
        $taskDate = (string) ($input['task_date'] ?? date('Y-m-d'));
        $description = trim((string) ($input['description'] ?? ''));
        $status = trim((string) ($input['status'] ?? 'pending'));
        $priority = trim((string) ($input['priority'] ?? 'medium'));
        $dueTime = trim((string) ($input['due_time'] ?? ''));

        if ($title === '') {
            jsonResponse(['ok' => false, 'message' => 'El titulo de la tarea es obligatorio.'], 422);
        }

        $params = [
            'user_id' => $user['id'],
            'task_date' => $taskDate,
            'title' => $title,
            'description' => $description !== '' ? $description : null,
            'status' => $status,
            'priority' => $priority,
            'due_time' => $dueTime !== '' ? $dueTime : null,
            'completed_at' => $status === 'done' ? date('Y-m-d H:i:s') : null,
        ];

        if ($taskId > 0) {
            $params['id'] = $taskId;
            executeStatement($pdo, 'UPDATE daily_tasks SET task_date = :task_date, title = :title, description = :description, status = :status, priority = :priority, due_time = :due_time, completed_at = :completed_at WHERE id = :id AND user_id = :user_id', $params);
        } else {
            executeStatement($pdo, 'INSERT INTO daily_tasks (user_id, task_date, title, description, status, priority, due_time, completed_at) VALUES (:user_id, :task_date, :title, :description, :status, :priority, :due_time, :completed_at)', $params);
        }

        jsonResponse(['ok' => true, 'message' => 'Tarea guardada.']);
        break;

    case 'task_toggle':
        $taskId = (int) ($input['id'] ?? 0);
        $task = fetchOne($pdo, 'SELECT * FROM daily_tasks WHERE id = :id AND user_id = :user_id LIMIT 1', ['id' => $taskId, 'user_id' => $user['id']]);
        if (!$task) {
            jsonResponse(['ok' => false, 'message' => 'Tarea no encontrada.'], 404);
        }
        $newStatus = $task['status'] === 'done' ? 'pending' : 'done';
        executeStatement($pdo, 'UPDATE daily_tasks SET status = :status, completed_at = :completed_at WHERE id = :id AND user_id = :user_id', [
            'status' => $newStatus,
            'completed_at' => $newStatus === 'done' ? date('Y-m-d H:i:s') : null,
            'id' => $taskId,
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Tarea actualizada.']);
        break;

    case 'task_delete':
        executeStatement($pdo, 'DELETE FROM daily_tasks WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Tarea eliminada.']);
        break;

    case 'goals_list':
        jsonResponse([
            'ok' => true,
            'goals' => fetchAll($pdo, 'SELECT * FROM goals WHERE user_id = :user_id ORDER BY status = "active" DESC, target_date IS NULL ASC, target_date ASC, id DESC', ['user_id' => $user['id']]),
        ]);
        break;

    case 'goal_save':
        $goalId = (int) ($input['id'] ?? 0);
        $title = trim((string) ($input['title'] ?? ''));
        if ($title === '') {
            jsonResponse(['ok' => false, 'message' => 'El titulo de la meta es obligatorio.'], 422);
        }

        $params = [
            'user_id' => $user['id'],
            'title' => $title,
            'description' => trim((string) ($input['description'] ?? '')) ?: null,
            'target_date' => trim((string) ($input['target_date'] ?? '')) ?: null,
            'progress_percent' => max(0, min(100, (int) ($input['progress_percent'] ?? 0))),
            'status' => trim((string) ($input['status'] ?? 'active')),
        ];

        if ($goalId > 0) {
            $params['id'] = $goalId;
            executeStatement($pdo, 'UPDATE goals SET title = :title, description = :description, target_date = :target_date, progress_percent = :progress_percent, status = :status WHERE id = :id AND user_id = :user_id', $params);
        } else {
            executeStatement($pdo, 'INSERT INTO goals (user_id, title, description, target_date, progress_percent, status) VALUES (:user_id, :title, :description, :target_date, :progress_percent, :status)', $params);
        }

        jsonResponse(['ok' => true, 'message' => 'Meta guardada.']);
        break;

    case 'goal_delete':
        executeStatement($pdo, 'DELETE FROM goals WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Meta eliminada.']);
        break;

    case 'calendar_list':
        jsonResponse([
            'ok' => true,
            'events' => fetchAll($pdo, 'SELECT * FROM calendar_events WHERE user_id = :user_id ORDER BY start_at ASC', ['user_id' => $user['id']]),
            'sources' => fetchAll($pdo, 'SELECT provider, external_account_email, sync_enabled, sync_status, last_synced_at FROM calendar_sources WHERE user_id = :user_id ORDER BY provider ASC', ['user_id' => $user['id']]),
        ]);
        break;

    case 'calendar_save':
        $eventId = (int) ($input['id'] ?? 0);
        $title = trim((string) ($input['title'] ?? ''));
        $startAt = trim((string) ($input['start_at'] ?? ''));
        $endAt = trim((string) ($input['end_at'] ?? ''));

        if ($title === '' || $startAt === '' || $endAt === '') {
            jsonResponse(['ok' => false, 'message' => 'Titulo, inicio y fin son obligatorios.'], 422);
        }

        $params = [
            'user_id' => $user['id'],
            'title' => $title,
            'description' => trim((string) ($input['description'] ?? '')) ?: null,
            'start_at' => $startAt,
            'end_at' => $endAt,
            'location' => trim((string) ($input['location'] ?? '')) ?: null,
            'meeting_url' => trim((string) ($input['meeting_url'] ?? '')) ?: null,
            'source_type' => trim((string) ($input['source_type'] ?? 'manual')) ?: 'manual',
        ];

        if ($eventId > 0) {
            $params['id'] = $eventId;
            executeStatement($pdo, 'UPDATE calendar_events SET title = :title, description = :description, start_at = :start_at, end_at = :end_at, location = :location, meeting_url = :meeting_url, source_type = :source_type WHERE id = :id AND user_id = :user_id', $params);
        } else {
            executeStatement($pdo, 'INSERT INTO calendar_events (user_id, title, description, start_at, end_at, location, meeting_url, source_type) VALUES (:user_id, :title, :description, :start_at, :end_at, :location, :meeting_url, :source_type)', $params);
        }

        jsonResponse(['ok' => true, 'message' => 'Evento guardado.']);
        break;

    case 'calendar_delete':
        executeStatement($pdo, 'DELETE FROM calendar_events WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Evento eliminado.']);
        break;

    case 'profile_get':
        jsonResponse([
            'ok' => true,
            'user' => currentUser(),
            'public_profile' => getPublicProfile(env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo')),
        ]);
        break;

    case 'profile_update':
        $fullName = trim((string) ($input['full_name'] ?? ''));
        $email = strtolower(trim((string) ($input['email'] ?? '')));
        $password = trim((string) ($input['password'] ?? ''));

        if ($fullName === '' || $email === '') {
            jsonResponse(['ok' => false, 'message' => 'Nombre y correo son obligatorios.'], 422);
        }

        $sql = 'UPDATE users SET full_name = :full_name, email = :email';
        $params = [
            'full_name' => $fullName,
            'email' => $email,
            'id' => $user['id'],
        ];
        if ($password !== '') {
            $sql .= ', password_hash = :password_hash';
            $params['password_hash'] = password_hash($password, PASSWORD_DEFAULT);
        }
        $sql .= ' WHERE id = :id';
        executeStatement($pdo, $sql, $params);

        jsonResponse(['ok' => true, 'message' => 'Perfil actualizado.']);
        break;

    case 'profile_photo_upload':
        if (!isset($_FILES['photo'])) {
            jsonResponse(['ok' => false, 'message' => 'No se recibio ninguna foto.'], 422);
        }

        try {
            $path = saveProfilePhoto($user, $_FILES['photo'], !empty($_POST['make_public_profile']));
        } catch (Throwable $throwable) {
            jsonResponse(['ok' => false, 'message' => $throwable->getMessage()], 422);
        }

        jsonResponse([
            'ok' => true,
            'message' => 'Foto actualizada.',
            'photo_url' => buildAssetUrl($path),
        ]);
        break;

    case 'public_profile_update':
        executeStatement($pdo, 'UPDATE public_profiles SET display_name = :display_name, role_title = :role_title, company_name = :company_name, bio = :bio, updated_by_user_id = :updated_by_user_id WHERE slug = :slug', [
            'display_name' => trim((string) ($input['display_name'] ?? env('PUBLIC_PROFILE_NAME', 'Cristian Jesus Bonelo Rios'))),
            'role_title' => trim((string) ($input['role_title'] ?? env('PUBLIC_PROFILE_ROLE', 'Software Quality Analyst'))),
            'company_name' => trim((string) ($input['company_name'] ?? env('PUBLIC_PROFILE_COMPANY', 'Olimpia IT'))) ?: null,
            'bio' => trim((string) ($input['bio'] ?? '')) ?: null,
            'updated_by_user_id' => $user['id'],
            'slug' => env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo'),
        ]);

        jsonResponse(['ok' => true, 'message' => 'Perfil publico actualizado.']);
        break;

    default:
        jsonResponse(['ok' => false, 'message' => 'Accion no soportada.'], 400);
}
} catch (Throwable $throwable) {
    jsonResponse([
        'ok' => false,
        'message' => applicationErrorMessage($throwable),
    ], 500);
}

function fetchScalar(PDO $pdo, string $sql, array $params = [])
{
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    return $stmt->fetchColumn();
}

function fetchOne(PDO $pdo, string $sql, array $params = []): ?array
{
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $row = $stmt->fetch();
    return is_array($row) ? $row : null;
}

function fetchAll(PDO $pdo, string $sql, array $params = []): array
{
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $rows = $stmt->fetchAll();
    return is_array($rows) ? $rows : [];
}

function executeStatement(PDO $pdo, string $sql, array $params = []): void
{
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
}