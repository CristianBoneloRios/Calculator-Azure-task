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
        $teamsMinutesToday = (int) fetchScalar($pdo, 'SELECT COALESCE(SUM(TIMESTAMPDIFF(MINUTE, start_at, end_at)), 0) FROM calendar_events WHERE user_id = :user_id AND source_type = :source_type AND start_at BETWEEN :day_start AND :day_end', [
            'user_id' => $user['id'],
            'source_type' => 'power_automate_teams',
            'day_start' => $today . ' 00:00:00',
            'day_end' => $today . ' 23:59:59',
        ]);
        $summary = [
            'notes' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM notes WHERE user_id = :user_id', ['user_id' => $user['id']]),
            'tasks_today' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM daily_tasks WHERE user_id = :user_id AND task_date = :task_date', ['user_id' => $user['id'], 'task_date' => $today]),
            'goals_active' => (int) fetchScalar($pdo, "SELECT COUNT(*) FROM goals WHERE user_id = :user_id AND status = 'active'", ['user_id' => $user['id']]),
            'events_upcoming' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM calendar_events WHERE user_id = :user_id AND start_at >= :start_at', ['user_id' => $user['id'], 'start_at' => date('Y-m-d H:i:s')]),
            'teams_sessions_today' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM calendar_events WHERE user_id = :user_id AND source_type = :source_type AND start_at BETWEEN :day_start AND :day_end', [
                'user_id' => $user['id'],
                'source_type' => 'power_automate_teams',
                'day_start' => $today . ' 00:00:00',
                'day_end' => $today . ' 23:59:59',
            ]),
            'teams_minutes_today' => $teamsMinutesToday,
            'teams_hours_today_label' => formatMinutesAsHoursLabel($teamsMinutesToday),
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
        ensureNotesCollaborationSchema($pdo);

        $notesSql =
            'SELECT n.*,\
                    CASE WHEN n.user_id = :user_id THEN 1 ELSE 0 END AS can_edit,\
                    CASE WHEN n.user_id = :user_id THEN 1 ELSE 0 END AS is_owner,\
                    CASE WHEN n.user_id = :user_id THEN "mine" ELSE "shared" END AS note_scope,\
                    (SELECT COUNT(*) FROM note_comments nc WHERE nc.note_id = n.id) AS comments_count\
                         FROM notes n\
                         WHERE n.user_id = :user_id\
                                OR EXISTS (\
                                        SELECT 1\
                                        FROM note_shares ns\
                                        WHERE ns.note_id = n.id\
                                            AND ns.is_active = 1\
                                            AND LOWER(ns.invited_email) = :user_email\
                                )\
             ORDER BY n.is_pinned DESC, n.updated_at DESC';

        jsonResponse([
            'ok' => true,
            'notes' => fetchAll($pdo, $notesSql, [
                'user_id' => (int) $user['id'],
                'user_email' => strtolower((string) ($user['email'] ?? '')),
            ]),
        ]);
        break;

    case 'note_save':
        ensureNotesCollaborationSchema($pdo);

        $noteId = (int) ($input['id'] ?? 0);
        $title = trim((string) ($input['title'] ?? ''));
        $content = trim((string) ($input['content'] ?? ''));
        $color = trim((string) ($input['color'] ?? 'blue'));
        $isPinned = !empty($input['is_pinned']) ? 1 : 0;
        $allowedColors = ['blue', 'yellow', 'purple', 'cyan'];

        if ($title === '' || $content === '') {
            jsonResponse(['ok' => false, 'message' => 'Titulo y contenido son obligatorios.'], 422);
        }

        if (!in_array($color, $allowedColors, true)) {
            $color = 'blue';
        }

        if ($noteId > 0) {
            $editable = fetchOne($pdo, 'SELECT id FROM notes WHERE id = :id AND user_id = :user_id LIMIT 1', [
                'id' => $noteId,
                'user_id' => (int) $user['id'],
            ]);
            if ($editable === null) {
                jsonResponse(['ok' => false, 'message' => 'Solo el propietario puede editar esta nota.'], 403);
            }

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
        ensureNotesCollaborationSchema($pdo);

        executeStatement($pdo, 'DELETE FROM notes WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Nota eliminada.']);
        break;

    case 'note_comments_list':
        ensureNotesCollaborationSchema($pdo);

        $noteId = (int) ($input['note_id'] ?? 0);
        if ($noteId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Nota invalida.'], 422);
        }

        if (!userCanAccessNote($pdo, $noteId, (int) $user['id'], strtolower((string) ($user['email'] ?? '')))) {
            jsonResponse(['ok' => false, 'message' => 'No tienes acceso a esta nota.'], 403);
        }

        $comments = fetchAll(
            $pdo,
            'SELECT nc.id, nc.note_id, nc.user_id, nc.parent_comment_id, nc.content, nc.created_at, nc.updated_at,
                    u.full_name AS author_name, u.email AS author_email
             FROM note_comments nc
             INNER JOIN users u ON u.id = nc.user_id
             WHERE nc.note_id = :note_id
             ORDER BY nc.created_at ASC, nc.id ASC',
            ['note_id' => $noteId]
        );

        jsonResponse([
            'ok' => true,
            'comments' => $comments,
        ]);
        break;

    case 'note_comment_add':
        ensureNotesCollaborationSchema($pdo);

        $noteId = (int) ($input['note_id'] ?? 0);
        $parentCommentId = (int) ($input['parent_comment_id'] ?? 0);
        $content = trim((string) ($input['content'] ?? ''));

        if ($noteId <= 0 || $content === '') {
            jsonResponse(['ok' => false, 'message' => 'Nota y comentario son obligatorios.'], 422);
        }

        if (mb_strlen($content) > 3000) {
            jsonResponse(['ok' => false, 'message' => 'El comentario supera el limite de 3000 caracteres.'], 422);
        }

        if (!userCanAccessNote($pdo, $noteId, (int) $user['id'], strtolower((string) ($user['email'] ?? '')))) {
            jsonResponse(['ok' => false, 'message' => 'No tienes acceso a esta nota.'], 403);
        }

        if ($parentCommentId > 0) {
            $parentExists = fetchOne(
                $pdo,
                'SELECT id FROM note_comments WHERE id = :id AND note_id = :note_id LIMIT 1',
                ['id' => $parentCommentId, 'note_id' => $noteId]
            );
            if ($parentExists === null) {
                jsonResponse(['ok' => false, 'message' => 'El comentario padre no existe para esta nota.'], 422);
            }
        }

        executeStatement(
            $pdo,
            'INSERT INTO note_comments (note_id, user_id, parent_comment_id, content)
             VALUES (:note_id, :user_id, :parent_comment_id, :content)',
            [
                'note_id' => $noteId,
                'user_id' => (int) $user['id'],
                'parent_comment_id' => $parentCommentId > 0 ? $parentCommentId : null,
                'content' => $content,
            ]
        );

        jsonResponse(['ok' => true, 'message' => 'Comentario agregado.']);
        break;

    case 'note_comment_delete':
        ensureNotesCollaborationSchema($pdo);

        $commentId = (int) ($input['comment_id'] ?? 0);
        if ($commentId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Comentario invalido.'], 422);
        }

        $comment = fetchOne(
            $pdo,
            'SELECT nc.id, nc.note_id, nc.user_id, n.user_id AS note_owner_user_id
             FROM note_comments nc
             INNER JOIN notes n ON n.id = nc.note_id
             WHERE nc.id = :id
             LIMIT 1',
            ['id' => $commentId]
        );

        if ($comment === null) {
            jsonResponse(['ok' => false, 'message' => 'Comentario no encontrado.'], 404);
        }

        $isCommentAuthor = (int) $comment['user_id'] === (int) $user['id'];
        $isNoteOwner = (int) $comment['note_owner_user_id'] === (int) $user['id'];
        if (!$isCommentAuthor && !$isNoteOwner) {
            jsonResponse(['ok' => false, 'message' => 'No tienes permiso para eliminar este comentario.'], 403);
        }

        executeStatement($pdo, 'DELETE FROM note_comments WHERE id = :id', ['id' => $commentId]);
        jsonResponse(['ok' => true, 'message' => 'Comentario eliminado.']);
        break;

    case 'note_shares_list':
        ensureNotesCollaborationSchema($pdo);

        $noteId = (int) ($input['note_id'] ?? 0);
        if ($noteId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Nota invalida.'], 422);
        }

        $ownedNote = fetchOne($pdo, 'SELECT id FROM notes WHERE id = :id AND user_id = :user_id LIMIT 1', [
            'id' => $noteId,
            'user_id' => (int) $user['id'],
        ]);
        if ($ownedNote === null) {
            jsonResponse(['ok' => false, 'message' => 'Solo el propietario puede gestionar invitaciones.'], 403);
        }

        jsonResponse([
            'ok' => true,
            'shares' => fetchAll(
                $pdo,
                'SELECT id, note_id, invited_email, invited_user_id, is_active, created_at, updated_at
                 FROM note_shares
                 WHERE note_id = :note_id
                 ORDER BY created_at DESC',
                ['note_id' => $noteId]
            ),
        ]);
        break;

    case 'note_share_invite':
        ensureNotesCollaborationSchema($pdo);

        $noteId = (int) ($input['note_id'] ?? 0);
        $invitedEmail = strtolower(trim((string) ($input['email'] ?? '')));

        if ($noteId <= 0 || $invitedEmail === '' || !filter_var($invitedEmail, FILTER_VALIDATE_EMAIL)) {
            jsonResponse(['ok' => false, 'message' => 'Debes enviar una nota y un correo valido.'], 422);
        }

        $ownedNote = fetchOne($pdo, 'SELECT id FROM notes WHERE id = :id AND user_id = :user_id LIMIT 1', [
            'id' => $noteId,
            'user_id' => (int) $user['id'],
        ]);
        if ($ownedNote === null) {
            jsonResponse(['ok' => false, 'message' => 'Solo el propietario puede invitar por correo.'], 403);
        }

        $invitedUser = fetchOne($pdo, 'SELECT id FROM users WHERE email = :email LIMIT 1', ['email' => $invitedEmail]);

        executeStatement(
            $pdo,
            'INSERT INTO note_shares (note_id, owner_user_id, invited_email, invited_user_id, is_active)
             VALUES (:note_id, :owner_user_id, :invited_email, :invited_user_id, 1)
             ON DUPLICATE KEY UPDATE
                 owner_user_id = VALUES(owner_user_id),
                 invited_user_id = VALUES(invited_user_id),
                 is_active = 1,
                 updated_at = CURRENT_TIMESTAMP',
            [
                'note_id' => $noteId,
                'owner_user_id' => (int) $user['id'],
                'invited_email' => $invitedEmail,
                'invited_user_id' => $invitedUser ? (int) $invitedUser['id'] : null,
            ]
        );

        jsonResponse(['ok' => true, 'message' => 'Invitacion por correo registrada.']);
        break;

    case 'note_share_revoke':
        ensureNotesCollaborationSchema($pdo);

        $shareId = (int) ($input['share_id'] ?? 0);
        if ($shareId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Invitacion invalida.'], 422);
        }

        $share = fetchOne(
            $pdo,
            'SELECT id FROM note_shares WHERE id = :id AND owner_user_id = :owner_user_id LIMIT 1',
            ['id' => $shareId, 'owner_user_id' => (int) $user['id']]
        );
        if ($share === null) {
            jsonResponse(['ok' => false, 'message' => 'Solo el propietario puede revocar invitaciones.'], 403);
        }

        executeStatement($pdo, 'UPDATE note_shares SET is_active = 0, updated_at = CURRENT_TIMESTAMP WHERE id = :id', ['id' => $shareId]);
        jsonResponse(['ok' => true, 'message' => 'Invitacion revocada.']);
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

    case 'finance_list':
        $movements = fetchAll(
            $pdo,
            'SELECT id, movement_type, amount, category, description, movement_date, source_channel, created_at
             FROM finance_movements
             WHERE user_id = :user_id
             ORDER BY movement_date DESC, id DESC',
            ['user_id' => (int) $user['id']]
        );

        $totals = [
            'income' => 0.0,
            'expense' => 0.0,
            'investment' => 0.0,
            'tax' => 0.0,
        ];

        foreach ($movements as $movement) {
            $type = (string) ($movement['movement_type'] ?? '');
            if (!array_key_exists($type, $totals)) {
                continue;
            }
            $totals[$type] += (float) ($movement['amount'] ?? 0);
        }

        $net = $totals['income'] - $totals['expense'] - $totals['investment'] - $totals['tax'];
        $savingsRate = $totals['income'] > 0 ? max(0, ($net / $totals['income']) * 100) : 0;

        jsonResponse([
            'ok' => true,
            'movements' => $movements,
            'summary' => [
                'income' => round($totals['income'], 2),
                'expense' => round($totals['expense'], 2),
                'investment' => round($totals['investment'], 2),
                'tax' => round($totals['tax'], 2),
                'net' => round($net, 2),
                'savings_rate' => round($savingsRate, 2),
            ],
        ]);
        break;

    case 'finance_save':
        $type = trim((string) ($input['type'] ?? 'income'));
        $amount = (float) ($input['amount'] ?? 0);
        $category = trim((string) ($input['category'] ?? ''));
        $movementDate = trim((string) ($input['date'] ?? ''));
        $description = trim((string) ($input['description'] ?? ''));

        $allowedTypes = ['income', 'expense', 'investment', 'tax'];
        if (!in_array($type, $allowedTypes, true)) {
            jsonResponse(['ok' => false, 'message' => 'Tipo de movimiento invalido.'], 422);
        }

        if (!is_finite($amount) || $amount <= 0) {
            jsonResponse(['ok' => false, 'message' => 'El monto debe ser mayor a cero.'], 422);
        }

        if ($movementDate === '') {
            $movementDate = date('Y-m-d');
        }
        if (!preg_match('/^\d{4}-\d{2}-\d{2}$/', $movementDate)) {
            jsonResponse(['ok' => false, 'message' => 'La fecha del movimiento no es valida.'], 422);
        }

        executeStatement(
            $pdo,
            'INSERT INTO finance_movements (user_id, movement_type, amount, category, description, movement_date, source_channel)
             VALUES (:user_id, :movement_type, :amount, :category, :description, :movement_date, :source_channel)',
            [
                'user_id' => (int) $user['id'],
                'movement_type' => $type,
                'amount' => round($amount, 2),
                'category' => $category !== '' ? mb_substr($category, 0, 120) : null,
                'description' => $description !== '' ? mb_substr($description, 0, 3000) : null,
                'movement_date' => $movementDate,
                'source_channel' => 'workspace_finanzas',
            ]
        );

        jsonResponse([
            'ok' => true,
            'message' => 'Movimiento guardado en base de datos.',
            'id' => (int) $pdo->lastInsertId(),
        ]);
        break;

    case 'finance_simulation_save':
        $simulatorType = trim((string) ($input['simulator_type'] ?? ''));
        $inputPayload = $input['input_payload'] ?? null;
        $resultPayload = $input['result_payload'] ?? null;
        $resultSummary = trim((string) ($input['result_summary'] ?? ''));
        $simulatedAtRaw = trim((string) ($input['simulated_at'] ?? ''));

        $allowedSimulatorTypes = ['compound_interest', 'home_capacity', 'leasing_vs_credit', 'tax_burden'];
        if (!in_array($simulatorType, $allowedSimulatorTypes, true)) {
            jsonResponse(['ok' => false, 'message' => 'Tipo de simulador invalido.'], 422);
        }

        if (!is_array($inputPayload) || !is_array($resultPayload)) {
            jsonResponse(['ok' => false, 'message' => 'Los datos de entrada y resultado deben ser JSON valido.'], 422);
        }

        if ($resultSummary === '') {
            jsonResponse(['ok' => false, 'message' => 'El resumen del resultado es obligatorio.'], 422);
        }

        $simulatedAtTs = $simulatedAtRaw !== '' ? strtotime($simulatedAtRaw) : time();
        if ($simulatedAtTs === false) {
            $simulatedAtTs = time();
        }
        $simulatedAt = date('Y-m-d H:i:s', $simulatedAtTs);

        executeStatement(
            $pdo,
            'INSERT INTO finance_simulation_history (user_id, simulator_type, input_payload, result_payload, result_summary, simulated_at)
             VALUES (:user_id, :simulator_type, :input_payload, :result_payload, :result_summary, :simulated_at)',
            [
                'user_id' => (int) $user['id'],
                'simulator_type' => $simulatorType,
                'input_payload' => json_encode($inputPayload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES),
                'result_payload' => json_encode($resultPayload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES),
                'result_summary' => mb_substr($resultSummary, 0, 2000),
                'simulated_at' => $simulatedAt,
            ]
        );

        $insertId = (int) $pdo->lastInsertId();
        jsonResponse([
            'ok' => true,
            'message' => 'Simulacion guardada en historial.',
            'simulation' => [
                'id' => $insertId,
                'simulator_type' => $simulatorType,
                'input_payload' => $inputPayload,
                'result_payload' => $resultPayload,
                'result_summary' => mb_substr($resultSummary, 0, 2000),
                'simulated_at' => $simulatedAt,
            ],
        ]);
        break;

    case 'finance_simulation_list':
        $simulatorTypeFilter = trim((string) ($input['simulator_type'] ?? ''));
        $fromDateFilter = trim((string) ($input['from_date'] ?? ''));
        $toDateFilter = trim((string) ($input['to_date'] ?? ''));
        $page = (int) ($input['page'] ?? 1);
        if ($page < 1) {
            $page = 1;
        }

        $pageSize = (int) ($input['page_size'] ?? ($input['limit'] ?? 20));
        if ($pageSize < 1) {
            $pageSize = 20;
        }
        if ($pageSize > 100) {
            $pageSize = 100;
        }

        $allowedSimulatorTypes = ['compound_interest', 'home_capacity', 'leasing_vs_credit', 'tax_burden'];
        if ($simulatorTypeFilter !== '' && !in_array($simulatorTypeFilter, $allowedSimulatorTypes, true)) {
            jsonResponse(['ok' => false, 'message' => 'Filtro de simulador invalido.'], 422);
        }

        if ($fromDateFilter !== '' && !preg_match('/^\d{4}-\d{2}-\d{2}$/', $fromDateFilter)) {
            jsonResponse(['ok' => false, 'message' => 'Filtro fecha desde invalido.'], 422);
        }

        if ($toDateFilter !== '' && !preg_match('/^\d{4}-\d{2}-\d{2}$/', $toDateFilter)) {
            jsonResponse(['ok' => false, 'message' => 'Filtro fecha hasta invalido.'], 422);
        }

        $conditions = ['user_id = :user_id'];
        $params = ['user_id' => (int) $user['id']];

        if ($simulatorTypeFilter !== '') {
            $conditions[] = 'simulator_type = :simulator_type';
            $params['simulator_type'] = $simulatorTypeFilter;
        }

        if ($fromDateFilter !== '') {
            $conditions[] = 'simulated_at >= :from_date_start';
            $params['from_date_start'] = $fromDateFilter . ' 00:00:00';
        }

        if ($toDateFilter !== '') {
            $conditions[] = 'simulated_at <= :to_date_end';
            $params['to_date_end'] = $toDateFilter . ' 23:59:59';
        }

        $whereClause = implode(' AND ', $conditions);
        $total = (int) fetchScalar(
            $pdo,
            'SELECT COUNT(*) FROM finance_simulation_history WHERE ' . $whereClause,
            $params
        );
        $totalPages = $total > 0 ? (int) ceil($total / $pageSize) : 1;
        if ($page > $totalPages) {
            $page = $totalPages;
        }
        $offset = ($page - 1) * $pageSize;

        $rows = fetchAll(
            $pdo,
            'SELECT id, simulator_type, input_payload, result_payload, result_summary, simulated_at, created_at
             FROM finance_simulation_history
             WHERE ' . $whereClause . '
             ORDER BY simulated_at DESC, id DESC
             LIMIT ' . $pageSize . ' OFFSET ' . $offset,
            $params
        );

        $simulations = array_map(static function (array $row): array {
            $inputPayload = json_decode((string) ($row['input_payload'] ?? '{}'), true);
            $resultPayload = json_decode((string) ($row['result_payload'] ?? '{}'), true);
            return [
                'id' => (int) ($row['id'] ?? 0),
                'simulator_type' => (string) ($row['simulator_type'] ?? ''),
                'input_payload' => is_array($inputPayload) ? $inputPayload : [],
                'result_payload' => is_array($resultPayload) ? $resultPayload : [],
                'result_summary' => (string) ($row['result_summary'] ?? ''),
                'simulated_at' => (string) ($row['simulated_at'] ?? ''),
                'created_at' => (string) ($row['created_at'] ?? ''),
            ];
        }, $rows);

        jsonResponse([
            'ok' => true,
            'simulations' => $simulations,
            'pagination' => [
                'page' => $page,
                'page_size' => $pageSize,
                'total' => $total,
                'total_pages' => $totalPages,
                'has_prev' => $page > 1,
                'has_next' => $page < $totalPages,
            ],
        ]);
        break;

    case 'finance_simulation_delete':
        $simulationId = (int) ($input['simulation_id'] ?? 0);
        if ($simulationId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Debes seleccionar una simulacion valida.'], 422);
        }

        executeStatement(
            $pdo,
            'DELETE FROM finance_simulation_history WHERE id = :id AND user_id = :user_id',
            [
                'id' => $simulationId,
                'user_id' => (int) $user['id'],
            ]
        );

        jsonResponse([
            'ok' => true,
            'message' => 'Simulacion eliminada del historial.',
        ]);
        break;

    case 'finance_simulation_purge':
        $olderThanDays = (int) ($input['older_than_days'] ?? 90);
        if ($olderThanDays < 7) {
            $olderThanDays = 7;
        }
        if ($olderThanDays > 3650) {
            $olderThanDays = 3650;
        }

        $cutoff = date('Y-m-d H:i:s', time() - ($olderThanDays * 86400));

        $stmt = $pdo->prepare('DELETE FROM finance_simulation_history WHERE user_id = :user_id AND simulated_at < :cutoff');
        $stmt->execute([
            'user_id' => (int) $user['id'],
            'cutoff' => $cutoff,
        ]);

        jsonResponse([
            'ok' => true,
            'message' => 'Se eliminaron simulaciones antiguas.',
            'deleted_count' => (int) $stmt->rowCount(),
            'cutoff' => $cutoff,
        ]);
        break;

    case 'finance_simulation_export_pdf':
        $simulationId = (int) ($input['simulation_id'] ?? 0);
        $simulatorType = trim((string) ($input['simulator_type'] ?? ''));
        $inputPayload = is_array($input['input_payload'] ?? null) ? $input['input_payload'] : [];
        $resultPayload = is_array($input['result_payload'] ?? null) ? $input['result_payload'] : [];
        $resultSummary = trim((string) ($input['result_summary'] ?? ''));
        $simulatedAtRaw = trim((string) ($input['simulated_at'] ?? ''));

        if ($simulationId > 0) {
            $stored = fetchOne(
                $pdo,
                'SELECT simulator_type, input_payload, result_payload, result_summary, simulated_at
                 FROM finance_simulation_history
                 WHERE id = :id AND user_id = :user_id
                 LIMIT 1',
                ['id' => $simulationId, 'user_id' => (int) $user['id']]
            );

            if ($stored === null) {
                jsonResponse(['ok' => false, 'message' => 'No se encontro esa simulacion en tu historial.'], 404);
            }

            $simulatorType = (string) ($stored['simulator_type'] ?? $simulatorType);
            $decodedInput = json_decode((string) ($stored['input_payload'] ?? '{}'), true);
            $decodedResult = json_decode((string) ($stored['result_payload'] ?? '{}'), true);
            $inputPayload = is_array($decodedInput) ? $decodedInput : [];
            $resultPayload = is_array($decodedResult) ? $decodedResult : [];
            $resultSummary = (string) ($stored['result_summary'] ?? $resultSummary);
            $simulatedAtRaw = (string) ($stored['simulated_at'] ?? $simulatedAtRaw);
        }

        if ($simulatorType === '' || $resultSummary === '') {
            jsonResponse(['ok' => false, 'message' => 'Faltan datos para generar el PDF.'], 422);
        }

        $simulatorLabel = financeSimulatorLabel($simulatorType);
        $simulatedTs = $simulatedAtRaw !== '' ? strtotime($simulatedAtRaw) : time();
        if ($simulatedTs === false) {
            $simulatedTs = time();
        }

        $lines = [
            'Reporte de simulacion financiera',
            'Simulador: ' . $simulatorLabel,
            'Fecha: ' . date('Y-m-d H:i:s', $simulatedTs),
            'Usuario: ' . ((string) ($user['full_name'] ?? 'Usuario')),
            'Email: ' . ((string) ($user['email'] ?? '')),
            'Resumen: ' . $resultSummary,
            '--- Entradas ---',
        ];

        foreach (flattenFinancePdfFields($inputPayload, 'input') as $fieldLine) {
            $lines[] = $fieldLine;
        }

        $lines[] = '--- Resultados ---';
        foreach (flattenFinancePdfFields($resultPayload, 'result') as $fieldLine) {
            $lines[] = $fieldLine;
        }

        $pdfBinary = buildSimpleFinancePdf($lines);
        $filename = 'simulacion-' . preg_replace('/[^a-z0-9\-]+/i', '-', strtolower($simulatorType)) . '-' . date('Ymd-His') . '.pdf';

        jsonResponse([
            'ok' => true,
            'file_name' => $filename,
            'file_base64' => base64_encode($pdfBinary),
            'mime_type' => 'application/pdf',
        ]);
        break;

    case 'calendar_list':
        jsonResponse([
            'ok' => true,
            'events' => fetchAll($pdo, 'SELECT * FROM calendar_events WHERE user_id = :user_id ORDER BY start_at ASC', ['user_id' => $user['id']]),
            'sources' => fetchAll($pdo, 'SELECT provider, external_account_email, sync_enabled, sync_status, last_synced_at FROM calendar_sources WHERE user_id = :user_id ORDER BY provider ASC', ['user_id' => $user['id']]),
            'power_automate' => getPowerAutomateSetup((int) $user['id']),
            'teams_today' => [
                'sessions' => (int) fetchScalar($pdo, 'SELECT COUNT(*) FROM calendar_events WHERE user_id = :user_id AND source_type = :source_type AND start_at BETWEEN :day_start AND :day_end', [
                    'user_id' => $user['id'],
                    'source_type' => 'power_automate_teams',
                    'day_start' => date('Y-m-d') . ' 00:00:00',
                    'day_end' => date('Y-m-d') . ' 23:59:59',
                ]),
                'minutes' => (int) fetchScalar($pdo, 'SELECT COALESCE(SUM(TIMESTAMPDIFF(MINUTE, start_at, end_at)), 0) FROM calendar_events WHERE user_id = :user_id AND source_type = :source_type AND start_at BETWEEN :day_start AND :day_end', [
                    'user_id' => $user['id'],
                    'source_type' => 'power_automate_teams',
                    'day_start' => date('Y-m-d') . ' 00:00:00',
                    'day_end' => date('Y-m-d') . ' 23:59:59',
                ]),
            ],
        ]);
        break;

    case 'power_automate_config_get':
        jsonResponse([
            'ok' => true,
            'config' => getPowerAutomateSetup((int) $user['id']),
        ]);
        break;

    case 'power_automate_config_rotate':
        $externalAccountEmail = strtolower(trim((string) ($input['external_account_email'] ?? '')));
        if ($externalAccountEmail !== '' && !filter_var($externalAccountEmail, FILTER_VALIDATE_EMAIL)) {
            jsonResponse(['ok' => false, 'message' => 'El correo origen de Power Automate no es valido.'], 422);
        }

        jsonResponse([
            'ok' => true,
            'message' => 'Clave de Power Automate generada.',
            'config' => createOrRotatePowerAutomateSecret((int) $user['id'], $externalAccountEmail !== '' ? $externalAccountEmail : null),
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

        if (!preg_match('/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/', $startAt) || !preg_match('/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/', $endAt)) {
            jsonResponse(['ok' => false, 'message' => 'Formato de fecha invalido. Usa YYYY-MM-DD HH:mm:ss.'], 422);
        }

        if (strtotime($endAt) <= strtotime($startAt)) {
            jsonResponse(['ok' => false, 'message' => 'La fecha/hora de fin debe ser mayor a la de inicio.'], 422);
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

        $isInsert = $eventId <= 0;

        if ($isInsert) {
            $existingDuplicateId = (int) fetchScalar(
                $pdo,
                'SELECT COALESCE(MAX(id), 0) FROM calendar_events WHERE user_id = :user_id AND title = :title AND start_at = :start_at AND end_at = :end_at',
                [
                    'user_id' => $user['id'],
                    'title' => $title,
                    'start_at' => $startAt,
                    'end_at' => $endAt,
                ]
            );

            if ($existingDuplicateId > 0) {
                jsonResponse([
                    'ok' => true,
                    'id' => $existingDuplicateId,
                    'message' => 'Evento ya existente. Se evito duplicado.',
                    'outbound_sync' => [
                        'attempted' => false,
                        'ok' => false,
                        'status' => '',
                        'message' => 'Sin envio externo por deteccion de duplicado.',
                    ],
                ]);
            }
        }

        if (!$isInsert) {
            $params['id'] = $eventId;
            executeStatement($pdo, 'UPDATE calendar_events SET title = :title, description = :description, start_at = :start_at, end_at = :end_at, location = :location, meeting_url = :meeting_url, source_type = :source_type WHERE id = :id AND user_id = :user_id', $params);
        } else {
            executeStatement($pdo, 'INSERT INTO calendar_events (user_id, title, description, start_at, end_at, location, meeting_url, source_type) VALUES (:user_id, :title, :description, :start_at, :end_at, :location, :meeting_url, :source_type)', $params);
        }

        $savedId = $eventId > 0 ? $eventId : (int) $pdo->lastInsertId();

        $syncResult = [
            'attempted' => false,
            'ok' => false,
            'status' => '',
            'message' => '',
        ];

        $isPowerAutomateSource = in_array($params['source_type'], ['power_automate', 'power_automate_teams'], true);

        // Automatic outbound sync only for newly-created local events (avoid loops).
        if ($isInsert && !$isPowerAutomateSource) {
            $paSetup = getPowerAutomateSetup((int) $user['id']);
            $outboundUrl = (string) ($paSetup['outbound_webhook_url'] ?? '');

            if ($outboundUrl !== '') {
                $syncPayload = json_encode([
                    'action' => 'create_event',
                    'source_email' => (string) ($paSetup['external_account_email'] ?? ($user['email'] ?? '')),
                    'event' => [
                        'title' => $title,
                        'start' => $startAt,
                        'end' => $endAt,
                        'meeting_url' => (string) ($params['meeting_url'] ?? ''),
                        'type' => (string) ($params['source_type'] ?? 'meeting'),
                    ],
                ], JSON_THROW_ON_ERROR);

                $outResult = sendOutboundJsonRequest($outboundUrl, $syncPayload, 12);
                $syncResult = [
                    'attempted' => true,
                    'ok' => (bool) $outResult['ok'],
                    'status' => (string) ($outResult['status_line'] ?? ''),
                    'message' => (bool) $outResult['ok']
                        ? 'Power Automate confirmo create_event.'
                        : buildPowerAutomateErrorMessage('Power Automate no confirmo create_event.', $outResult),
                ];

                if (!$syncResult['ok']) {
                    error_log('calendar_save outbound create_event failed [user_id=' . (int) $user['id'] . ', event_id=' . $savedId . ']: ' . $syncResult['message']);
                }
            }
        }

        jsonResponse([
            'ok' => true,
            'id' => $savedId,
            'message' => 'Evento guardado.',
            'outbound_sync' => $syncResult,
        ]);
        break;

    case 'calendar_push_teams':
        $paSetup = getPowerAutomateSetup((int) $user['id']);
        $outboundUrl = (string) ($paSetup['outbound_webhook_url'] ?? '');
        if ($outboundUrl === '') {
            jsonResponse(['ok' => false, 'message' => 'No hay URL de flujo de salida configurada en Power Automate.'], 422);
        }

        $pushTitle    = trim((string) ($input['title']       ?? ''));
        $pushStart    = trim((string) ($input['start_at']    ?? ''));
        $pushEnd      = trim((string) ($input['end_at']      ?? ''));
        $pushUrl      = trim((string) ($input['meeting_url'] ?? ''));
        $pushType     = trim((string) ($input['event_type']  ?? 'manual'));

        if ($pushTitle === '' || $pushStart === '' || $pushEnd === '') {
            jsonResponse(['ok' => false, 'message' => 'Datos del evento incompletos para enviar a Teams.'], 422);
        }

        $pushPayload = json_encode([
            'action'       => 'create_event',
            'source_email' => $paSetup['external_account_email'] ?? '',
            'event'        => [
                'title'       => $pushTitle,
                'start'       => $pushStart,
                'end'         => $pushEnd,
                'meeting_url' => $pushUrl,
                'type'        => $pushType,
            ],
        ], JSON_THROW_ON_ERROR);

        $pushResult = sendOutboundJsonRequest($outboundUrl, $pushPayload, 12);
        $pushOk = $pushResult['ok'];

        jsonResponse([
            'ok' => $pushOk,
            'status' => $pushResult['status_line'],
            'message' => $pushOk
                ? 'Evento enviado a Power Automate.'
                : buildPowerAutomateErrorMessage('Power Automate respondio con error.', $pushResult),
        ], $pushOk ? 200 : 502);
        break;

    case 'calendar_sync_request':
        $paSetup = getPowerAutomateSetup((int) $user['id']);
        $outboundUrl = (string) ($paSetup['outbound_webhook_url'] ?? '');
        if ($outboundUrl === '') {
            jsonResponse(['ok' => false, 'message' => 'Configura primero la URL de salida de Power Automate.'], 422);
        }

        $syncPayload = json_encode([
            'action' => 'sync_request',
            'requested_at' => date(DATE_ATOM),
            'request_day' => date('Y-m-d'),
            'user' => [
                'id' => (int) $user['id'],
                'email' => (string) ($user['email'] ?? ''),
                'source_email' => (string) ($paSetup['external_account_email'] ?? ''),
            ],
        ], JSON_THROW_ON_ERROR);

        $syncResult = sendOutboundJsonRequest($outboundUrl, $syncPayload, 15);
        $syncOk = $syncResult['ok'];

        if ($syncOk) {
            executeStatement($pdo, 'UPDATE calendar_sources SET sync_status = :sync_status, updated_at = CURRENT_TIMESTAMP WHERE user_id = :user_id AND provider = :provider', [
                'sync_status' => 'sync_requested',
                'user_id' => (int) $user['id'],
                'provider' => 'power_automate',
            ]);
        }

        jsonResponse([
            'ok' => $syncOk,
            'status' => $syncResult['status_line'],
            'message' => $syncOk
                ? 'Solicitud de sincronizacion enviada a Power Automate.'
                : buildPowerAutomateErrorMessage('No se pudo solicitar sincronizacion. Revisa la URL del flujo y el trigger HTTP.', $syncResult),
        ], $syncOk ? 200 : 502);
        break;

    case 'power_automate_set_outbound':
        $outUrl = trim((string) ($input['outbound_webhook_url'] ?? ''));
        if ($outUrl !== '' && !filter_var($outUrl, FILTER_VALIDATE_URL)) {
            jsonResponse(['ok' => false, 'message' => 'URL de salida no valida.'], 422);
        }
        savePowerAutomateOutboundUrl((int) $user['id'], $outUrl !== '' ? $outUrl : null);
        jsonResponse(['ok' => true, 'message' => 'URL de salida guardada.']);
        break;

    case 'calendar_delete':
        executeStatement($pdo, 'DELETE FROM calendar_events WHERE id = :id AND user_id = :user_id', [
            'id' => (int) ($input['id'] ?? 0),
            'user_id' => $user['id'],
        ]);
        jsonResponse(['ok' => true, 'message' => 'Evento eliminado.']);
        break;

    case 'profile_get':
        $developerProfile = getDeveloperIdentityProfile();
        $canManageDeveloperProfile = userCanManageDeveloperIdentity($user, $developerProfile);

        jsonResponse([
            'ok' => true,
            'user' => currentUser(),
            'public_profile' => getPublicProfile(env('PUBLIC_PROFILE_SLUG', 'cristian-bonelo')),
            'developer_profile' => $developerProfile,
            'can_manage_developer_profile' => $canManageDeveloperProfile,
            'admin_users' => $canManageDeveloperProfile
                ? fetchAll($pdo, 'SELECT id, full_name, email, role FROM users WHERE role = :role ORDER BY id ASC', ['role' => 'admin'])
                : [],
        ]);
        break;

    case 'developer_profile_photo_upload':
        $developerProfile = getDeveloperIdentityProfile();
        if (!userCanManageDeveloperIdentity($user, $developerProfile)) {
            jsonResponse(['ok' => false, 'message' => 'Solo el admin propietario puede modificar la foto del desarrollador.'], 403);
        }

        if (!isset($_FILES['photo'])) {
            jsonResponse(['ok' => false, 'message' => 'No se recibio ninguna foto.'], 422);
        }

        try {
            $path = saveDeveloperIdentityPhoto($user, $_FILES['photo']);
        } catch (Throwable $throwable) {
            jsonResponse(['ok' => false, 'message' => $throwable->getMessage()], 422);
        }

        jsonResponse([
            'ok' => true,
            'message' => 'Foto del desarrollador actualizada.',
            'photo_url' => buildAssetUrl($path),
            'developer_profile' => getDeveloperIdentityProfile(),
        ]);
        break;

    case 'developer_profile_transfer_owner':
        $developerProfile = getDeveloperIdentityProfile();
        if (!userCanManageDeveloperIdentity($user, $developerProfile)) {
            jsonResponse(['ok' => false, 'message' => 'Solo el admin propietario puede transferir la propiedad del perfil del desarrollador.'], 403);
        }

        $targetUserId = (int) ($input['target_user_id'] ?? 0);
        if ($targetUserId <= 0) {
            jsonResponse(['ok' => false, 'message' => 'Debes seleccionar un admin destino valido.'], 422);
        }

        $targetUser = fetchOne($pdo, 'SELECT id, email, role FROM users WHERE id = :id LIMIT 1', ['id' => $targetUserId]);
        if ($targetUser === null || (string) ($targetUser['role'] ?? '') !== 'admin') {
            jsonResponse(['ok' => false, 'message' => 'El usuario seleccionado no es admin.'], 422);
        }

        executeStatement(
            $pdo,
            'UPDATE developer_identity_profile
             SET owner_user_id = :owner_user_id,
                 owner_email = :owner_email,
                 updated_by_user_id = :updated_by_user_id,
                 updated_at = CURRENT_TIMESTAMP
             WHERE id = 1',
            [
                'owner_user_id' => (int) $targetUser['id'],
                'owner_email' => (string) $targetUser['email'],
                'updated_by_user_id' => (int) $user['id'],
            ]
        );

        jsonResponse([
            'ok' => true,
            'message' => 'Propiedad del perfil del desarrollador transferida al nuevo admin.',
            'developer_profile' => getDeveloperIdentityProfile(),
        ]);
        break;

    case 'developer_profile_promote_admin':
        $developerProfile = getDeveloperIdentityProfile();
        if (!userCanManageDeveloperIdentity($user, $developerProfile)) {
            jsonResponse(['ok' => false, 'message' => 'Solo el admin propietario puede crear nuevos admins desde ajustes.'], 403);
        }

        $targetEmail = strtolower(trim((string) ($input['email'] ?? '')));
        if ($targetEmail === '' || !filter_var($targetEmail, FILTER_VALIDATE_EMAIL)) {
            jsonResponse(['ok' => false, 'message' => 'Debes enviar un correo valido.'], 422);
        }

        $targetUser = fetchOne($pdo, 'SELECT id, role, email FROM users WHERE email = :email LIMIT 1', ['email' => $targetEmail]);
        if ($targetUser === null) {
            jsonResponse(['ok' => false, 'message' => 'No existe una cuenta con ese correo.'], 404);
        }

        if ((string) ($targetUser['role'] ?? '') === 'admin') {
            jsonResponse(['ok' => true, 'message' => 'Ese usuario ya es admin.']);
        }

        executeStatement($pdo, 'UPDATE users SET role = :role WHERE id = :id', [
            'role' => 'admin',
            'id' => (int) $targetUser['id'],
        ]);

        jsonResponse([
            'ok' => true,
            'message' => 'Cuenta promovida a admin correctamente.',
            'admin_users' => fetchAll($pdo, 'SELECT id, full_name, email, role FROM users WHERE role = :role ORDER BY id ASC', ['role' => 'admin']),
        ]);
        break;

    case 'profile_2fa_generate':
        ensureTwoFactorSchemaReady($pdo);

        $secret = generateTwoFactorSecret();
        $_SESSION['_2fa_setup_user_id'] = (int) $user['id'];
        $_SESSION['_2fa_setup_secret'] = $secret;

        executeStatement(
            $pdo,
            'UPDATE users
             SET two_factor_setup_secret = :setup_secret,
                 two_factor_setup_expires_at = DATE_ADD(NOW(), INTERVAL 15 MINUTE)
             WHERE id = :id',
            [
                'setup_secret' => $secret,
                'id' => (int) $user['id'],
            ]
        );

        $label = urlencode('Azure Task Suite (' . (string) $user['email'] . ')');
        $issuer = urlencode('Azure Task Suite');
        $otpauthUrl = 'otpauth://totp/' . $label . '?secret=' . $secret . '&issuer=' . $issuer;

        jsonResponse([
            'ok' => true,
            'secret' => $secret,
            'otpauth_url' => $otpauthUrl,
            'qr_url' => 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=' . urlencode($otpauthUrl),
            'message' => 'Escanea el QR y confirma con un codigo para activar 2FA.',
        ]);
        break;

    case 'profile_2fa_enable':
        ensureTwoFactorSchemaReady($pdo);

        $rawRequestBody = file_get_contents('php://input');
        $rawRequestCode = '';
        if (is_string($rawRequestBody) && $rawRequestBody !== '') {
            parse_str($rawRequestBody, $rawParsed);
            if (is_array($rawParsed)) {
                $rawRequestCode = (string) ($rawParsed['code'] ?? '');
            }
        }

        $code = trim((string) (
            $input['code']
            ?? $_POST['code']
            ?? $_GET['code']
            ?? $_REQUEST['code']
            ?? $rawRequestCode
            ?? ''
        ));
        $setupUserId = (int) ($_SESSION['_2fa_setup_user_id'] ?? 0);
        $setupSecret = (string) ($_SESSION['_2fa_setup_secret'] ?? '');

        $setupDbRow = fetchOne(
            $pdo,
            'SELECT two_factor_setup_secret, two_factor_setup_expires_at
             FROM users
             WHERE id = :id
             LIMIT 1',
            ['id' => (int) $user['id']]
        );

        $setupSecretFromDb = '';
        if (is_array($setupDbRow)) {
            $dbSecret = (string) ($setupDbRow['two_factor_setup_secret'] ?? '');
            $dbExpiresAt = (string) ($setupDbRow['two_factor_setup_expires_at'] ?? '');
            $isDbSecretUsable = $dbSecret !== '' && (
                $dbExpiresAt === '' ||
                strtotime($dbExpiresAt) === false ||
                strtotime($dbExpiresAt) >= time()
            );

            if ($isDbSecretUsable) {
                $setupSecretFromDb = $dbSecret;
            }
        }

        if ($setupSecret === '' && $setupSecretFromDb !== '') {
            $setupSecret = $setupSecretFromDb;
            $setupUserId = (int) $user['id'];
        }

        if (!$code) {
            jsonResponse(['ok' => false, 'message' => 'Ingresa el codigo 2FA de 6 digitos.'], 422);
        }

        if ($setupUserId !== (int) $user['id'] || $setupSecret === '') {
            jsonResponse(['ok' => false, 'message' => 'Primero genera y escanea un QR de 2FA. Si el QR expiro, genera uno nuevo.'], 422);
        }

        if (!verifyTwoFactorCode($setupSecret, $code)) {
            jsonResponse(['ok' => false, 'message' => 'Codigo 2FA invalido. Verifica que tu telefono este sincronizado con la hora correcta.'], 422);
        }

        try {
            $recoveryCodes = generateTwoFactorRecoveryCodes();

            $pdo->beginTransaction();
            try {
                executeStatement(
                    $pdo,
                    'UPDATE users
                     SET two_factor_secret = :secret,
                         two_factor_enabled = 1,
                         two_factor_setup_secret = NULL,
                         two_factor_setup_expires_at = NULL
                     WHERE id = :id',
                    [
                        'secret' => $setupSecret,
                        'id' => (int) $user['id'],
                    ]
                );
                storeTwoFactorRecoveryCodes($pdo, (int) $user['id'], $recoveryCodes);
                $pdo->commit();
            } catch (Throwable $throwable) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                throw $throwable;
            }

            unset($_SESSION['_2fa_setup_user_id'], $_SESSION['_2fa_setup_secret']);

            $updatedUser = $user;
            $updatedUser['two_factor_enabled'] = true;

            jsonResponse([
                'ok' => true,
                'message' => '2FA activado correctamente. Guarda tus codigos de recuperacion en un lugar seguro.',
                'user' => $updatedUser,
                'recovery_codes' => $recoveryCodes,
            ]);
        } catch (Throwable $e) {
            jsonResponse([
                'ok' => false,
                'message' => 'Error al activar 2FA: ' . ($e->getMessage() ?? 'Unknown error')
            ], 500);
        }
        break;

    case 'profile_2fa_disable':
        ensureTwoFactorSchemaReady($pdo);

        $pdo->beginTransaction();
        try {
            executeStatement(
                $pdo,
                'UPDATE users
                 SET two_factor_secret = NULL,
                     two_factor_enabled = 0,
                     two_factor_setup_secret = NULL,
                     two_factor_setup_expires_at = NULL
                 WHERE id = :id',
                [
                    'id' => (int) $user['id'],
                ]
            );
            executeStatement($pdo, 'DELETE FROM two_factor_recovery_codes WHERE user_id = :user_id', [
                'user_id' => (int) $user['id'],
            ]);
            $pdo->commit();
        } catch (Throwable $throwable) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            throw $throwable;
        }

        unset($_SESSION['_2fa_setup_user_id'], $_SESSION['_2fa_setup_secret']);

        $updatedUser = $user;
        $updatedUser['two_factor_enabled'] = false;

        jsonResponse([
            'ok' => true,
            'message' => '2FA desactivado.',
            'user' => $updatedUser,
        ]);
        break;

    case 'profile_update':
        $fullName = trim((string) ($input['full_name'] ?? ''));
        $email = strtolower(trim((string) ($input['email'] ?? '')));
        $password = trim((string) ($input['password'] ?? ''));

        if ($fullName === '' || $email === '') {
            jsonResponse(['ok' => false, 'message' => 'Nombre y correo son obligatorios.'], 422);
        }

        if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
            jsonResponse(['ok' => false, 'message' => 'El correo ingresado no es valido.'], 422);
        }

        $emailOwnerId = (int) fetchScalar($pdo, 'SELECT COALESCE(MAX(id), 0) FROM users WHERE email = :email AND id <> :id', [
            'email' => $email,
            'id' => $user['id'],
        ]);
        if ($emailOwnerId > 0) {
            jsonResponse(['ok' => false, 'message' => 'Ya existe otra cuenta con ese correo.'], 422);
        }

        if ($password !== '' && strlen($password) < 8) {
            jsonResponse(['ok' => false, 'message' => 'La nueva contrasena debe tener al menos 8 caracteres.'], 422);
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

        jsonResponse([
            'ok' => true,
            'message' => 'Perfil actualizado.',
            'user' => currentUser(),
        ]);
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

    case 'document_security_status':
        $security = fetchOne($pdo, 'SELECT user_id, is_enabled, last_verified_at FROM user_document_security WHERE user_id = :user_id LIMIT 1', [
            'user_id' => (int) $user['id'],
        ]);

        jsonResponse([
            'ok' => true,
            'security' => [
                'configured' => $security !== null,
                'enabled' => $security !== null && (int) ($security['is_enabled'] ?? 0) === 1,
                'last_verified_at' => $security['last_verified_at'] ?? null,
            ],
        ]);
        break;

    case 'document_security_set':
        $accessKey = trim((string) ($input['access_key'] ?? ''));
        if (strlen($accessKey) < 6) {
            jsonResponse(['ok' => false, 'message' => 'La clave secundaria debe tener al menos 6 caracteres.'], 422);
        }

        $hash = password_hash($accessKey, PASSWORD_BCRYPT);
        executeStatement(
            $pdo,
            'INSERT INTO user_document_security (user_id, access_key_hash, is_enabled)
             VALUES (:user_id, :access_key_hash, 1)
             ON DUPLICATE KEY UPDATE access_key_hash = VALUES(access_key_hash), is_enabled = 1, updated_at = CURRENT_TIMESTAMP',
            [
                'user_id' => (int) $user['id'],
                'access_key_hash' => $hash,
            ]
        );

        jsonResponse(['ok' => true, 'message' => 'Clave secundaria guardada correctamente.']);
        break;

    case 'document_security_remove':
        executeStatement($pdo, 'DELETE FROM user_document_security WHERE user_id = :user_id', [
            'user_id' => (int) $user['id'],
        ]);

        unset($_SESSION['doc_access_user_id'], $_SESSION['doc_access_verified_until']);

        jsonResponse(['ok' => true, 'message' => 'Clave secundaria eliminada.']);
        break;

    default:
        jsonResponse(['ok' => false, 'message' => 'Accion no soportada.'], 400);
}
} catch (Throwable $throwable) {
    error_log(
        'dashboard.php error [action=' . ($action ?? 'unknown') . ']: ' .
        $throwable->getMessage() . ' in ' . $throwable->getFile() . ':' . $throwable->getLine()
    );

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

function userCanAccessNote(PDO $pdo, int $noteId, int $userId, string $userEmail): bool
{
    $owner = fetchOne(
        $pdo,
        'SELECT id FROM notes WHERE id = :id AND user_id = :user_id LIMIT 1',
        ['id' => $noteId, 'user_id' => $userId]
    );
    if ($owner !== null) {
        return true;
    }

    if ($userEmail === '') {
        return false;
    }

    $shared = fetchOne(
        $pdo,
        'SELECT id
         FROM note_shares
         WHERE note_id = :note_id
           AND is_active = 1
           AND LOWER(invited_email) = :invited_email
         LIMIT 1',
        [
            'note_id' => $noteId,
            'invited_email' => strtolower($userEmail),
        ]
    );

    return $shared !== null;
}

function sendOutboundJsonRequest(string $url, string $jsonPayload, int $timeoutSeconds = 15): array
{
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        if ($ch !== false) {
            curl_setopt_array($ch, [
                CURLOPT_RETURNTRANSFER => true,
                CURLOPT_POST => true,
                CURLOPT_HTTPHEADER => ['Content-Type: application/json'],
                CURLOPT_POSTFIELDS => $jsonPayload,
                CURLOPT_TIMEOUT => $timeoutSeconds,
            ]);

            $body = curl_exec($ch);
            $curlError = curl_error($ch);
            $httpCode = (int) curl_getinfo($ch, CURLINFO_HTTP_CODE);
            curl_close($ch);

            $statusLine = $httpCode > 0 ? ('HTTP ' . $httpCode) : '';
            $responseBody = is_string($body) ? trim($body) : '';

            return [
                'ok' => $httpCode >= 200 && $httpCode < 300,
                'status_line' => $statusLine,
                'body' => $responseBody,
                'transport_error' => trim($curlError),
            ];
        }
    }

    $ctx = stream_context_create([
        'http' => [
            'method' => 'POST',
            'header' => "Content-Type: application/json\r\n",
            'content' => $jsonPayload,
            'timeout' => $timeoutSeconds,
            'ignore_errors' => true,
        ],
    ]);

    $body = @file_get_contents($url, false, $ctx);
    $headers = $http_response_header ?? [];
    $statusLine = (string) ($headers[0] ?? '');
    $responseBody = is_string($body) ? trim($body) : '';

    return [
        'ok' => (bool) preg_match('/HTTP\/\S+ 2/', $statusLine),
        'status_line' => $statusLine,
        'body' => $responseBody,
        'transport_error' => is_string($body) ? '' : 'Sin respuesta HTTP desde la URL configurada.',
    ];
}

function buildPowerAutomateErrorMessage(string $baseMessage, array $result): string
{
    $details = [];

    if (!empty($result['status_line'])) {
        $details[] = (string) $result['status_line'];
    }

    if (!empty($result['transport_error'])) {
        $details[] = (string) $result['transport_error'];
    }

    if (!empty($result['body'])) {
        $details[] = mb_substr((string) $result['body'], 0, 220);
    }

    if ($details === []) {
        return $baseMessage;
    }

    return $baseMessage . ' Detalle: ' . implode(' | ', $details);
}

function financeSimulatorLabel(string $simulatorType): string
{
    $labels = [
        'compound_interest' => 'Interes compuesto',
        'home_capacity' => 'Capacidad de vivienda',
        'leasing_vs_credit' => 'Leasing vs credito',
        'tax_burden' => 'Carga tributaria',
    ];

    return $labels[$simulatorType] ?? 'Simulador financiero';
}

function flattenFinancePdfFields(array $payload, string $prefix = ''): array
{
    $lines = [];
    foreach ($payload as $key => $value) {
        $safeKey = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', (string) $key);
        $path = $prefix !== '' ? ($prefix . '.' . $safeKey) : $safeKey;

        if (is_array($value)) {
            foreach (flattenFinancePdfFields($value, $path) as $nestedLine) {
                $lines[] = $nestedLine;
            }
            continue;
        }

        if (is_bool($value)) {
            $value = $value ? 'true' : 'false';
        }

        if ($value === null) {
            $value = 'null';
        }

        $lineValue = mb_substr((string) $value, 0, 120);
        $lines[] = $path . ': ' . $lineValue;
    }

    if (count($lines) > 40) {
        $lines = array_slice($lines, 0, 40);
        $lines[] = '... (truncado por limite de reporte)';
    }

    return $lines;
}

function buildSimpleFinancePdf(array $rawLines): string
{
    $lines = [];
    foreach ($rawLines as $line) {
        $normalized = preg_replace('/[^\x20-\x7E]/', ' ', (string) $line);
        $normalized = trim((string) preg_replace('/\s+/', ' ', (string) $normalized));
        if ($normalized === '') {
            continue;
        }
        $lines[] = mb_substr($normalized, 0, 110);
    }

    if ($lines === []) {
        $lines = ['Sin contenido para exportar.'];
    }

    $maxLines = 46;
    if (count($lines) > $maxLines) {
        $lines = array_slice($lines, 0, $maxLines);
        $lines[] = '... contenido truncado';
    }

    $content = "BT\n/F1 10 Tf\n50 800 Td\n14 TL\n";
    foreach ($lines as $index => $line) {
        $escaped = str_replace(['\\', '(', ')'], ['\\\\', '\\(', '\\)'], $line);
        if ($index === 0) {
            $content .= '(' . $escaped . ") Tj\n";
        } else {
            $content .= 'T* (' . $escaped . ") Tj\n";
        }
    }
    $content .= "ET\n";

    $objects = [];
    $objects[] = '<< /Type /Catalog /Pages 2 0 R >>';
    $objects[] = '<< /Type /Pages /Kids [3 0 R] /Count 1 >>';
    $objects[] = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>';
    $objects[] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
    $objects[] = '<< /Length ' . strlen($content) . " >>\nstream\n" . $content . 'endstream';

    $pdf = "%PDF-1.4\n";
    $offsets = [0];

    foreach ($objects as $index => $objectBody) {
        $offsets[] = strlen($pdf);
        $objNumber = $index + 1;
        $pdf .= $objNumber . " 0 obj\n" . $objectBody . "\nendobj\n";
    }

    $xrefOffset = strlen($pdf);
    $pdf .= 'xref' . "\n";
    $pdf .= '0 ' . (count($objects) + 1) . "\n";
    $pdf .= "0000000000 65535 f \n";

    for ($i = 1; $i <= count($objects); $i++) {
        $pdf .= str_pad((string) $offsets[$i], 10, '0', STR_PAD_LEFT) . " 00000 n \n";
    }

    $pdf .= 'trailer << /Size ' . (count($objects) + 1) . ' /Root 1 0 R >>' . "\n";
    $pdf .= 'startxref' . "\n" . $xrefOffset . "\n%%EOF";

    return $pdf;
}