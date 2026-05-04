<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

// Optional CORS support for custom domains/subdomains if needed.
$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin !== '') {
    header('Vary: Origin');
    header('Access-Control-Allow-Origin: ' . $origin);
    header('Access-Control-Allow-Headers: Content-Type, Authorization');
    header('Access-Control-Allow-Methods: POST, OPTIONS');
}

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    respondError(405, 'Metodo no permitido. Usa POST.');
}

$rawBody = file_get_contents('php://input');
$payload = json_decode((string) $rawBody, true);

if (!is_array($payload)) {
    respondError(400, 'Body JSON invalido.');
}

$org = trim((string) ($payload['org'] ?? ''));
$project = normalizeProjectName((string) ($payload['project'] ?? ''));
$pat = trim((string) ($payload['pat'] ?? ''));

if ($org === '' || $project === '' || $pat === '') {
    respondError(400, 'org, project y pat son requeridos.');
}

if (!preg_match('/^[a-zA-Z0-9_-]+$/', $org)) {
    respondError(400, 'Organizacion invalida.');
}

$orgUrl = 'https://dev.azure.com/' . $org;
$projectPath = rawurlencode($project);
$authHeader = 'Authorization: Basic ' . base64_encode(':' . $pat);

$wiql = 'SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @project ORDER BY [System.ChangedDate] DESC';
$wiqlUrl = $orgUrl . '/' . $projectPath . '/_apis/wit/wiql?api-version=7.0';

$wiqlResponse = azureRequest($wiqlUrl, 'POST', [
    'query' => $wiql
], [
    $authHeader,
    'Content-Type: application/json'
]);

if ($wiqlResponse['status'] < 200 || $wiqlResponse['status'] >= 300) {
    $azureMsg = extractAzureMessage($wiqlResponse['body']);

    if ($wiqlResponse['status'] === 401) {
        respondError(401, 'PAT invalido o expirado.');
    }
    if ($wiqlResponse['status'] === 403) {
        respondError(403, 'Acceso denegado. Verifica permisos del PAT.');
    }
    if ($wiqlResponse['status'] === 404) {
        respondError(404, 'Proyecto u organizacion no encontrada.');
    }
    if ($wiqlResponse['status'] === 400) {
        respondError(400, $azureMsg !== '' ? $azureMsg : 'Consulta WIQL invalida.');
    }

    respondError(502, $azureMsg !== '' ? $azureMsg : 'Error al consultar Azure DevOps.');
}

$wiqlBody = json_decode($wiqlResponse['body'], true);
$workItems = is_array($wiqlBody['workItems'] ?? null) ? $wiqlBody['workItems'] : [];

if (count($workItems) === 0) {
    respondOk([
        'rows' => []
    ]);
}

$ids = [];
foreach ($workItems as $item) {
    if (isset($item['id']) && is_numeric($item['id'])) {
        $ids[] = (int) $item['id'];
    }
}

if (count($ids) === 0) {
    respondOk([
        'rows' => []
    ]);
}

$allWorkItems = [];
$chunks = array_chunk($ids, 200);

foreach ($chunks as $chunk) {
    $idsParam = implode(',', $chunk);
    $detailsUrl = $orgUrl . '/_apis/wit/workitems?ids=' . $idsParam . '&%24expand=all&api-version=7.0';

    $detailResponse = azureRequest($detailsUrl, 'GET', null, [
        $authHeader,
        'Content-Type: application/json'
    ]);

    if ($detailResponse['status'] < 200 || $detailResponse['status'] >= 300) {
        $azureMsg = extractAzureMessage($detailResponse['body']);
        respondError(502, $azureMsg !== '' ? $azureMsg : 'No se pudieron obtener los detalles de tareas.');
    }

    $detailBody = json_decode($detailResponse['body'], true);
    $items = is_array($detailBody['value'] ?? null) ? $detailBody['value'] : [];
    foreach ($items as $it) {
        $allWorkItems[] = $it;
    }
}

$rows = array_map(static function (array $wi): array {
    $fields = is_array($wi['fields'] ?? null) ? $wi['fields'] : [];
    $assignedTo = $fields['System.AssignedTo'] ?? '';

    if (is_array($assignedTo)) {
        $assignedTo = (string) ($assignedTo['displayName'] ?? $assignedTo['uniqueName'] ?? '');
    } else {
        $assignedTo = (string) $assignedTo;
    }

    return [
        'ID' => (string) ($fields['System.Id'] ?? ''),
        'Titulo' => (string) ($fields['System.Title'] ?? ''),
        'Tipo' => (string) ($fields['System.WorkItemType'] ?? ''),
        'Estado' => (string) ($fields['System.State'] ?? ''),
        'Asignado a' => $assignedTo,
        'Estimacion Original' => (float) ($fields['Microsoft.VSTS.Scheduling.OriginalEstimate'] ?? 0),
        'Trabajo Completado' => (float) ($fields['Microsoft.VSTS.Scheduling.CompletedWork'] ?? 0),
        'Trabajo Restante' => (float) ($fields['Microsoft.VSTS.Scheduling.RemainingWork'] ?? 0),
        'Etiquetas' => (string) ($fields['System.Tags'] ?? ''),
        'Ruta de Area' => (string) ($fields['System.AreaPath'] ?? ''),
        'Iteracion' => (string) ($fields['System.IterationPath'] ?? '')
    ];
}, $allWorkItems);

respondOk([
    'rows' => $rows,
    'count' => count($rows)
]);

function normalizeProjectName(string $project): string
{
    $value = trim($project);

    for ($i = 0; $i < 2; $i++) {
        $decoded = rawurldecode($value);
        if ($decoded === $value) {
            break;
        }
        $value = $decoded;
    }

    return $value;
}

function azureRequest(string $url, string $method, ?array $body, array $headers): array
{
    $ch = curl_init($url);
    if ($ch === false) {
        respondError(500, 'No se pudo inicializar cURL.');
    }

    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $method);
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_TIMEOUT, 45);

    if ($body !== null) {
        $encoded = json_encode($body, JSON_UNESCAPED_UNICODE);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $encoded);
    }

    $resp = curl_exec($ch);
    if ($resp === false) {
        $err = curl_error($ch);
        curl_close($ch);
        respondError(502, 'Error de red al conectar con Azure DevOps: ' . $err);
    }

    $status = (int) curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    return [
        'status' => $status,
        'body' => (string) $resp
    ];
}

function extractAzureMessage(string $body): string
{
    $json = json_decode($body, true);
    if (!is_array($json)) {
        return '';
    }

    $message = $json['message'] ?? '';
    return is_string($message) ? trim($message) : '';
}

function respondOk(array $data): void
{
    echo json_encode([
        'ok' => true
    ] + $data, JSON_UNESCAPED_UNICODE);
    exit;
}

function respondError(int $status, string $message): void
{
    http_response_code($status);
    echo json_encode([
        'ok' => false,
        'message' => $message
    ], JSON_UNESCAPED_UNICODE);
    exit;
}
