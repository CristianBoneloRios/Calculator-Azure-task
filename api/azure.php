<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

$backendVersion = '2026-05-03.2';

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

$configPath = __DIR__ . '/config.php';
if (is_file($configPath)) {
    require_once $configPath;
}

$rawBody = file_get_contents('php://input');
$payload = json_decode((string) $rawBody, true);

if (!is_array($payload)) {
    respondError(400, 'Body JSON invalido.');
}

$action = trim((string) ($payload['action'] ?? 'fetch'));

if ($action === 'status') {
    respondOk([
        'hasPat' => getServerPat() !== '',
        'backendVersion' => $backendVersion,
        'wiqlMode' => 'query+url-$top'
    ]);
}

if ($action === 'delete_pat') {
    if (!deleteServerPat()) {
        respondError(500, 'No se pudo borrar el PAT del servidor.');
    }

    respondOk([
        'message' => 'PAT eliminado del servidor.'
    ]);
}

$org = trim((string) ($payload['org'] ?? ''));
$project = normalizeProjectName((string) ($payload['project'] ?? ''));
$incomingPat = trim((string) ($payload['pat'] ?? ''));

$maxItems = (int) ($payload['maxItems'] ?? 10000);
if ($maxItems < 1) {
    $maxItems = 10000;
}
if ($maxItems > 20000) {
    $maxItems = 20000;
}

if ($incomingPat !== '') {
    if (preg_match('/\s/', $incomingPat) === 1) {
        respondError(400, 'El PAT contiene espacios o saltos de linea.');
    }
    if (!persistServerPat($incomingPat)) {
        respondError(500, 'No se pudo guardar el PAT en servidor. Verifica permisos de escritura en api/storage.');
    }
}

$pat = getServerPat();

if ($org === '' || $project === '' || $pat === '') {
    respondError(400, 'org y project son requeridos, y el PAT debe configurarse en servidor.');
}

if (!preg_match('/^[a-zA-Z0-9_-]+$/', $org)) {
    respondError(400, 'Organizacion invalida.');
}

$orgUrl = 'https://dev.azure.com/' . $org;
$projectPath = rawurlencode($project);
$authHeader = 'Authorization: Basic ' . base64_encode(':' . $pat);

$wiql = 'SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @project ORDER BY [System.ChangedDate] DESC';
$wiqlUrl = $orgUrl . '/' . $projectPath . '/_apis/wit/wiql?api-version=7.0&%24top=' . $maxItems;

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
        if (strpos($azureMsg, 'VS402337') !== false) {
            respondError(400, 'El proyecto tiene demasiados items. Reduce el alcance de consulta (por sprint/estado) o baja maxItems.');
        }
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
    $assignedEmail = '';

    if (is_array($assignedTo)) {
        $assignedEmail = (string) ($assignedTo['uniqueName'] ?? $assignedTo['mailAddress'] ?? '');
        $assignedTo = (string) ($assignedTo['displayName'] ?? $assignedEmail ?? '');
    } else {
        $assignedTo = (string) $assignedTo;
    }

    return [
        'ID' => (string) ($fields['System.Id'] ?? ''),
        'Titulo' => (string) ($fields['System.Title'] ?? ''),
        'Tipo' => (string) ($fields['System.WorkItemType'] ?? ''),
        'Estado' => (string) ($fields['System.State'] ?? ''),
        'Asignado a' => $assignedTo,
        'Asignado correo' => $assignedEmail,
        'Creado' => (string) ($fields['System.CreatedDate'] ?? ''),
        'Actualizado' => (string) ($fields['System.ChangedDate'] ?? ''),
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
    'count' => count($rows),
    'limitApplied' => $maxItems,
    'backendVersion' => $backendVersion
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

function getServerPat(): string
{
    $stored = readServerPat();
    if ($stored !== '') {
        return $stored;
    }

    if (defined('AZURE_DEVOPS_PAT') && is_string(AZURE_DEVOPS_PAT) && trim(AZURE_DEVOPS_PAT) !== '') {
        return trim(AZURE_DEVOPS_PAT);
    }

    $env = getenv('AZURE_DEVOPS_PAT');
    if (is_string($env) && trim($env) !== '') {
        return trim($env);
    }

    return '';
}

function getPatStoragePath(): string
{
    return __DIR__ . '/storage/pat.store.php';
}

function persistServerPat(string $pat): bool
{
    $dir = __DIR__ . '/storage';
    if (!is_dir($dir) && !mkdir($dir, 0700, true) && !is_dir($dir)) {
        return false;
    }

    $path = getPatStoragePath();
    $content = "<?php\nreturn " . var_export($pat, true) . ";\n";
    $written = @file_put_contents($path, $content, LOCK_EX);
    if ($written === false) {
        return false;
    }

    @chmod($path, 0600);
    return true;
}

function readServerPat(): string
{
    $path = getPatStoragePath();
    if (!is_file($path)) {
        return '';
    }

    $value = require $path;
    if (!is_string($value)) {
        return '';
    }

    return trim($value);
}

function deleteServerPat(): bool
{
    $path = getPatStoragePath();
    if (!is_file($path)) {
        return true;
    }

    return @unlink($path);
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
