<?php

declare(strict_types=1);

require_once __DIR__ . '/../api/app.php';

ensureApplicationInstalled();
logoutCurrentUser();

header('Location: login.php');
exit;