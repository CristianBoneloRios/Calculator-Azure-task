<?php

declare(strict_types=1);

// Copy this file to api/config.php and set your real PAT.
// Never commit api/config.php with real secrets.
define('AZURE_DEVOPS_PAT', 'PUT_YOUR_AZURE_PAT_HERE');

// Copilot 365 Agent bridge (for api/copilot_agent.php)
define('COPILOT_AGENT_API_KEY', 'PUT_A_LONG_RANDOM_KEY_HERE');
define('COPILOT_ALLOWED_ORIGIN', '*');
