# Base de Datos Completa y Detallada

Este documento contiene el esquema SQL completo consolidado del sistema.

## Motor y codificacion recomendada

- Motor: InnoDB
- Charset: utf8mb4
- Collation: utf8mb4_unicode_ci

## Script SQL completo

```sql
-- =============================================================
-- Schema completo — u400335795_AzureDevOPs
-- Todas las tablas con CREATE TABLE IF NOT EXISTS
-- =============================================================

-- -------------------------------------------------------------
-- Tabla base: users
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS users (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    full_name VARCHAR(150) NOT NULL,
    email VARCHAR(190) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    two_factor_secret VARCHAR(255) DEFAULT NULL,
    two_factor_enabled TINYINT(1) NOT NULL DEFAULT 0,
    role VARCHAR(50) NOT NULL DEFAULT 'member',
    profile_photo_path VARCHAR(255) DEFAULT NULL,
    last_login_at DATETIME DEFAULT NULL,
    last_seen_at DATETIME DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Sesiones de usuario
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS user_sessions (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    session_token VARCHAR(128) NOT NULL UNIQUE,
    login_at DATETIME NOT NULL,
    logout_at DATETIME DEFAULT NULL,
    last_seen_at DATETIME NOT NULL,
    ip_address VARCHAR(45) DEFAULT NULL,
    user_agent VARCHAR(255) DEFAULT NULL,
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY fk_user_sessions_user (user_id),
    CONSTRAINT fk_user_sessions_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- 2FA: codigos de recuperacion
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS two_factor_recovery_codes (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    code_hash VARCHAR(255) NOT NULL,
    used_at DATETIME DEFAULT NULL,
    is_valid TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_user_id (user_id),
    KEY idx_used_at (used_at),
    CONSTRAINT fk_two_factor_recovery_codes_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Perfil publico
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public_profiles (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    slug VARCHAR(120) NOT NULL UNIQUE,
    display_name VARCHAR(150) NOT NULL,
    role_title VARCHAR(150) NOT NULL,
    company_name VARCHAR(150) DEFAULT NULL,
    bio TEXT DEFAULT NULL,
    photo_path VARCHAR(255) DEFAULT NULL,
    updated_by_user_id INT UNSIGNED DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_public_profiles_user (updated_by_user_id),
    CONSTRAINT fk_public_profiles_user FOREIGN KEY (updated_by_user_id) REFERENCES users(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Historial de cambios de foto de perfil
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS profile_photo_changes (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    changed_by_user_id INT UNSIGNED DEFAULT NULL,
    file_path VARCHAR(255) NOT NULL,
    original_name VARCHAR(255) DEFAULT NULL,
    mime_type VARCHAR(120) DEFAULT NULL,
    file_size INT UNSIGNED DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY fk_profile_photo_changes_user (user_id),
    KEY fk_profile_photo_changes_changed_by (changed_by_user_id),
    CONSTRAINT fk_profile_photo_changes_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    CONSTRAINT fk_profile_photo_changes_changed_by FOREIGN KEY (changed_by_user_id) REFERENCES users(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Notas
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS notes (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    title VARCHAR(180) NOT NULL,
    content MEDIUMTEXT NOT NULL,
    color VARCHAR(20) NOT NULL DEFAULT 'blue',
    is_pinned TINYINT(1) NOT NULL DEFAULT 0,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_notes_user (user_id),
    CONSTRAINT fk_notes_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Tareas diarias
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS daily_tasks (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    task_date DATE NOT NULL,
    title VARCHAR(180) NOT NULL,
    description TEXT DEFAULT NULL,
    status VARCHAR(30) NOT NULL DEFAULT 'pending',
    priority VARCHAR(20) NOT NULL DEFAULT 'medium',
    due_time TIME DEFAULT NULL,
    sort_order INT NOT NULL DEFAULT 0,
    completed_at DATETIME DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_daily_tasks_user (user_id),
    CONSTRAINT fk_daily_tasks_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Metas / Objetivos
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS goals (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    title VARCHAR(180) NOT NULL,
    description TEXT DEFAULT NULL,
    target_date DATE DEFAULT NULL,
    progress_percent TINYINT UNSIGNED NOT NULL DEFAULT 0,
    status VARCHAR(30) NOT NULL DEFAULT 'active',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_goals_user (user_id),
    CONSTRAINT fk_goals_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Fuentes de calendario
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS calendar_sources (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    provider VARCHAR(50) NOT NULL,
    external_account_email VARCHAR(190) DEFAULT NULL,
    sync_enabled TINYINT(1) NOT NULL DEFAULT 0,
    sync_status VARCHAR(40) NOT NULL DEFAULT 'pending',
    last_synced_at DATETIME DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_calendar_sources_user (user_id),
    CONSTRAINT fk_calendar_sources_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Eventos de calendario
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS calendar_events (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    source_id INT UNSIGNED DEFAULT NULL,
    external_event_id VARCHAR(190) DEFAULT NULL,
    title VARCHAR(180) NOT NULL,
    description TEXT DEFAULT NULL,
    start_at DATETIME NOT NULL,
    end_at DATETIME NOT NULL,
    location VARCHAR(190) DEFAULT NULL,
    meeting_url VARCHAR(255) DEFAULT NULL,
    source_type VARCHAR(40) NOT NULL DEFAULT 'manual',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY fk_calendar_events_user (user_id),
    KEY fk_calendar_events_source (source_id),
    CONSTRAINT fk_calendar_events_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    CONSTRAINT fk_calendar_events_source FOREIGN KEY (source_id) REFERENCES calendar_sources(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Integraciones (Power Automate calendar, etc.)
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS integrations (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    provider VARCHAR(50) NOT NULL,
    status VARCHAR(40) NOT NULL DEFAULT 'pending',
    access_token TEXT DEFAULT NULL,
    refresh_token TEXT DEFAULT NULL,
    token_expires_at DATETIME DEFAULT NULL,
    metadata LONGTEXT DEFAULT NULL CHECK (JSON_VALID(metadata)),
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uniq_integration_provider_user (user_id, provider),
    CONSTRAINT fk_integrations_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Configuracion de Power Automate (token entrante + webhook saliente)
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS power_automate_config (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL UNIQUE,
    token VARCHAR(255) NOT NULL UNIQUE,
    token_generated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    webhook_url VARCHAR(500) NOT NULL,
    header_name VARCHAR(50) NOT NULL DEFAULT 'X-Power-Automate-Key',
    outbound_webhook_url VARCHAR(500) DEFAULT NULL,
    external_account_email VARCHAR(190) DEFAULT NULL,
    is_enabled TINYINT(1) NOT NULL DEFAULT 0,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY idx_user_id (user_id),
    CONSTRAINT fk_power_automate_config_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Logs de sincronizacion Power Automate
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS power_automate_logs (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id INT UNSIGNED NOT NULL,
    event_type ENUM('sync_request_sent','sync_response_received','token_rotated','webhook_error') NOT NULL,
    direction ENUM('inbound','outbound') NOT NULL,
    status_code INT(3) DEFAULT NULL,
    error_message TEXT DEFAULT NULL,
    events_count INT DEFAULT NULL,
    payload_preview LONGTEXT DEFAULT NULL CHECK (JSON_VALID(payload_preview)),
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_user_id (user_id),
    KEY idx_event_type (event_type),
    KEY idx_created_at (created_at),
    CONSTRAINT fk_power_automate_logs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Jobs de generacion de documentos (manual, guia, informe)
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS document_generation_jobs (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id BIGINT UNSIGNED NOT NULL,
    user_email VARCHAR(255) NOT NULL,
    input_file_name VARCHAR(255) DEFAULT NULL,
    input_file_type VARCHAR(50) DEFAULT NULL,
    input_file_size BIGINT DEFAULT NULL,
    input_file_path TEXT DEFAULT NULL,
    input_file_hash VARCHAR(255) DEFAULT NULL,
    input_base64 LONGTEXT DEFAULT NULL,
    generation_type VARCHAR(50) DEFAULT NULL,
    description TEXT DEFAULT NULL,
    pa_request_id VARCHAR(255) DEFAULT NULL,
    pa_status VARCHAR(50) DEFAULT NULL,
    pa_response TEXT DEFAULT NULL,
    output_file_name VARCHAR(255) DEFAULT NULL,
    output_file_type VARCHAR(50) DEFAULT NULL,
    output_file_path TEXT DEFAULT NULL,
    output_file_size BIGINT DEFAULT NULL,
    output_file_url TEXT DEFAULT NULL,
    ai_summary TEXT DEFAULT NULL,
    ai_confidence DECIMAL(5,2) DEFAULT NULL,
    status ENUM('pending','processing','completed','error') NOT NULL DEFAULT 'pending',
    error_message TEXT DEFAULT NULL,
    ip_address VARCHAR(45) DEFAULT NULL,
    user_agent TEXT DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    completed_at TIMESTAMP NULL DEFAULT NULL,
    KEY idx_user (user_id),
    KEY idx_status (status),
    KEY idx_created_at (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Seguridad de acceso a documentos (clave secundaria por usuario)
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS user_document_security (
    id BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    user_id BIGINT UNSIGNED NOT NULL UNIQUE,
    access_key_hash VARCHAR(255) NOT NULL,
    is_enabled TINYINT(1) NOT NULL DEFAULT 1,
    last_verified_at TIMESTAMP NULL DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY idx_user_doc_security_user (user_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -------------------------------------------------------------
-- Configuracion general de la aplicacion (clave-valor)
-- -------------------------------------------------------------
CREATE TABLE IF NOT EXISTS app_settings (
    id INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    setting_key VARCHAR(150) NOT NULL UNIQUE,
    setting_value TEXT DEFAULT NULL,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
```

## Verificacion rapida

```sql
SHOW TABLES;
```

```sql
SELECT COUNT(*) AS total_tablas
FROM information_schema.tables
WHERE table_schema = DATABASE();
```
