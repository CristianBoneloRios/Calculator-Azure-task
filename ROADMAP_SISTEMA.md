# Roadmap del Workspace

## Objetivo general

Convertir el analizador actual en un workspace personal con autenticacion, persistencia en MySQL y base preparada para integraciones de productividad sin perder la apariencia visual principal del sistema.

## Alcance implementado en esta fase

- Conexion MySQL centralizada con `.env` en raiz.
- Esquema SQL autoinstalable para usuarios, sesiones, notas, tareas, metas, calendario, integraciones y perfil publico.
- Login y logout con sesiones PHP y auditoria de ultimo acceso.
- Workspace con hojas en PHP, HTML, Bootstrap, CSS y JavaScript:
  - `workspace/index.php`
  - `workspace/profile.php`
  - `workspace/notes.php`
  - `workspace/tasks.php`
  - `workspace/goals.php`
  - `workspace/calendar.php`
- Subida de foto de perfil del usuario y opcion para reutilizarla en el perfil publico del desarrollador.
- Boton de inicio de sesion en el sidebar principal con animacion cuando hay sesion activa.

## Modelo de datos pensado para el sistema

- `users`: credenciales, rol, foto, ultima sesion.
- `user_sessions`: login, logout, ultimo seen, IP y agente.
- `public_profiles`: datos publicos que hoy muestra la portada, incluyendo el bloque de Cristian Jesus Bonelo Rios.
- `profile_photo_changes`: historial de cambios de foto.
- `notes`: notas importantes tipo workspace.
- `daily_tasks`: tareas del dia y backlog operativo.
- `goals`: metas con avance porcentual.
- `calendar_sources`: origen de datos para Teams o Calendar.
- `calendar_events`: eventos manuales o sincronizados.
- `integrations`: tokens y metadatos por proveedor.
- `app_settings`: ajustes globales del sistema.

## Siguiente fase recomendada

1. Crear aplicacion en Azure AD / Microsoft Entra ID para usar Microsoft Graph.
2. Agregar `MICROSOFT_CLIENT_ID`, `MICROSOFT_CLIENT_SECRET`, `MICROSOFT_TENANT_ID` y `MICROSOFT_REDIRECT_URI` al `.env`.
3. Implementar OAuth para Microsoft 365 y persistir tokens en `integrations`.
4. Sincronizar calendario de Teams/Outlook hacia `calendar_events` y actualizar `calendar_sources`.
5. Sincronizar reuniones programadas del dia y mostrarlas en el dashboard principal.

## Riesgos tecnicos a controlar

- El `.env` contiene secretos y no debe subirse al repositorio.
- La subida de imagenes requiere permisos de escritura sobre `uploads/profile-photos`.
- La sincronizacion real con Teams no depende de la base de datos sino de credenciales OAuth y consentimiento de Microsoft Graph.
- En hosting compartido conviene confirmar que `PDO`, `pdo_mysql` y `file_uploads` estan habilitados.