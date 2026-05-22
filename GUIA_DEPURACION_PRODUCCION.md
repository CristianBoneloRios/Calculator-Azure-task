# Guia de depuracion en produccion (Hostinger)

## Error reportado

`No resource with given URL found`

## 3 causas raiz mas probables

1. **Ruta principal o archivos estaticos no encontrados**
   - Ejemplos: `app.js`, `styles.css`, `workspace/login.php`, `workspace/assets/workspace.css`.
   - Sintoma: el HTML carga parcialmente y el navegador reporta recursos faltantes.

2. **Error interno PHP (500) en endpoints clave**
   - Ejemplos: `api/auth.php`, `api/public_profile.php`, `workspace/login.php`.
   - Sintoma: al navegar a login o al consultar APIs, aparece pagina de error y luego errores de recursos en consola.

3. **Configuracion de hosting incompleta para base de datos**
   - Ejemplos: `.env` ausente, `DB_HOST` incorrecto, `pdo_mysql` deshabilitado.
   - Sintoma: login no inicializa, APIs responden con error y la sesion no se crea.

## Flujo recomendado de depuracion

1. Abrir `https://TU_DOMINIO/api/health.php`.
2. Verificar estado de checks:
   - `pdo_mysql_extension`
   - `env_file`
   - `database_connection`
3. Abrir `https://TU_DOMINIO/workspace/login.php` y confirmar si muestra alerta de bootstrap.
4. Abrir DevTools (F12) en `https://TU_DOMINIO/index.php` y revisar:
   - Consola: tabla `Diagnostico de recursos`
   - Network: recursos con estado distinto de 200.
5. Corregir y volver a desplegar.

## Check list de Hostinger

- Subir archivo `.env` en la raiz real del sitio.
- Confirmar `DB_HOST` con el valor exacto de hPanel (no asumir `localhost`).
- Confirmar `DB_DATABASE`, `DB_USERNAME`, `DB_PASSWORD`.
- Confirmar extension `pdo_mysql` activa.
- Confirmar que `DirectoryIndex` prioriza `index.php`.

## Rutas utiles

- `/index.php`
- `/workspace/login.php`
- `/api/health.php`
- `/api/auth.php?action=session`
