# Guía de Integración - Autenticación Biométrica Morpho E3/E4

## Descripción General

Este documento describe cómo utilizar la autenticación biométrica por huella digital en Azure Task Suite usando un captor Morpho E3 o E4.

## Características

✅ **Registro de Huella Digital**: Captura y almacenamiento seguro de huellas digitales  
✅ **Autenticación Rápida**: Inicia sesión con tu huella sin escribir contraseña  
✅ **Verificación Segura**: Los datos biométricos se almacenan como hashes seguros  
✅ **Historial de Auditoría**: Registro de todos los intentos de autenticación  
✅ **Gestión en Perfil**: Opción para registrar, reemplazar o eliminar tu huella

---

## Requisitos Previos

### Hardware
- **Captor Morpho E3 o E4** conectado vía USB
- **Computadora** con puertos USB disponibles

### Software

#### 1. Driver de Morpho
Descarga e instala el driver oficial de tu captor Morpho:
- [Morpho E3 Driver](https://www.morpho.com) (descarga desde el sitio oficial)
- [Morpho E4 Driver](https://www.morpho.com) (descarga desde el sitio oficial)

**Pasos de instalación:**
```
1. Descargar el archivo .exe o .msi del driver
2. Ejecutar el instalador
3. Seguir las instrucciones del asistente
4. Reiniciar la computadora
5. Conectar el captor Morpho
6. El sistema debería detectar el dispositivo automáticamente
```

#### 2. Middleware Morpho (Opcional)
Para aplicaciones web, es recomendable instalar el middleware de Morpho:
```
1. Instalar Morpho WebEnroll SDK (si está disponible)
2. Configurar el servicio en localhost:5000
3. Permitir acceso CORS desde tu aplicación
```

#### 3. Browser Compatibility
- Chrome/Chromium 90+
- Firefox 88+
- Edge 90+
- Opera 76+

---

## Configuración en Azure Task Suite

### Paso 1: Verificar la conexión del captor

1. Ve a tu **Perfil** (esquina superior derecha)
2. Desplázate hasta la sección **"Autenticacion por huella digital"**
3. Verifica el estado del dispositivo:
   - 🟢 **Verde**: Captor conectado correctamente
   - 🔴 **Rojo**: Captor no detectado

### Paso 2: Registrar tu huella digital

1. En la sección de Biometría, haz clic en **"Registrar huella digital"**
2. Aparecerá una pantalla indicando: **"Acerca tu dedo al captor"**
3. Coloca tu dedo en el captor y espera a que se complete la captura
4. La pantalla mostrará ✓ cuando la huella sea capturada exitosamente
5. Haz clic en **"Guardar huella digital"**
6. ¡Listo! Tu huella ha sido registrada

### Paso 3: Usar autenticación biométrica en el login

1. Ve a la página de **Login** (Iniciar sesión)
2. En la sección de biometría, verás un botón **"Autenticarse con huella"**
3. Ingresa tu **email**
4. Haz clic en el botón de huella digital
5. Acerca tu dedo al captor
6. Espera a que se verifique tu identidad
7. ¡Serás autenticado automáticamente!

---

## Gestión de la Huella Digital

### Reemplazar tu huella
1. Ve a tu **Perfil** → **Biometría**
2. Haz clic en **"Reemplazar huella digital"**
3. Sigue los pasos de captura nuevamente

### Eliminar tu huella
1. Ve a tu **Perfil** → **Biometría**
2. Haz clic en **"Eliminar huella"**
3. Confirma la eliminación
4. Tu huella será removida de forma permanente

### Ver información de tu huella
En el panel de Biometría puedes ver:
- **Dispositivo**: Tipo de captor (ej: Morpho E3/E4)
- **Número de verificaciones**: Cuántas veces has usado la autenticación biométrica
- **Última verificación**: La última fecha/hora que iniciaste sesión con tu huella

---

## Solución de Problemas

### ❌ "Captor no detectado"

**Causa posible**: El driver no está instalado o el captor no está conectado.

**Solución**:
1. Verifica que el captor esté conectado a un puerto USB
2. Intenta reconectar el cable USB
3. Reinicia tu navegador
4. Si persiste, reinstala el driver:
   - Desinstala el driver actual desde Panel de Control
   - Descarga e instala la versión más reciente
   - Reinicia la computadora

### ❌ "Tiempo de captura agotado"

**Causa posible**: La captura tardó más de 30 segundos.

**Solución**:
1. Asegúrate de que tu dedo esté limpio y seco
2. Coloca el dedo completamente en el sensor
3. Mantén el dedo estable durante la captura
4. Intenta nuevamente

### ❌ "La huella digital no coincide"

**Causa posible**: Tu huella no coincide con la registrada.

**Solución**:
1. Asegúrate de usar el mismo dedo que registraste
2. Presiona el dedo correctamente en el sensor
3. Intenta registrar nuevamente con mejor presión/posición
4. Como alternativa, usa tu contraseña para iniciar sesión

### ❌ "Error de base de datos"

**Causa posible**: Problema al guardar los datos.

**Solución**:
1. Recarga la página
2. Intenta registrarte nuevamente
3. Si persiste, contacta al administrador

---

## Medidas de Seguridad

### ¿Cómo se almacenan mis datos biométricos?

1. **Hashing SHA-256**: Tu huella se convierte en un hash criptográfico
2. **Almacenamiento seguro**: Los datos se guardan en la base de datos MySQL
3. **Sin datos sin cifrar**: Los templates de huellas se almacenan como BLOB codificado
4. **Auditoría**: Todos los intentos se registran con fecha, hora e IP

### ¿Qué sucede si alguien obtiene acceso a mi huella?

Los datos almacenados en Azure Task Suite son:
- **Hashes criptográficos** (no son la huella original)
- **Específicos de este sistema** (no pueden usarse en otro lugar)
- **Protegidos** por la seguridad de la base de datos

### ¿Puedo usar la autenticación biométrica junto a 2FA?

Sí. Si tienes ambas habilitadas:
1. **Primera opción**: Autenticarte con huella (más rápido)
2. **Alternativa**: Usar contraseña y después código 2FA

---

## Información Técnica

### Arquitectura

```
┌─────────────────────────────────────────┐
│        Browser / Cliente Web            │
│  - biometric-manager.js (SDK)           │
│  - profile-biometric.js (UI)            │
└────────────────┬────────────────────────┘
                 │ HTTPS API
┌────────────────▼────────────────────────┐
│      api/biometric.php                  │
│  - Captura: /action=capture-fingerprint │
│  - Verificación: /action=verify         │
│  - Guardado: /action=save-fingerprint   │
└────────────────┬────────────────────────┘
                 │ SQL
┌────────────────▼────────────────────────┐
│    Base de Datos MySQL (Tablas)         │
│  - user_biometric_fingerprints          │
│  - biometric_login_attempts (auditoría) │
└─────────────────────────────────────────┘
```

### Tablas de Base de Datos

#### `user_biometric_fingerprints`
```sql
- id: Identificador único
- user_id: Usuario propietario
- template_data: Template biométrico (BLOB)
- fingerprint_hash: Hash SHA-256
- device_type: Tipo de captor
- device_serial: Serial del dispositivo
- is_enabled: Activo/Inactivo
- last_verified_at: Última verificación
- verification_count: Total de verificaciones
- created_at: Fecha de creación
- updated_at: Última actualización
```

#### `biometric_login_attempts`
```sql
- id: Identificador único
- user_id: Usuario (puede ser NULL si falló antes de identificar)
- device_type: Tipo de captor
- attempt_type: 'capture', 'verify', 'failed_verify', 'success'
- status: 'success', 'failure', 'timeout', 'device_error'
- error_message: Mensaje de error (si aplica)
- ip_address: IP del cliente
- user_agent: Navegador/sistema
- created_at: Fecha del intento
```

### Endpoints de API

#### 1. Capturar huella
```
POST /api/biometric.php?action=capture-fingerprint
Authorization: Session requerida
Response:
{
  "ok": true,
  "message": "Listo para capturar huella digital...",
  "device_type": "morpho_e3_e4"
}
```

#### 2. Guardar huella
```
POST /api/biometric.php?action=save-fingerprint
Authorization: Session requerida
Body:
{
  "template_data": "base64_encoded_template",
  "fingerprint_hash": "sha256_hash",
  "device_serial": "device_serial_number"
}
Response:
{
  "ok": true,
  "message": "Huella digital guardada exitosamente."
}
```

#### 3. Verificar huella (Login)
```
POST /api/biometric.php?action=verify-fingerprint
Body:
{
  "email": "usuario@ejemplo.com",
  "fingerprint_hash": "sha256_hash"
}
Response:
{
  "ok": true,
  "message": "Autenticacion biometrica exitosa.",
  "user": {...}
}
```

#### 4. Verificar estado
```
GET /api/biometric.php?action=check-biometric-status
Authorization: Session requerida
Response:
{
  "ok": true,
  "enrolled": true,
  "enabled": true,
  "device_type": "morpho_e3_e4",
  "verification_count": 5,
  "last_verified": "2024-06-30 14:30:00"
}
```

#### 5. Desabilitar biometría
```
POST /api/biometric.php?action=disable-biometric
Authorization: Session requerida
Response:
{
  "ok": true,
  "message": "Autenticacion biometrica deshabilitada."
}
```

---

## FAQ (Preguntas Frecuentes)

**P: ¿Es segura la autenticación biométrica?**
R: Sí, usamos hashing SHA-256 y almacenamiento seguro en base de datos. Tus huellas nunca se guardan en texto plano.

**P: ¿Puedo registrar múltiples dedos?**
R: Actualmente solo se soporta un registro biométrico por usuario. Puedes reemplazarlo en cualquier momento.

**P: ¿Qué pasa si tengo un dedo lesionado?**
R: Puedes usar tu contraseña normal o registrar un dedo diferente reemplazando el anterior.

**P: ¿Funciona en dispositivos móviles?**
R: La autenticación biométrica requiere un captor Morpho E3/E4 conectado a USB. Los dispositivos móviles pueden usar autenticación por huella nativa del dispositivo en futuras versiones.

**P: ¿Puedo descargar mis datos biométricos?**
R: No es posible descargar los datos biométricos por razones de seguridad. Solo el dispositivo biométrico original puede capturar nuevos datos.

**P: ¿Qué sucede si pierdo acceso al captor?**
R: Siempre puedes iniciar sesión con tu correo y contraseña. Registra nuevamente con tu captor cuando lo recuperes.

---

## Soporte

Para problemas técnicos:
1. Verifica la [Sección de Solución de Problemas](#solución-de-problemas)
2. Reinicia tu navegador y dispositivo
3. Reinstala los drivers del captor
4. Contacta al administrador de sistemas

---

**Última actualización**: 30 de junio de 2024  
**Versión**: 1.0  
**Captor soportado**: Morpho E3, Morpho E4
