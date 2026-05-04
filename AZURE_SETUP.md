# 🚀 Configuración de Azure DevOps - Calculador de Horas

## ¿Cómo funciona?

Tu calculador ahora puede traer tareas **directamente desde Azure DevOps** sin necesidad de descargar CSV manualmente. Solo necesitas:

1. **Organización de Azure DevOps**
2. **Nombre del Proyecto**
3. **Personal Access Token (PAT)** - Tu contraseña segura

---

## 📋 PASO 1: Obtener tu Personal Access Token (PAT)

### Opción A: Desde Azure DevOps Web

1. Ve a: `https://dev.azure.com/[tu-organizacion]`
2. Haz clic en **tu perfil** (arriba a la derecha) → **Personal access tokens**
3. Haz clic en **+ New Token**
4. Completa los datos:
   - **Name:** `Calculador-Horas` (o el nombre que quieras)
   - **Organization:** Selecciona tu organización
   - **Expiration:** Elige cuánto tiempo quieres que sea válido (90 días es seguro)
   - **Scopes:** Selecciona **Work Items (Read)** ✓
5. Haz clic en **Create**
6. **COPIA el token** (aparecerá una sola vez - guárdalo en un lugar seguro)

### Aspecto del Token
```
patxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

---

## 📍 PASO 2: Usar el Calculador

### En la aplicación:

1. Abre tu calculador en el navegador
2. Haz clic en el botón **"Conectar Azure DevOps"** (en la sección de upload)
3. Completa el formulario:
   - **Organización:** Tu nombre de organización
     - Ejemplo: `mi-organizacion` 
     - O la URL completa: `https://dev.azure.com/mi-organizacion`
   - **Nombre del Proyecto:** El proyecto exacto de donde traer tareas
     - Ejemplo: `Mi Proyecto`
   - **Personal Access Token:** El PAT que copiaste en el Paso 1
4. Haz clic en **"Conectar"**
5. ¡Listo! Las tareas se cargarán automáticamente

---

## 🔒 Seguridad

✅ **Tu PAT es 100% seguro:**
- Se guarda **solo en tu navegador** (localStorage)
- Se **encripta con base64** (protección básica)
- **NUNCA se envía** a servidores externos
- Solo se usa para conectar directamente con Azure DevOps
- Puedes **revocar el token en cualquier momento** desde Azure DevOps

---

## 📊 ¿Qué datos trae?

El calculador automáticamente obtiene:

| Campo | Descripción |
|-------|-------------|
| **ID** | Identificador del elemento |
| **Título** | Nombre de la tarea |
| **Tipo** | Bug, Tarea, Historia, etc. |
| **Estado** | Nuevo, En Progreso, Completado |
| **Asignado a** | Persona responsable |
| **Estimación Original** | Horas estimadas |
| **Trabajo Completado** | Horas trabajadas |
| **Trabajo Restante** | Horas pendientes |
| **Etiquetas** | Tags del elemento |
| **Ruta de Área** | Categoría/Área |
| **Iteración** | Sprint/Ciclo |

---

## ⚙️ Cómo cambiar de proyecto

1. Haz clic nuevamente en **"Conectar Azure DevOps"**
2. Cambia cualquiera de los datos (Org, Proyecto, PAT)
3. Haz clic en **"Conectar"**
4. Las nuevas tareas se cargarán

---

## ❌ Solucionar Problemas

### Error: "PAT inválido o expirado"
- ✓ Copia el token nuevamente desde Azure DevOps
- ✓ Verifica que no haya espacios en blanco
- ✓ Revisa que el token siga siendo válido (no expirado)

### Error: "Acceso denegado"
- ✓ El token necesita permiso **"Work Items (Read)"**
- ✓ Crea un token nuevo con los permisos correctos

### Error: "Proyecto no encontrado"
- ✓ Verifica el nombre exacto del proyecto (incluyendo mayúsculas/minúsculas)
- ✓ Asegúrate que tienes acceso a ese proyecto

### No aparecen tareas
- ✓ El proyecto podría estar vacío
- ✓ Verifica que tengas tareas con información de horas

---

## 🔑 Dónde encontrar tu información de Azure

### Organización
URL de Azure DevOps: `https://dev.azure.com/**mi-organizacion**`

### Proyecto
Aparece en el selector de proyectos (izquierda superior en Azure DevOps)

### Personal Access Token
Settings → Personal access tokens → + New Token

---

## 💾 Próximas Actualizaciones Planificadas

- ✨ Actualizar tareas directamente en Azure DevOps desde el calculador
- ✨ Chat con IA para análisis inteligente de horas
- ✨ Predicciones de velocidad del equipo
- ✨ Alertas de tareas en riesgo
- ✨ Exportar reportes a Power BI

---

## 📞 ¿Dudas?

Si tienes problemas:
1. Revisa que tu PAT sea válido
2. Verifica los permisos en Azure DevOps
3. Comprueba que el nombre del proyecto es exacto
4. Abre la consola del navegador (F12) para más detalles

**¡Listo para usar!** 🚀
