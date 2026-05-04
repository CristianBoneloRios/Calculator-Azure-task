# 🧪 Guía de Prueba - Azure DevOps Integration

## ✅ Verificación Rápida

Antes de usar con datos reales, verifica que todo funciona correctamente.

---

## 📋 Checklist de Pruebas

### Prueba 1: Interfaz de Usuario
- [ ] Abre la aplicación en el navegador
- [ ] En la sección "Cargar Archivos" deberías ver:
  - [ ] Botón "Seleccionar Archivos" (azul)
  - [ ] Texto "o" en el medio
  - [ ] Botón "Conectar Azure DevOps" (azul Microsoft)

### Prueba 2: Modal de Conexión
- [ ] Haz clic en "Conectar Azure DevOps"
- [ ] Debe abrirse un modal con:
  - [ ] Ícono de Microsoft
  - [ ] Título "Conectar Azure DevOps"
  - [ ] 3 campos de entrada:
    - [ ] Organización
    - [ ] Nombre del Proyecto
    - [ ] Personal Access Token
  - [ ] Botones: "Cancelar" y "Conectar"

### Prueba 3: Validación de Campos
**Prueba vacío:**
- [ ] Deja todos los campos en blanco
- [ ] Haz clic en "Conectar"
- [ ] Debe mostrar: ⚠️ "Todos los campos son requeridos"

**Prueba org inválida:**
- [ ] Completa solo Organización con: `!!!`
- [ ] Haz clic en "Conectar"
- [ ] Debe mostrar: ⚠️ "Nombre de organización inválido"

**Prueba PAT con espacios:**
- [ ] Completa con: `pat XXXX (con espacio)`
- [ ] Haz clic en "Conectar"
- [ ] Debe mostrar: ⚠️ "El PAT tiene espacios"

### Prueba 4: Cancelar
- [ ] Abre el modal de conexión
- [ ] Haz clic en "Cancelar"
- [ ] El modal debe cerrarse sin hacer nada

### Prueba 5: Recuperación de Datos Guardados
- [ ] Ingresa datos válidos y conecta
- [ ] Recarga la página (F5)
- [ ] Abre el modal nuevamente
- [ ] Los campos Organización y Proyecto deben estar pre-llenados
- [ ] *El PAT NO se muestra (solo se guarda)*

---

## 🔑 Pruebas con Datos Reales

Si tienes una cuenta de Azure DevOps y quieres probar completamente:

### Paso 1: Obtén tu PAT
1. Ve a: `https://dev.azure.com/[tu-org]`
2. Perfil (arriba derecha) → Personal access tokens
3. + New Token
   - Name: `Calculador-Test`
   - Scopes: **Work Items (Read)** ✓
   - Expiration: 7 días
4. Copia el token

### Paso 2: Busca tu Organización
- La encuentras en la URL: `https://dev.azure.com/**mi-org**`

### Paso 3: Busca un Proyecto
- Lo encuentras en el selector de proyectos (Azure DevOps)

### Paso 4: Prueba la Conexión
- Haz clic en "Conectar Azure DevOps"
- Ingresa:
  - **Organización:** `mi-org`
  - **Proyecto:** `Mi Proyecto`
  - **PAT:** Tu token copiado
- Haz clic en "Conectar"

### Resultado Esperado:
- [ ] Modal de carga aparece
- [ ] Muestra: "Obteniendo tareas de Azure DevOps..."
- [ ] Después de unos segundos (depende del proyecto):
  - [ ] Modal se cierra
  - [ ] Aparece notificación verde: ✓ "XX tareas cargadas desde Azure DevOps"
  - [ ] Las tareas aparecen en la pantalla
  - [ ] Estadísticas se actualizan

---

## ⚠️ Errores Comunes en Pruebas

### Error: "PAT inválido o expirado"
**Causas:**
- PAT expirado
- PAT copiado incorrectamente
- Espacios en blanco

**Solución:**
- Copia el PAT nuevamente desde Azure
- Verifica que no tenga espacios

### Error: "Acceso denegado"
**Causas:**
- PAT sin permisos de "Work Items (Read)"
- PAT creado sin los scopes correctos

**Solución:**
- Crea un nuevo PAT con permiso "Work Items (Read)"

### Error: "Proyecto no encontrado"
**Causas:**
- Nombre de proyecto incorrecto
- No tienes acceso a ese proyecto

**Solución:**
- Verifica el nombre exacto en Azure DevOps
- Copia y pega desde Azure DevOps directamente

### No aparecen tareas
**Causas:**
- El proyecto no tiene tareas
- Las tareas no tienen información de horas

**Solución:**
- Prueba con otro proyecto que tenga tareas
- Verifica que haya tareas con estimación de horas

---

## 🔍 Debugging

Si algo no funciona, abre la consola del navegador:

### En Chrome/Edge:
1. Presiona `F12`
2. Haz clic en pestaña **"Console"**
3. Intenta conectar nuevamente
4. Verás mensajes de error detallados

### Mensajes útiles:
```javascript
// Si ves esto → Conexión exitosa
"✓ XX tareas cargadas desde Azure DevOps"

// Si ves esto → Error de autenticación
"Error: PAT inválido o expirado"

// Si ves esto → Problema con la API
"Error fetching Azure data: ..."
```

---

## 📊 Datos que Debe Traer

Si la conexión es exitosa, deberías ver:

| Campo | Ejemplo |
|-------|---------|
| ID | 12345 |
| Título | Implementar login |
| Tipo | Tarea |
| Estado | Activo |
| Asignado a | Juan Pérez |
| Estimación Original | 8 horas |
| Trabajo Completado | 5 horas |
| Trabajo Restante | 3 horas |

---

## ✨ Checklist Final

- [ ] Botón "Conectar Azure DevOps" aparece
- [ ] Modal se abre correctamente
- [ ] Validaciones funcionan
- [ ] Cancelar cierra el modal
- [ ] Los datos se guardan (recuperación tras recarga)
- [ ] Conexión con datos válidos funciona
- [ ] Tareas aparecen en la pantalla
- [ ] Estadísticas se actualizan

---

## 🚀 ¡Todo Listo!

Si todas las pruebas pasaron, tu integración con Azure DevOps está completa.

**Nota:** La primera conexión puede tomar unos segundos dependiendo de:
- Cantidad de tareas en el proyecto
- Conexión a internet
- API de Azure DevOps

---

## 📞 Si Necesitas Ayuda

1. Revisa [AZURE_SETUP.md](./AZURE_SETUP.md) para pasos detallados
2. Consulta [CAMBIOS_REALIZADOS.md](./CAMBIOS_REALIZADOS.md) para ver qué se modificó
3. Abre la consola (F12) para ver errores específicos
4. Verifica tu conexión a internet

---

**¡Listo para empezar! 🎉**
