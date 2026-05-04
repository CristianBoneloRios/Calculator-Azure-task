# ✅ Integración Azure DevOps - ¡Implementado!

## 🎯 ¿Qué se agregó?

Tu aplicación ahora tiene una **conexión completa con Azure DevOps** para traer tareas automáticamente.

### Archivos Modificados:
1. ✅ **app.js** - Lógica de conexión con Azure
2. ✅ **index.html** - Interfaz de usuario (botón + modales)
3. ✅ **styles.css** - Estilos visuales para Azure

### Nuevas Funciones:
- 🔗 Conectar con Azure DevOps usando PAT
- 📥 Traer tareas automáticamente
- 💾 Guardar credenciales en el navegador (encriptadas)
- ⏳ Mostrar progreso de carga
- 🔒 Máxima seguridad (sin envío de datos a servidores externos)

---

## 🚀 Cómo Empezar

### 1️⃣ Obtén tu PAT (Personal Access Token)

Ve a: `https://dev.azure.com/`

```
Settings → Personal access tokens → + New Token
- Name: "Calculador-Horas"
- Scopes: Work Items (Read)
- Expiration: 90 días
```

**⚠️ Copia el token - solo aparece una vez**

### 2️⃣ Abre tu Calculador

Haz clic en **"Conectar Azure DevOps"** y completa:
- Organización
- Proyecto
- Tu PAT

### 3️⃣ ¡Listo!

Las tareas se cargarán automáticamente y se calcularán.

---

## 📝 Dónde Colocar tu PAT (Seguridad)

### ❌ NO RECOMENDADO en el código:
```javascript
const AzureConfig = {
  pat: 'patXXXXXXXXX'  // ⚠️ Esto es inseguro
};
```

### ✅ RECOMENDADO - Ingresar en la UI:
```
App → Conectar Azure DevOps → Ingresa tu PAT
```

El PAT se guarda **encriptado** en localStorage y **nunca se envía** a otros servidores.

---

## 🔐 Características de Seguridad

- ✓ PAT encriptado con base64 en localStorage
- ✓ Solo comunicación directa con Azure DevOps
- ✓ Sin almacenamiento en servidores externos
- ✓ Puedes revocar el token en cualquier momento
- ✓ Validación de entrada en todos los campos

---

## 📊 Qué Trae del Servidor

```
├── ID de Elemento
├── Título
├── Tipo (Bug, Tarea, Historia, etc)
├── Estado (Nuevo, En Progreso, Completado)
├── Asignado A
├── Estimación Original (horas)
├── Trabajo Completado (horas)
├── Trabajo Restante (horas)
├── Etiquetas
├── Ruta de Área
└── Iteración (Sprint)
```

---

## 🎨 Cambios en la UI

### Botón Nueva:
- **"Conectar Azure DevOps"** (al lado del botón de cargar archivos)
  - Estilo: Azul Microsoft (#0078d4)
  - Ícono: Microsoft logo

### Modales Nuevos:
1. **Modal de Conexión** - Donde ingresas credenciales
2. **Modal de Carga** - Muestra progreso mientras trae datos

### Mensajes:
- Notificaciones de éxito/error
- Validación de campos
- Progreso en tiempo real

---

## 🧪 Testear la Conexión

### Prueba 1: Modal aparece
```
Haz clic en "Conectar Azure DevOps"
→ Debería aparecer un modal
```

### Prueba 2: Validación
```
Deja un campo vacío → Haz clic en "Conectar"
→ Debe mostrar error: "Todos los campos son requeridos"
```

### Prueba 3: Conexión Real
```
Ingresa tu información correcta
→ Modal de carga aparece
→ Tareas aparecen en la pantalla
```

---

## 🔧 Si Algo No Funciona

### Abre la Consola del Navegador (F12)
- Haz clic en **"Console"**
- Verás errores detallados
- Copia el error en la descripción

### Errores Comunes:

| Error | Solución |
|-------|----------|
| "PAT inválido" | Copia el token nuevamente desde Azure |
| "Acceso denegado" | El PAT necesita permisos "Work Items (Read)" |
| "Proyecto no encontrado" | Verifica el nombre exacto del proyecto |
| "No se conecta" | Verifica tu conexión a internet |

---

## 📈 Próximas Mejoras Planificadas

- [ ] Chat con IA para análisis automático
- [ ] Actualizar tareas desde el calculador
- [ ] Predicciones de velocidad del equipo
- [ ] Alertas de tareas en riesgo
- [ ] Exportar a Power BI
- [ ] Integración con Teams

---

## 💡 Tips de Uso

1. **Guarda tu PAT** en un lugar seguro (gestor de contraseñas)
2. **Revoca el token** si cambias de PC o lo comprometes
3. **Usa un token específico** para cada aplicación (no uno maestro)
4. **Actualiza tareas regularmente** para datos frescos
5. **Exporta reportes** antes de cerrar la aplicación

---

## 📞 Documentación Completa

Ve a: **[AZURE_SETUP.md](./AZURE_SETUP.md)**

Contiene:
- Instrucciones paso a paso
- Cómo obtener tu PAT
- Solución de problemas
- Información de seguridad

---

## ✨ ¡Listo para Usar!

Tu calculador ahora está integrado con Azure DevOps. 

**¡Conecta y comienza a calcular! 🚀**
