# 📌 REFERENCIA RÁPIDA - Azure DevOps Integration

## 🎯 ¿Qué Necesito?

### Para Conectar:
- [ ] URL de Azure DevOps: `https://dev.azure.com/[TU_ORG]`
- [ ] Nombre de tu organización: `mi-org`
- [ ] Nombre de tu proyecto: `Mi Proyecto`
- [ ] Personal Access Token (PAT)

### Para Obtener el PAT:
```
1. dev.azure.com → Tu perfil → Personal access tokens
2. + New Token
3. Name: "Calculador-Horas"
4. Scopes: ✓ Work Items (Read)
5. Expiration: 90 días
6. Create → COPIAR
```

---

## 🚀 Cómo Usar (3 pasos)

### 1️⃣ Abre la App
```
Tu navegador → Calculador de Horas
```

### 2️⃣ Conecta Azure
```
Botón: "Conectar Azure DevOps"
├─ Organización: mi-org
├─ Proyecto: Mi Proyecto
└─ PAT: patsXXXXXXXXXXXX
```

### 3️⃣ ¡Listo!
```
Espera → Tareas cargan → Ver análisis
```

---

## 🔑 Dónde Poner Cada Dato

| Campo | Dónde Obtenerlo | Ejemplo |
|-------|-----------------|---------|
| **Organización** | URL de Azure DevOps | `mi-org` de `https://dev.azure.com/mi-org` |
| **Proyecto** | Selector proyectos (Azure) | `Mi Proyecto` |
| **PAT** | Personal Access Tokens (Azure) | `pat1234567890ABCDEF...` |

---

## 📊 Qué Trae del Servidor

```
Automáticamente obtiene de Azure DevOps:
✓ ID de tareas
✓ Títulos
✓ Tipos (Bug, Tarea, etc)
✓ Estados
✓ Asignados a
✓ Horas estimadas
✓ Horas completadas
✓ Horas restantes
✓ Etiquetas
✓ Áreas
✓ Sprints/Iteraciones
```

---

## 🔒 Seguridad en 3 Puntos

```
1. PAT en localStorage → Encriptado con base64
2. Comunicación directa → Solo con Azure DevOps API
3. Sin envíos externos → Nunca sale del navegador
```

---

## ⚠️ Si Algo Falla

### Error: "PAT inválido"
```
→ Copia PAT nuevamente desde Azure
→ Verifica sin espacios en blanco
```

### Error: "Acceso denegado"
```
→ Crea PAT con permiso "Work Items (Read)"
```

### Error: "Proyecto no encontrado"
```
→ Verifica nombre exacto del proyecto
→ Prueba copiar-pegar desde Azure DevOps
```

### No cargan tareas
```
→ Proyecto podría estar vacío
→ O sin tareas con información de horas
```

---

## 💾 Datos Guardados

Se guarda automáticamente en tu navegador:
```
- Organización
- Proyecto
- PAT (encriptado)

Se recupera al recargar la página.
```

---

## 🧪 Validar que Funciona

```
1. Botón "Conectar Azure DevOps" → ¿Aparece?
2. Haz clic → ¿Se abre modal?
3. Ingresa datos → ¿Se validan?
4. Conecta → ¿Cargan tareas?
5. Recarga página → ¿Datos siguen guardados?
```

---

## 📁 Archivos Nuevos

```
AZURE_SETUP.md         → Instrucciones detalladas
SETUP_COMPLETE.md      → Resumen instalación
CAMBIOS_REALIZADOS.md  → Qué se modificó (técnico)
GUIA_PRUEBAS.md        → Cómo testear
README_AZURE.md        → Resumen ejecutivo
REFERENCIA_RAPIDA.md   → Este archivo
```

---

## 🎯 Checklist Inicial

- [ ] Tengo mi organización de Azure DevOps
- [ ] Tengo acceso a un proyecto con tareas
- [ ] Obtuve mi PAT (no expirado)
- [ ] Abrí la aplicación en el navegador
- [ ] Vi el botón "Conectar Azure DevOps"
- [ ] Ingresé mis datos sin errores
- [ ] Se cargaron las tareas
- [ ] Veo análisis de horas

---

## 💡 Tips Útiles

```
✓ Guarda tu PAT en un gestor de contraseñas
✓ Usa PATs específicos por aplicación
✓ Revoca PATs que no uses
✓ El token tarda segundos en crearse
✓ La primera carga puede demorar más
✓ El PAT se guardará para futuros usos
```

---

## 🚀 Próximos Pasos (Futuros)

```
Después de conectar con Azure:

1. Análisis IA
   → Chat con Copilot sobre tus tareas
   
2. Actualizar en Azure
   → Cambiar estado, agregar comentarios
   
3. Predicciones
   → Velocidad del equipo
   → Estimaciones automáticas
   
4. Reportes
   → Power BI
   → Teams Alerts
```

---

## 📞 Documentación Completa

**Leer en este orden:**

1. **REFERENCIA_RAPIDA.md** ← Estás aquí (visión general)
2. **AZURE_SETUP.md** ← Paso a paso completo
3. **GUIA_PRUEBAS.md** ← Cómo validar
4. **CAMBIOS_REALIZADOS.md** ← Detalles técnicos
5. **README_AZURE.md** ← Resumen ejecutivo

---

## 🎉 ¡Listo!

Solo necesitas:
1. Tu PAT
2. 30 segundos
3. ¡Conectar!

**Luego:** Todas tus tareas analizadas automáticamente.

---

**Fecha de implementación:** Mayo 3, 2026
**Estado:** ✅ LISTO PARA USAR
**Versión:** 1.0 - Azure DevOps Integration
