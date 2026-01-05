# Plan Maestro de Refactorización Total: Arquitectura Limpia (Domain-Driven)

## 1. Visión General
Transformar el sistema actual (basado en Active Record y acoplamiento fuerte) en una arquitectura de capas desacoplada. El objetivo es que **ninguna Clase (.cls)** tenga código SQL y que **ningún Formulario** gestione transacciones o cachés directamente.

---

## 2. Definición de Capas (Estándar de Implementación)

### Capa 1: Dominio (Entidades POCO)
- **Clases:** `Proyecto`, `Edicion`, `Riesgo`, `RiesgoExterno`, `PM` (Mitigación), `PC` (Contingencia), `PMAccion`, `PCAccion`, `Suministrador`, `Usuario`, `Anexo`.
- **Estado:** Solo variables públicas y Properties Get calculadas.
- **Regla:** Sin `DAO`, sin `Recordsets`, sin `getdb()`.

### Capa 2: Repositorios (Persistencia CRUD)
- **Módulos:** `ProyectoRepository`, `RiesgoRepository`, `EdicionRepository`, `PlanRepository`, `SuministradorRepository`, `AnexoRepository`.
- **Responsabilidad:** SQL puro (`INSERT`, `UPDATE`, `DELETE`, `SELECT`). 
- **Firma:** `Function Save(p_Objeto, Optional p_db)`. Si `p_db` está presente, se usa para la transacción.

### Capa 3: Servicios (Orquestación de Casos de Uso)
- **Módulos:** `ProyectoService`, `RiesgoService`, `PlanService`, `CacheService`, `NotificationService`.
- **Responsabilidad:** Gestión de transacciones (`ws.BeginTrans`), validación cruzada y refresco de cachés.

### Capa 4: Factorías (Instanciación)
- **Módulo:** `EntityFactory`.
- **Responsabilidad:** Crear objetos con sus valores por defecto (ej: un nuevo Riesgo con su colección de Planes inicializada).

---

## 3. Catálogo Total de Casos de Uso por Módulo

### Módulo A: Ciclo de Vida del Proyecto y Ediciones
- **CU-A1: Alta de Proyecto:** (Service -> ProyectoService). Crea Proyecto + Edición 1 + Tareas iniciales + Registro en Último Proyecto.
- **CU-A2: Cierre de Edición y Apertura de Siguiente:** (Service -> ProyectoService). 
    1. Transacción única. 
    2. Setea `FechaPublicacion` en Edición N. 
    3. Crea Edición N+1. 
    4. Copia todos los Riesgos (No Retirados/No Cerrados) de N a N+1. 
    5. Copia Planes y Acciones asociados.
    6. Reconstruye `CacheArbol` completo.
- **CU-A3: Actualización de Fechas Máximas:** Recalcular `FechaMaxProximaPublicacion` tras cualquier cambio en riesgos.

### Módulo B: Gestión Integral de Riesgos
- **CU-B1: Upsert de Riesgo (Alta/Edición):** (Service -> RiesgoService). Guarda datos de riesgo + Actualiza `CachePublicabilidad` + Refresca Nodo en Árbol.
- **CU-B2: Flujo de Aprobación (Calidad):** `AprobarAceptacion`, `RechazarAceptacion`, `AprobarRetiro`, `RechazarRetiro`. 
    - Orquestación: Cambia campos de validación + Dispara Correo a Técnico + Actualiza Estado del Riesgo + Sincroniza Árbol.
- **CU-B3: Retipificación:** Cambio de código de biblioteca. Valida el nuevo código y limpia flags de "Pendiente Retipificar".
- **CU-B4: Priorización Masiva:** (Service -> RiesgoService). Recibe lista de IDs, reasigna `Priorizacion` en un solo bloque SQL y actualiza el orden en la caché del árbol.

### Módulo C: Riesgos Externos y Promoción
- **CU-C1: Integración de Riesgo de Oferta/Suministrador/Pedido:** 
    - Si `Trasladar = "Sí"`: Crea Riesgo + Vincula IDs + Envía Aviso.
    - Si `Trasladar = "No"`: Borra Riesgo asociado (si existe) + Desvincula.
    - Todo en transacción única mediante `RiesgoService.ProcesarRiesgoExterno`.

### Módulo D: Planes de Acción (Mitigación y Contingencia)
- **CU-D1: Gestión de Acciones:** Alta/Baja/Cierre de acciones.
    - Orquestación: Si se cierra la última acción, el Plan debe cambiar de estado automáticamente. Si el riesgo era "Alto", validar que tenga planes activos.
- **CU-D2: Reversión de Acciones:** Gestión de `PMAccionReversa` / `PCAccionReversa`.

### Módulo E: Suministradores y Entidades
- **CU-E1: Asociación de Suministrador a Proyecto:** Crear vínculo en `TbProyectoSuministradores` y activar/desactivar flag de Gestión de Calidad.

### Módulo F: Gestión de Anexos
- **CU-F1: Persistencia de Archivos:** (Service -> AnexoService). Copia física de archivo a servidor + Registro en `TbAnexos` + Vinculación a Riesgo/Edición/Proyecto.

---

### Regla de Oro: Codificación y Tipos
- **Prohibido el uso de acentos en el código:** Ningún Enum, variable o nombre de función debe contener acentos (ej: NO usar `EnumSiNo.Sí`).
- **Uso de Booleanos:** Para flags de Si/No, usar el tipo nativo `Boolean` (`True`/`False`).
- **Enums de Estado:** Si se requiere un tercer estado (ej: Indefinido), usar un Enum con nombres en inglés o sin acentos (ej: `tsYes`, `tsNo`, `tsUndefined`).
- **Firma Obligatoria:** TODAS las funciones en las capas de Repositorio, Servicio y Factoría DEBEN terminar con el parámetro: `Optional ByRef p_Error As String`.

### Inyección de Dependencias Manual
- Al mover lógica a un Repositorio, asegúrate de que cada función reciba el objeto de dominio completo, no solo IDs.
- Los Repositorios deben aceptar `Optional ByVal p_db As DAO.Database` para participar en transacciones orquestadas por el Servicio.

### Gestión de Transacciones
- El **Servicio** es el responsable único de abrir y cerrar transacciones (`ws.BeginTrans`).
- Siempre usar un flag `blnEnTransaccion` para ejecutar `ws.Rollback` en el bloque `errores:` si algo falla.

### Eliminación de Redundancias
- Al usar el Repositorio, elimina todos los `db.Execute "UPDATE..."` que se hacían después de un `.Update` de recordset. El Repositorio debe dejar el objeto persistido correctamente en una sola operación.
- La UI nunca debe llamar a la caché directamente después de un cambio de datos; es el **Servicio** quien dispara la actualización de la caché tras el Commit.


---

## 5. Hoja de Ruta de Implementación para la IA

1. **Paso 1: Repositorios Puros.** Crear los módulos `.bas` de Repository y mover el SQL de las clases.
2. **Paso 2: Servicios de Orquestación.** Crear los módulos `.bas` de Service y unificar los métodos `Registrar...Transaccional`.
3. **Paso 3: Desacoplo de UI.** Modificar formularios para que solo llamen a una línea del `Service`.
4. **Paso 4: Limpieza de Clases.** Eliminar todo rastro de DAO y variables temporales de las clases `.cls`.
5. **Paso 5: Validación Final.** Asegurar que `RefrescarDerivadosTx` sea llamado siempre por el Service tras un Commit.