CHANGELOG_INTEGRATION

Modulo / Archivo | Metodo / Funcion | Cambio Realizado | Descripcion del Ajuste
--- | --- | --- | ---
cachearbol.bas | CacheArbolRiesgos_UpsertNodo | Nueva funcion | Upsert idempotente para nodos del cache con Update + Insert si no existe.
cachearbol.bas | CacheArbolRiesgos_ActualizarRiesgo | Nueva funcion | Actualiza el nodo de riesgo y reconstruye su rama (planes y acciones) dentro del build activo.
cachearbol.bas | CacheArbolRiesgos_ActualizarPlan | Nueva funcion | Actualiza el nodo de plan y sus acciones, sincronizando tambien el nodo de riesgo.
cachearbol.bas | CacheArbolRiesgos_BorrarPlan | Nueva funcion | Elimina un plan y sus acciones del cache por NodeKey/ParentKey.
cachearbol.bas | CacheArbolRiesgos_BorrarRiesgo | Nueva funcion | Elimina toda la rama del riesgo del cache por IDRiesgo.
Riesgo.cls | Borrar | Insercion de codigo | Llamada a CacheArbolRiesgos_BorrarRiesgo tras eliminar en TbRiesgos.
Riesgo.cls | Registrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarRiesgo tras registrar actualizacion de edicion.
Riesgo.cls | RetiroRegistrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarRiesgo tras registrar actualizacion de edicion.
Riesgo.cls | RetiroRegistrarQuitar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarRiesgo tras registrar actualizacion de edicion.
Riesgo.cls | MaterializacionRegistrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarRiesgo tras registrar actualizacion de edicion.
Riesgo.cls | MaterializacionQuitarRegistrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarRiesgo tras registrar actualizacion de edicion.
PM.cls | Registrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarPlan tras registrar actualizacion de edicion.
PM.cls | Borrar | Insercion de codigo | Llamada a CacheArbolRiesgos_BorrarPlan tras registrar actualizacion de edicion.
PC.cls | Registrar | Insercion de codigo | Llamada a CacheArbolRiesgos_ActualizarPlan tras registrar actualizacion de edicion.
PC.cls | Borrar | Insercion de codigo | Llamada a CacheArbolRiesgos_BorrarPlan tras registrar actualizacion de edicion.
PMAccion.cls | Registrar/Borrar | Modificacion | Usa CacheArbolRiesgos_ActualizarPlan para recalcular el plan completo para recalcular el plan completo.
PCAccion.cls | Registrar/Borrar | Modificacion | Usa CacheArbolRiesgos_ActualizarPlan para recalcular el plan completo para recalcular el plan completo.
cachearbol.bas | CacheArbolRiesgos_ActualizarAccion | Eliminacion | Se elimina la funcion al unificar actualizacion de acciones via CacheArbolRiesgos_ActualizarPlan.
cachearbol.bas | CacheArbolRiesgos_ActualizarOrdenRiesgos | Nueva funcion | Actualiza SortIndex en cache para riesgos segun priorizacion guardada.
Form_FormRiesgosEstablecerPrioridades.cls | EstablecerPriorizaciones | Insercion de codigo | Actualiza el orden de riesgos en cache tras grabar la priorizacion.
cachearbol.bas | CacheArbolRiesgos_RebuildEdicion | Modificacion | Ordena riesgos por Priorizacion y luego IDRiesgo para el SortIndex en cache.
cachearbol.bas | CacheArbolRiesgos_ActualizarRiesgo/ActualizarPlan/ActualizarOrdenRiesgos | Modificacion | Ajusta el SortIndex siguiendo Priorizacion y luego IDRiesgo.\r\ncachearbol.bas | CacheArbolRiesgos_BorrarEdicionCache | Nueva funcion | Borra toda la cache de una edicion (nodos y meta) para forzar reconstruccion limpia.\r\nForm_FormRiesgosGestion.cls | ComandoActualizarContador_Click | Modificacion | Limpia la cache de la edicion antes de reconstruir el arbol para evitar colision de keys.\r\n