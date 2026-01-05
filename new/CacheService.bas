Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : CacheService
' Purpose   : Centraliza el refresco de las tablas de caché del sistema.
'---------------------------------------------------------------------------------------

Public Function InicializarCacheEdicion( _
                            ByVal p_IDEdicion As Long, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim db As DAO.Database
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    
    ' 1. Limpiar caché antigua de esta edición (si hubiera)
    db.Execute "DELETE FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & p_IDEdicion
    db.Execute "DELETE FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_IDEdicion
    
    ' 2. Insertar Meta inicial
    db.Execute "INSERT INTO TbCacheArbolRiesgosMeta (IDEdicion, ActiveBuildId, UpdatedAt) " & _
               "VALUES (" & p_IDEdicion & ", 1, Now())"
               
    ' Nota: La construcción de los nodos hijos se delega en la lógica de negocio 
    ' de refresco del árbol (legacy/cachearbol.bas adaptado).
    
    InicializarCacheEdicion = "OK"
    Exit Function

errores:
    p_Error = "Error en CacheService.InicializarCacheEdicion: " & Err.Description
End Function
