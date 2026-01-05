Option Compare Database
Option Explicit

Public Function CacheArbolRiesgosTx_ActualizarRiesgo( _
                                                    p_Riesgo As riesgo, _
                                                    Optional ByVal p_db As DAO.Database, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean

    On Error GoTo errores

    If p_Riesgo Is Nothing Then Exit Function

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If

    Set wksLocal = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    CacheArbolRiesgos_ActualizarRiesgo p_Riesgo, db, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If

    CacheArbolRiesgosTx_ActualizarRiesgo = "OK"
    Exit Function

errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_ActualizarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_BorrarRiesgo( _
                                                p_Riesgo As riesgo, _
                                                Optional ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean

    On Error GoTo errores

    If p_Riesgo Is Nothing Then Exit Function

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If

    Set wksLocal = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    CacheArbolRiesgos_BorrarRiesgo p_Riesgo, db, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If

    CacheArbolRiesgosTx_BorrarRiesgo = "OK"
    Exit Function

errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_BorrarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_RebuildEdicion( _
                                                    p_Edicion As Edicion, _
                                                    Optional ByVal p_db As DAO.Database, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    On Error GoTo errores

    CacheArbolRiesgosTx_RebuildEdicion = CacheArbolRiesgos_RebuildEdicion(p_Edicion, p_db, p_Error)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_RebuildEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_ActualizarOrdenRiesgos( _
                                                            p_Edicion As Edicion, _
                                                            p_ColPriorizacion As Scripting.Dictionary, _
                                                            Optional ByVal p_db As DAO.Database, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    On Error GoTo errores

    CacheArbolRiesgosTx_ActualizarOrdenRiesgos = CacheArbolRiesgos_ActualizarOrdenRiesgos(p_Edicion, p_ColPriorizacion, p_db, p_Error)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_ActualizarOrdenRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_ActualizarEdicion( _
                                                    p_Edicion As Edicion, _
                                                    Optional ByVal p_db As DAO.Database, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean

    On Error GoTo errores

    If p_Edicion Is Nothing Then Exit Function

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If

    Set wksLocal = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    CacheArbolRiesgos_ActualizarEdicion p_Edicion, db, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If

    CacheArbolRiesgosTx_ActualizarEdicion = "OK"
    Exit Function

errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_ActualizarEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgo( _
                                                                p_Riesgo As riesgo, _
                                                                p_SQL As String, _
                                                                Optional ByVal p_db As DAO.Database, _
                                                                Optional p_SQLExtra As Variant, _
                                                                Optional ByRef p_Error As String _
                                                                ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean

    On Error GoTo errores

    If p_Riesgo Is Nothing Then Exit Function
    If Trim$(p_SQL) = "" Then Exit Function

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If

    Set wksLocal = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    CacheArbolRiesgosTx_EjecutarSqls db, p_SQL, p_SQLExtra

    CacheArbolRiesgos_ActualizarRiesgo p_Riesgo, db, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If

    CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgo = "OK"
    Exit Function

errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgoId( _
                                                                    p_IDRiesgo As Long, _
                                                                    p_SQL As String, _
                                                                    Optional ByVal p_db As DAO.Database, _
                                                                    Optional p_SQLExtra As Variant, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean
    Dim m_Riesgo As riesgo

    On Error GoTo errores

    If p_IDRiesgo <= 0 Then Exit Function
    If Trim$(p_SQL) = "" Then Exit Function

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If

    Set wksLocal = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    CacheArbolRiesgosTx_EjecutarSqls db, p_SQL, p_SQLExtra

    Set m_Riesgo = Constructor.getRiesgo(CStr(p_IDRiesgo), , , p_Error)
    If p_Error <> "" Then Err.Raise 1000

    CacheArbolRiesgos_ActualizarRiesgo m_Riesgo, db, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If

    CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgoId = "OK"
    Exit Function

errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgosTx_EjecutarSqlYActualizarRiesgoId ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub CacheArbolRiesgosTx_EjecutarSqls( _
                                            p_db As DAO.Database, _
                                            p_SQL As String, _
                                            Optional p_SQLExtra As Variant _
                                            )
    Dim m_Idx As Long

    If Trim$(p_SQL) <> "" Then
        p_db.Execute p_SQL
    End If

    If IsMissing(p_SQLExtra) Then
        Exit Sub
    End If

    If IsArray(p_SQLExtra) Then
        For m_Idx = LBound(p_SQLExtra) To UBound(p_SQLExtra)
            If Trim$(CStr(p_SQLExtra(m_Idx))) <> "" Then
                p_db.Execute CStr(p_SQLExtra(m_Idx))
            End If
        Next
    Else
        If Trim$(CStr(p_SQLExtra)) <> "" Then
            p_db.Execute CStr(p_SQLExtra)
        End If
    End If
End Sub
