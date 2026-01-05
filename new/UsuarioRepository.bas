Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : UsuarioRepository
' Purpose   : Acceso a datos de usuarios y permisos.
'---------------------------------------------------------------------------------------

Public Function GetByUsuarioRed(ByVal p_UsuarioRed As String, Optional ByRef p_Error As String) As Usuario
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Usr As Usuario
    
    On Error GoTo errores
    Set db = DatabaseProvider.GetLanzaderaDB(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ' Nota: Ajustar nombre de tabla y campos según realidad de la Lanzadera
    Set rcd = db.OpenRecordset("SELECT * FROM TbUsuarios WHERE UsuarioRed='" & p_UsuarioRed & "'")
    If Not rcd.EOF Then
        Set m_Usr = New Usuario
        With m_Usr
            .ID = Nz(rcd!ID, 0)
            .CorreoUsuario = Nz(rcd!CorreoUsuario, "")
            .UsuarioRed = Nz(rcd!UsuarioRed, "")
            .Nombre = Nz(rcd!Nombre, "")
            .Activado = Nz(rcd!Activado, False)
            .FechaBaja = rcd!FechaBaja
        End With
        Set GetByUsuarioRed = m_Usr
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en UsuarioRepository.GetByUsuarioRed: " & Err.Description
End Function

Public Function EsAdministrador(ByVal p_UsuarioRed As String, ByVal p_IDAplicacion As Integer, Optional ByRef p_Error As String) As Boolean
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    On Error GoTo errores
    Set db = DatabaseProvider.GetLanzaderaDB(p_Error)
    
    ' Consulta a la tabla de aplicaciones de usuario de la Lanzadera
    Set rcd = db.OpenRecordset("SELECT EsAdministrador FROM TbUsuariosAplicaciones " & _
                               "WHERE Usuario='" & p_UsuarioRed & "' AND IDAplicacion=" & p_IDAplicacion)
    If Not rcd.EOF Then
        EsAdministrador = (Nz(rcd!EsAdministrador, "No") = "Sí")
    End If
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en UsuarioRepository.EsAdministrador: " & Err.Description
End Function

Public Function EsUsuarioCalidad(ByVal p_UsuarioRed As String, ByVal p_IDAplicacion As Integer, Optional ByRef p_Error As String) As Boolean
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    On Error GoTo errores
    Set db = DatabaseProvider.GetLanzaderaDB(p_Error)
    
    ' Consulta el rol en la tabla de permisos/roles de la Lanzadera
    Set rcd = db.OpenRecordset("SELECT IDRol FROM TbUsuariosAplicaciones " & _
                               "WHERE Usuario='" & p_UsuarioRed & "' AND IDAplicacion=" & p_IDAplicacion)
    If Not rcd.EOF Then
        ' Asumimos que IDRol = 2 es Calidad (basado en lógica legacy)
        EsUsuarioCalidad = (Nz(rcd!IDRol, 0) = 2)
    End If
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en UsuarioRepository.EsUsuarioCalidad: " & Err.Description
End Function
