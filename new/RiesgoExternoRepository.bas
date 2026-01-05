Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : RiesgoExternoRepository
' Purpose   : Capa de persistencia para la tabla TbRiesgosAIntegrar.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_IDRiesgoExt As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As RiesgoExterno
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_RE As RiesgoExterno
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgosAIntegrar WHERE IDRiesgoExt=" & p_IDRiesgoExt)
    If Not rcd.EOF Then
        Set m_RE = New RiesgoExterno
        With m_RE
            .IDRiesgoExt = Nz(rcd!IDRiesgoExt, 0)
            .CodRiesgo = Nz(rcd!CodRiesgo, "")
            .IDRiesgo = Nz(rcd!IDRiesgo, 0)
            .Origen = Nz(rcd!Origen, "")
            .IDEdicion = Nz(rcd!IDEdicion, 0)
            .Descripcion = Nz(rcd!Descripcion, "")
            .CausaRaiz = Nz(rcd!CausaRaiz, "")
            .FechaDetectado = rcd!FechaDetectado
            .FechaAltaRegistro = rcd!FechaAltaRegistro
            .UsuarioRegistra = Nz(rcd!UsuarioRegistra, "")
            .MotivoNoIntegrado = Nz(rcd!MotivoNoIntegrado, "")
            .FechaMotivo = rcd!FechaMotivo
            .Trasladar = Nz(rcd!Trasladar, "")
            .Suministrador = Nz(rcd!Suministrador, "")
            .Pedido = Nz(rcd!Pedido, "")
            .ProveedorPedido = Nz(rcd!ProveedorPedido, "")
            .CausaRiesgoPedido = Nz(rcd!CausaRiesgoPedido, "")
            .RequiereRiesgoDeBiblioteca = Nz(rcd!RequiereRiesgoDeBiblioteca, "")
            .CodRiesgoBiblioteca = Nz(rcd!CodRiesgoBiblioteca, "")
            .RiesgoPendienteRetipificacion = Nz(rcd!RiesgoPendienteRetipificacion, "")
        End With
        Set GetById = m_RE
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en RiesgoExternoRepository.GetById: " & Err.Description
End Function

Public Function Save(p_RE As RiesgoExterno, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    If p_RE.IDRiesgoExt = 0 Then
        Set rcd = db.OpenRecordset("TbRiesgosAIntegrar")
        rcd.AddNew
        ' Aquí la IDRiesgoExt debería ser generada por una Factory o calculada antes del Save
        rcd!FechaAltaRegistro = Now()
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgosAIntegrar WHERE IDRiesgoExt=" & p_RE.IDRiesgoExt)
        If rcd.EOF Then
            p_Error = "No se encontró el riesgo externo para actualizar"
            GoTo errores
        End If
        rcd.Edit
    End If
    
    With rcd
        If p_RE.IDRiesgoExt <> 0 Then !IDRiesgoExt = p_RE.IDRiesgoExt
        !CodRiesgo = p_RE.CodRiesgo
        !IDRiesgo = IIf(p_RE.IDRiesgo = 0, Null, p_RE.IDRiesgo)
        !Origen = p_RE.Origen
        !IDEdicion = IIf(p_RE.IDEdicion = 0, Null, p_RE.IDEdicion)
        !Descripcion = p_RE.Descripcion
        !CausaRaiz = p_RE.CausaRaiz
        !UsuarioRegistra = p_RE.UsuarioRegistra
        !MotivoNoIntegrado = p_RE.MotivoNoIntegrado
        !FechaMotivo = p_RE.FechaMotivo
        !Trasladar = p_RE.Trasladar
        !Suministrador = p_RE.Suministrador
        !Pedido = p_RE.Pedido
        !ProveedorPedido = p_RE.ProveedorPedido
        !CausaRiesgoPedido = p_RE.CausaRiesgoPedido
        !RequiereRiesgoDeBiblioteca = p_RE.RequiereRiesgoDeBiblioteca
        !CodRiesgoBiblioteca = p_RE.CodRiesgoBiblioteca
        !RiesgoPendienteRetipificacion = p_RE.RiesgoPendienteRetipificacion
        .Update
    End With
    
    Save = "OK"
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en RiesgoExternoRepository.Save: " & Err.Description
End Function

Public Function Delete(ByVal p_IDRiesgoExt As String, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    db.Execute "DELETE FROM TbRiesgosAIntegrar WHERE IDRiesgoExt=" & p_IDRiesgoExt
    Delete = "OK"
    Exit Function

errores:
    p_Error = "Error en RiesgoExternoRepository.Delete: " & Err.Description
End Function
