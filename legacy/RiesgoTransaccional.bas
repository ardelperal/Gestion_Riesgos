Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : RiesgoTransaccional
' Purpose   : Gestiona la persistencia transaccional de Riesgos Externos y su promoción a Riesgos.
'---------------------------------------------------------------------------------------

Public Function RegistrarRiesgoExternoTransaccional( _
                            p_RiesgoExt As RiesgoExterno, _
                            Optional p_ObjRiesgoExternoAlInicio As RiesgoExterno, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim blnAltaRiesgoExterno As Boolean
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    Set db = getdb(p_Error)
    If p_Error <> "" Then GoTo errores
    
    m_EnTransaccion = False
    m_FechaRef = Date
    
    ' 1. VALIDACIÓN
    m_MotivoNoOK = p_RiesgoExt.MotivoNoOK(p_ObjRiesgoExternoAlInicio, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_MotivoNoOK <> "" Then
        p_Error = m_MotivoNoOK
        Err.Raise 1000
    End If
    
    ' 2. INICIO TRANSACCIÓN
    ws.BeginTrans
    m_EnTransaccion = True
    
    ' 3. GESTIÓN PREVIA DEL TRASLADO (Para tener el IDRiesgo antes de guardar el externo)
    If p_RiesgoExt.Trasladar = "Sí" Then
        ' Si no tiene riesgo asociado aún, lo creamos
        If p_RiesgoExt.IDRiesgo = "" Then
            Set m_ObjRiesgo = New riesgo
            With m_ObjRiesgo
                .IDEdicion = p_RiesgoExt.IDEdicion
                If p_RiesgoExt.Suministrador <> "" Then
                    .EntidadDetecta = "Empresa"
                    .DetectadoPor = p_RiesgoExt.Suministrador
                ElseIf p_RiesgoExt.ProveedorPedido <> "" Then
                    .EntidadDetecta = "Persona"
                    .DetectadoPor = m_ObjUsuarioConectado.Nombre
                End If
                .CodRiesgoBiblioteca = p_RiesgoExt.CodRiesgoBiblioteca
                .Descripcion = p_RiesgoExt.Descripcion
                If p_RiesgoExt.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
                    .CausaRaiz = p_RiesgoExt.CausaRaiz
                End If
                .Estado = .ESTADOCalculadoTexto
                .FechaDetectado = p_RiesgoExt.FechaDetectado
                .FechaEstado = CStr(m_FechaRef)
                .CodigoRiesgo = .CodigoRiesgoCalculado
                .Origen = "Oferta"
                If p_RiesgoExt.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
                    .Priorizacion = .PriorizacionCalculado
                End If
                .RiesgoParaRetipificar = EnumSiNo.No
                .CodigoUnico = .CodigoUnicoCalculado
                
                ' Pasamos el objeto db para que use la transacción actual
                .AltaDesdeRiesgoExterno p_Error, db
                If p_Error <> "" Then Err.Raise 1000
                
                p_RiesgoExt.IDRiesgo = .IDRiesgo
            End With
        End If
    Else
        ' Si deja de trasladarse, borramos el riesgo asociado transaccionalmente
        Set m_ObjRiesgo = p_RiesgoExt.riesgo
        If Not m_ObjRiesgo Is Nothing Then
            If m_ObjRiesgo.Edicion.Edicion = 1 And m_ObjRiesgo.EsActivo = EnumSiNo.Sí Then
                m_ObjRiesgo.Borrar p_Error, db
                If p_Error <> "" Then Err.Raise 1000
                p_RiesgoExt.IDRiesgo = ""
            End If
        End If
    End If
    
    ' 4. ÚNICO GUARDADO DE TbRiesgosAIntegrar
    If p_ObjRiesgoExternoAlInicio Is Nothing Then
        p_RiesgoExt.IDRiesgoExt = p_RiesgoExt.IDRiesgoExtCalculado
        m_SQL = "TbRiesgosAIntegrar"
    Else
        m_SQL = "SELECT * FROM TbRiesgosAIntegrar WHERE IDRiesgoExt=" & p_RiesgoExt.IDRiesgoExt
    End If
    
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjRiesgoExternoAlInicio Is Nothing Then
            .AddNew
            blnAltaRiesgoExterno = True
            .Fields("IDRiesgoExt") = p_RiesgoExt.IDRiesgoExt
            p_RiesgoExt.FechaAltaRegistro = Now()
            .Fields("FechaAltaRegistro") = p_RiesgoExt.FechaAltaRegistro
        Else
            .Edit
        End If
        
        .Fields("IDRiesgo") = IIf(p_RiesgoExt.IDRiesgo <> "", p_RiesgoExt.IDRiesgo, "")
        .Fields("CodRiesgo") = p_RiesgoExt.CodRiesgo
        .Fields("FechaDetectado") = p_RiesgoExt.FechaDetectado
        .Fields("Origen") = p_RiesgoExt.Origen
        .Fields("IDEdicion") = p_RiesgoExt.IDEdicion
        .Fields("Descripcion") = p_RiesgoExt.Descripcion
        .Fields("CausaRaiz") = p_RiesgoExt.CausaRaiz
        .Fields("UsuarioRegistra") = m_ObjUsuarioConectado.UsuarioRed
        .Fields("MotivoNoIntegrado") = p_RiesgoExt.MotivoNoIntegrado
        If IsDate(p_RiesgoExt.FechaMotivo) Then .Fields("FechaMotivo") = p_RiesgoExt.FechaMotivo
        
        .Fields("Trasladar") = p_RiesgoExt.Trasladar
        .Fields("Suministrador") = p_RiesgoExt.Suministrador
        .Fields("Pedido") = p_RiesgoExt.Pedido
        .Fields("ProveedorPedido") = p_RiesgoExt.ProveedorPedido
        .Fields("CausaRiesgoPedido") = p_RiesgoExt.CausaRiesgoPedido
        .Fields("RequiereRiesgoDeBiblioteca") = p_RiesgoExt.Edicion.Proyecto.RequiereRiesgoDeBiblioteca
        .Fields("CodRiesgoBiblioteca") = p_RiesgoExt.CodRiesgoBiblioteca
        .Fields("RiesgoPendienteRetipificacion") = p_RiesgoExt.RiesgoPendienteRetipificacion
        
        .Update
    End With
    rcdDatos.Close
    
    ' 5. ACTUALIZACIÓN DE CACHÉS Y NOTIFICACIONES
    If p_RiesgoExt.Trasladar = "Sí" And Not m_ObjRiesgo Is Nothing Then
        If p_RiesgoExt.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí And p_RiesgoExt.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
            If m_ObjRiesgo.RiesgoParaRetipificar = EnumSiNo.Sí Then
                EnviarCorreoRiesgoRequiereRetipificacion m_ObjRiesgo, p_Error
                If p_Error <> "" Then Err.Raise 1000
            End If
        End If
    End If

    CachePublicabilidad_RecalcularEdicionYResetear p_RiesgoExt.Edicion, , p_Error
    If p_Error <> "" Then Err.Raise 1000

    ws.CommitTrans
    m_EnTransaccion = False
    
    ' Limpieza de objetos para forzar recarga
    Set p_RiesgoExt.Edicion.ColRiesgosExternos = Nothing
    Set p_RiesgoExt.Edicion.Proyecto.ColRiesgosExternos = Nothing
    
    RegistrarRiesgoExternoTransaccional = ""
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en RegistrarRiesgoExternoTransaccional: " & Err.Description
    End If
    RegistrarRiesgoExternoTransaccional = p_Error
End Function

Public Function BorrarRiesgoExternoTransaccional( _
                                                    p_RiesgoExt As RiesgoExterno, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_ObjRiesgo As riesgo
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    Set db = getdb(p_Error)
    If p_Error <> "" Then GoTo errores
    
    m_EnTransaccion = False
    
    ' 1. VALIDACIÓN BORRADO (Riesgo asociado)
    Set m_ObjRiesgo = p_RiesgoExt.riesgo
    If Not m_ObjRiesgo Is Nothing Then
        m_ObjRiesgo.ValidacionEliminar m_MotivoNoOK
        If m_MotivoNoOK <> "" Then
            p_Error = "No se podría borrar el riesgo asociado por: " & vbNewLine & m_MotivoNoOK
            Err.Raise 1000
        End If
    End If
    
    ' 2. INICIO TRANSACCIÓN
    ws.BeginTrans
    m_EnTransaccion = True
    
    ' 3. BORRAR RIESGO ASOCIADO
    If Not m_ObjRiesgo Is Nothing Then
        m_ObjRiesgo.Borrar p_Error, db
        If p_Error <> "" Then Err.Raise 1000
    End If
    
    ' 4. BORRAR RIESGO EXTERNO
    db.Execute "DELETE * FROM TbRiesgosAIntegrar WHERE IDRiesgoExt=" & p_RiesgoExt.IDRiesgoExt
    
    ws.CommitTrans
    m_EnTransaccion = False
    
    BorrarRiesgoExternoTransaccional = ""
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en BorrarRiesgoExternoTransaccional: " & Err.Description
    End If
    BorrarRiesgoExternoTransaccional = p_Error
End Function






