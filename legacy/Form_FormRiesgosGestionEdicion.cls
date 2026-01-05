
Option Compare Database
Option Explicit

Private m_URLInforme As String
Private WithEvents m_FormRiesgo As Form_FormRiesgo

Private m_Error As String



Private Sub ComandoAbrirProyecto_Click()
    Dim ws As DAO.Workspace
    Dim blnEnTrans As Boolean
    Dim m_EdicionUltima As Edicion
    Dim m_ObjDocumento As Documento
    Dim m_ArchivosMovidos As Scripting.Dictionary

    On Error GoTo errores
    
    
    
    m_Error = ""
    pregunta = MsgBox("¿Desea realmente quitar la última publicación y convertir esa edición e la edición activa?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Abrir gestión de riesgos")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If m_ObjProyectoActivo Is Nothing Then
        m_Error = "No hay proyecto activo."
        Err.Raise 1000
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set m_EdicionUltima = m_ObjProyectoActivo.EdicionUltima
    If m_ObjProyectoActivo.Error <> "" Then
        m_Error = m_ObjProyectoActivo.Error
        Err.Raise 1000
    End If
    If Not m_EdicionUltima Is Nothing Then
        Set m_ObjDocumento = m_EdicionUltima.DocumentoObj
        If m_EdicionUltima.Error <> "" Then
            m_Error = m_EdicionUltima.Error
            Err.Raise 1000
        End If
        If Not m_ObjDocumento Is Nothing Then
            TxArchivos_PrepararMover m_ObjDocumento.URLAdjunto, m_ArchivosMovidos, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        ElseIf m_EdicionUltima.URLArchivoInforme <> "" Then
            TxArchivos_PrepararMover m_EdicionUltima.URLArchivoInforme, m_ArchivosMovidos, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    
    ws.BeginTrans
    blnEnTrans = True
    m_ObjProyectoActivo.Abrir m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ws.CommitTrans
    blnEnTrans = False
    
    TxArchivos_Confirmar m_ArchivosMovidos
    Set m_ObjProyectoActivo = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.cmdSalir.SetFocus
    Form_FormRiesgosGestion.ComandoActualizarContador_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    On Error Resume Next
    If blnEnTrans Then
        ws.Rollback
    End If
    TxArchivos_Deshacer m_ArchivosMovidos
    On Error GoTo 0
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAbrirProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminarEdicion_Click()
    Dim ws As DAO.Workspace
    Dim blnEnTrans As Boolean
    Dim m_EdicionAnterior As Edicion
    Dim m_ObjDocumento As Documento
    Dim m_ArchivosMovidos As Scripting.Dictionary
    
    On Error GoTo errores
    
    m_Error = ""
    
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay edicion activa."
        Err.Raise 1000
    End If
    If Form_FormRiesgosGestion.blnPermitidoEditar = False Then
        m_Error = "Usuario no autorizado a esa operacion"
        Err.Raise 1000
    End If
    If m_ObjProyectoActivo Is Nothing Then
        Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
        If m_ObjEdicionActiva.Error <> "" Then
            m_Error = m_ObjEdicionActiva.Error
            Err.Raise 1000
        End If
    End If
    
    pregunta = MsgBox("¿Desea realmente borrar la edicion activa?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar edicion")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Set m_ObjDocumento = m_ObjEdicionActiva.DocumentoObj
    If m_ObjEdicionActiva.Error <> "" Then
        m_Error = m_ObjEdicionActiva.Error
        Err.Raise 1000
    End If
    If Not m_ObjDocumento Is Nothing Then
        TxArchivos_PrepararMover m_ObjDocumento.URLAdjunto, m_ArchivosMovidos, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Set m_EdicionAnterior = m_ObjEdicionActiva.EdicionAnterior
    If m_ObjEdicionActiva.Error <> "" Then
        m_Error = m_ObjEdicionActiva.Error
        Err.Raise 1000
    End If
    If Not m_EdicionAnterior Is Nothing Then
        Set m_ObjDocumento = m_EdicionAnterior.DocumentoObj
        If m_EdicionAnterior.Error <> "" Then
            m_Error = m_EdicionAnterior.Error
            Err.Raise 1000
        End If
        If Not m_ObjDocumento Is Nothing Then
            TxArchivos_PrepararMover m_ObjDocumento.URLAdjunto, m_ArchivosMovidos, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    blnEnTrans = True
    
    m_ObjEdicionActiva.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ws.CommitTrans
    blnEnTrans = False
    
    TxArchivos_Confirmar m_ArchivosMovidos
    
    Set m_ObjProyectoActivo = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjProyectoActivo Is Nothing Then
        Set m_ObjEdicionActiva = m_ObjProyectoActivo.EdicionUltima
        If m_ObjProyectoActivo.Error <> "" Then
            m_Error = m_ObjProyectoActivo.Error
            Err.Raise 1000
        End If
    End If
    
    Me.Parent.cmdSalir.SetFocus
    Form_FormRiesgosGestion.ComandoActualizarContador_Click
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    On Error Resume Next
    If blnEnTrans Then
        ws.Rollback
    End If
    TxArchivos_Deshacer m_ArchivosMovidos
    On Error GoTo 0
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarEdicion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoAltaRiesgo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    If Form_FormRiesgosGestion.blnPermitidoEditar = False Then
        m_Error = "No Está permitido dar de alta más riesgos"
        Err.Raise 1000
    End If
    
    If FormularioAbierto("FormRiesgo") Then
        DoCmd.Close acForm, "FormRiesgo", acSaveNo
    End If
    Set m_ObjRiesgoActivo = Nothing
    Set m_ObjRiesgoAlInicio = Nothing
    m_EsAlta = EnumSiNo.S¡
    DoCmd.OpenForm "FormRiesgo"
    If FormularioAbierto("FormRiesgo") Then
        Set m_FormRiesgo = Forms("FormRiesgo")
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAltaRiesgo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGenerarInforme_Click()
    Dim m_URL As String
    Dim m_fechaUltimaEdicion As String
    Dim m_FechaCierre As String
    Dim m_FechaPub As String
    Dim m_edicion As Edicion
    
    On Error GoTo errores
    m_Error = ""
    
    If m_ObjProyectoActivo Is Nothing Then
        m_Error = "No hay proyecto activo."
        Err.Raise 1000
    End If
    
    ' Validar si hay cambios sin guardar? (Opcional)
    
    ' Obtener edición actual
    Set m_edicion = m_ObjProyectoActivo.EdicionActiva
    If m_ObjProyectoActivo.Error <> "" Then
        m_Error = m_ObjProyectoActivo.Error
        Err.Raise 1000
    End If
    If m_edicion Is Nothing Then
        m_Error = "No se ha encontrado la edición actual."
        Err.Raise 1000
    End If
    
    ' Fechas
    m_FechaCierre = Format(Date, "dd/mm/yyyy")
    m_FechaPub = Format(Date, "dd/mm/yyyy")
    
    ' Generar Informe
    DoCmd.Hourglass True
    m_URL = GenerarInformeEdicionHTML(m_edicion, 0, m_FechaCierre, m_FechaPub, m_Error)
    If m_Error <> "" Then Err.Raise 1000

    DoCmd.Hourglass False
    
    If m_URL <> "" Then
        MsgBox "Informe generado correctamente en:" & vbCrLf & m_URL, vbInformation, "Informe Generado"
        Application.FollowHyperlink m_URL
    Else
        m_Error = "No se ha podido generar el informe (URL vacía)."
        Err.Raise 1000
    End If
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Error al generar informe: " & Err.Description
        MsgBox m_Error, vbCritical, "Error"
    Else
        MsgBox m_Error, vbExclamation, "Advertencia"
    End If
End Sub

Private Sub ComandoVaciarCache_Click()
    Dim pregunta As VbMsgBoxResult
    Dim m_Error As String
    
    On Error GoTo errores
    m_Error = ""
    
    If m_ObjProyectoActivo Is Nothing Then Exit Sub
    
    pregunta = MsgBox("¿Desea vaciar la caché del Control de Cambios para este proyecto? Esto obligará a recalcular todos los cambios en el próximo informe.", vbQuestion + vbYesNo, "Vaciar Caché")
    If pregunta <> vbYes Then Exit Sub
    
    DoCmd.Hourglass True
    ControlCambios_LimpiarCache m_ObjProyectoActivo.IDProyecto, 0, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    DoCmd.Hourglass False
    MsgBox "Caché vaciada correctamente.", vbInformation, "Éxito"
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        MsgBox "Error al vaciar caché: " & Err.Description, vbCritical, "Error"
    Else
        MsgBox m_Error, vbExclamation, "Advertencia"
    End If
End Sub

Private Sub CheckUsarCache_AfterUpdate()
    Dim m_Error As String
    Dim valor As EnumSiNo
    
    On Error GoTo errores
    If m_ObjProyectoActivo Is Nothing Then Exit Sub
    
    ' Asumiendo que CheckUsarCache es un CheckBox donde True/False
    ' Si es CheckBox: -1 es True, 0 es False.
    If Me.CheckUsarCache.Value = True Then
        valor = EnumSiNo.S¡
    Else
        valor = EnumSiNo.No
    End If
    
    ControlCambios_SetUsarCacheProyecto m_ObjProyectoActivo.IDProyecto, valor, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    If valor = EnumSiNo.No Then
        If MsgBox("¿Desea vaciar la caché existente para este proyecto?", vbYesNo + vbQuestion, "Limpiar Caché") = vbYes Then
            ControlCambios_LimpiarCache m_ObjProyectoActivo.IDProyecto, 0, m_Error
            If m_Error <> "" Then Err.Raise 1000
            MsgBox "Caché vaciada correctamente.", vbInformation
        End If
    End If
    
    Exit Sub
errores:
    MsgBox "Error al guardar configuración de caché: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
    ' Cargar estado del checkbox si existe
    On Error Resume Next
    If Not m_ObjProyectoActivo Is Nothing Then
        If ControlCambios_UsarCacheProyecto(m_ObjProyectoActivo.IDProyecto) = EnumSiNo.S¡ Then
            Me.CheckUsarCache.Value = True
        Else
            Me.CheckUsarCache.Value = False
        End If
    End If
End Sub

Private Function TxArchivos_PrepararMover( _
                                        ByVal p_RutaOrigen As String, _
                                        ByRef p_Movidos As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_RutaTemp As String
    Dim m_BaseTemp As String
    Dim m_NombreTemp As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_RutaOrigen = "" Then
        Exit Function
    End If
    If Not FSO.FileExists(p_RutaOrigen) Then
        Exit Function
    End If
    If FicheroAbierto(p_RutaOrigen) Then
        p_Error = "El archivo esta abierto: " & p_RutaOrigen
        Err.Raise 1000
    End If
    If p_Movidos Is Nothing Then
        Set p_Movidos = New Scripting.Dictionary
        p_Movidos.CompareMode = TextCompare
    End If
    If p_Movidos.Exists(p_RutaOrigen) Then
        Exit Function
    End If
    
    If Not m_ObjEntorno Is Nothing Then
        m_BaseTemp = m_ObjEntorno.URLDirectorioTemporalAGEDO
    End If
    If m_BaseTemp = "" Then
        m_BaseTemp = Environ("TEMP")
    End If
    If Right(m_BaseTemp, 1) <> "\" Then
        m_BaseTemp = m_BaseTemp & "\"
    End If
    If Not FSO.FolderExists(m_BaseTemp) Then
        FSO.CreateFolder m_BaseTemp
    End If
    
    m_NombreTemp = "tx_borrado_" & Format(Now, "yyyymmdd_hhnnss") & "_" & _
                    Replace(Replace(Replace(FSO.GetFileName(p_RutaOrigen), " ", "_"), ":", "_"), "/", "_")
    m_RutaTemp = m_BaseTemp & m_NombreTemp
    If FSO.FileExists(m_RutaTemp) Then
        m_RutaTemp = m_RutaTemp & "_" & CStr(CLng(Timer * 1000))
    End If
    
    FSO.MoveFile p_RutaOrigen, m_RutaTemp
    p_Movidos.Add p_RutaOrigen, m_RutaTemp
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "TxArchivos_PrepararMover ha producido el error: " & Err.Description
    End If
End Function

Private Sub TxArchivos_Deshacer(ByRef p_Movidos As Scripting.Dictionary)
    Dim m_RutaOrigen As Variant
    Dim m_RutaTemp As String
    
    On Error Resume Next
    
    If p_Movidos Is Nothing Then
        Exit Sub
    End If
    
    For Each m_RutaOrigen In p_Movidos.Keys
        m_RutaTemp = p_Movidos(m_RutaOrigen)
        If FSO.FileExists(m_RutaTemp) Then
            If Not FSO.FileExists(CStr(m_RutaOrigen)) Then
                FSO.MoveFile m_RutaTemp, CStr(m_RutaOrigen)
            End If
        End If
    Next
End Sub

Private Sub TxArchivos_Confirmar(ByRef p_Movidos As Scripting.Dictionary)
    Dim m_RutaOrigen As Variant
    Dim m_RutaTemp As String
    
    On Error Resume Next
    
    If p_Movidos Is Nothing Then
        Exit Sub
    End If
    
    For Each m_RutaOrigen In p_Movidos.Keys
        m_RutaTemp = p_Movidos(m_RutaOrigen)
        If FSO.FileExists(m_RutaTemp) Then
            FSO.DeleteFile m_RutaTemp, True
        End If
    Next
End Sub

