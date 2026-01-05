VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormPublicacionCalidadPublicar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_Error As String

Private WithEvents m_FormAnexos As Form_FormAnexos
Attribute m_FormAnexos.VB_VarHelpID = -1
Private WithEvents m_FormPublicar As Form_FormPublicacionCalidadPublicarEjecutar
Attribute m_FormPublicar.VB_VarHelpID = -1
Private WithEvents m_FormMotivos As Form_FormPublicacionCalidadRechazoObservaciones
Attribute m_FormMotivos.VB_VarHelpID = -1

Private Sub ComandoEnviarCorreoTecnicoPropuestaPublicacion_Click()
    Dim m_EdicionCorreoRevision As EdicionCorreoRevision
    
    On Error GoTo errores
    
    m_Error = ""
    pregunta = MsgBox("¿Desea enviar un correo a los responsables para que revisen la edición para la propuesta de publicación?", vbExclamation + vbYesNo + vbDefaultButton1, "Correo a responsables")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_EdicionCorreoRevision = New EdicionCorreoRevision
    With m_EdicionCorreoRevision
        .IDEdicion = m_ObjEdicionActiva.IDEdicion
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set m_ObjEdicionActiva = Constructor.getEdicion(p_IDEdicion:=m_ObjEdicionActiva.IDEdicion, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Correo enviado correctamente", vbInformation, "Correo a responsables")
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEnviarCorreoTecnicoPropuestaPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGenerarInformePublicabilidad_Click()
    Dim m_URL As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_URL = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=m_ObjEdicionActiva, p_hWnd:=Me.hWnd, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLHTMLActivo = m_URL
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInformePublicabilidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
   
End Sub

Private Sub ComandoHistorialPublicaciones_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    HTMLENTXT m_ObjEdicionActiva.HTMLHistorialPublicaciones, , EnumSiNo.No, m_Error
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoHistorialPublicaciones_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoVerCorreosEnviados_Click()
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    HTMLENTXT m_ObjEdicionActiva.HTMLHistorialcorreosResponsables, , EnumSiNo.No, m_Error
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerCorreosEnviados_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerRechazo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        DoCmd.Close acForm, "FormPublicacionCalidadRechazoObservaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadRechazoObservaciones"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerRechazo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoPublicar_Click()
    
    
    
    On Error GoTo errores
    
    
    m_Error = ""
   
    If FormularioAbierto("FormPublicacionCalidadPublicarEjecutar") Then
        DoCmd.Close acForm, "FormPublicacionCalidadPublicarEjecutar", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadPublicarEjecutar"
    If FormularioAbierto("FormPublicacionCalidadPublicarEjecutar") Then
        Set m_FormPublicar = Forms("FormPublicacionCalidadPublicarEjecutar")
    End If
    
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPublicar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRechazar_Click()
    On Error GoTo errores
    
    
    m_Error = ""
    
    If m_ObjEdicionActiva.PublicacionRechazada = EnumSiNo.Sí Then
        pregunta = MsgBox("En Estos momentos ya está rechazada la propuesta de publicación" & vbNewLine & _
                        "¿Desea ver las observaciones registradasí", vbExclamation + vbYesNo + vbDefaultButton1, "Rechazo")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    Else
        pregunta = MsgBox("Para poder rechazar la propuesta rellene, en el siguiente formulario, los motivos", vbExclamation + vbYesNo + vbDefaultButton1, "Rechazo")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        DoCmd.Close acForm, "FormPublicacionCalidadRechazoObservaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormPublicacionCalidadRechazoObservaciones"
    If FormularioAbierto("FormPublicacionCalidadRechazoObservaciones") Then
        Set m_FormMotivos = Forms("FormPublicacionCalidadRechazoObservaciones")
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRechazar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerCambios_Click()
    Dim m_ColCambios As Scripting.Dictionary
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Avance "Generando HTML con los cambios"
    
    With m_ObjEdicionActiva
        If .EdicionAnterior Is Nothing Then
            m_Error = "Es la primera edición"
            Err.Raise 1000
        End If
        If .EsActivo = EnumSiNo.Sí Then
            .BorrarCambiosEdicion m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            .GrabarCambiosEnEdicion m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        Set m_ColCambios = .colCambiosConEdicionAnterior
        m_Error = .Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ColCambios Is Nothing Then
            m_Error = "No hay cambios entre ediciones"
            Err.Raise 1000
        End If
        MostrarEdicionCambios m_ColCambios, .Edicion, _
                            .EdicionAnterior.Edicion, _
                            .Proyecto.NombreProyecto, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerCambios_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    

    
    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "No hay ninguna edición activa"
        Err.Raise 1000
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub


Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    
    On Error GoTo errores
    m_Error = ""
    With m_ObjEdicionActiva
        If .Proyecto.EnUTE = "Sí" Then
            Me.lbl1Requerido.Caption = "Requerido"
            If .TieneAnexoEvidenciaUTE = EnumSiNo.Sí Then
                Me.lbl1Complementado.Caption = "Complementado"
                Me.lbl1Complementado.ForeColor = m_ColorComplementado
            Else
                Me.lbl1Complementado.Caption = "No complementado"
                Me.lbl1Complementado.ForeColor = m_ColorNoComplementado
            End If
        Else
            Me.lbl1Requerido.Caption = "No Requerido"
            Me.lbl1Complementado.Caption = "N/A"
            Me.lbl1Complementado.ForeColor = m_ColorComplementado
        End If
        If .EvidenciasSuministradoresRequeridas = EnumSiNo.Sí Then
            Me.lbl2Requerido.Caption = "Requerido"
            If .EvidenciasSuministradoresCompletadas = EnumSiNo.Sí Then
                Me.lbl2Complementado.Caption = "Complementado"
                Me.lbl2Complementado.ForeColor = m_ColorComplementado
            Else
                Me.lbl2Complementado.Caption = "No complementado"
                Me.lbl2Complementado.ForeColor = m_ColorNoComplementado
                
            End If
        Else
            Me.lbl2Requerido.Caption = "No requerido"
            Me.lbl2Complementado.Caption = "N/A"
            Me.lbl2Complementado.ForeColor = m_ColorComplementado
        End If
        If .EsActivo = EnumSiNo.Sí Then
            Me.ComandoRechazar.Enabled = True
            Me.ComandoPublicar.Enabled = True
            Me.ComandoPublicar.Caption = "Publicar"
            Me.ComandoGenerarInformePublicabilidad.Enabled = True
            Me.ComandoEnviarCorreoTecnicoPropuestaPublicacion.Enabled = True
            Me.lblPublicacion.Visible = True
            If Not IsDate(.FechaPreparadaParaPublicar) Then
                Me.lblPublicacion.Caption = "SIN PROPUESTA TÉCNICA"
                Me.lblPublicacion.ForeColor = m_ColorNoComplementado
            Else
                If .PropuestaRechazadaPorCalidadFecha = "" Then
                    Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") ESPERANDO A CALIDAD"
                    Me.lblPublicacion.ForeColor = m_ColorComplementado
                    Me.ComandoVerRechazo.Visible = False
                Else
                    Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") RECHAZADA (" & .PropuestaRechazadaPorCalidadFecha & ")"
                    Me.lblPublicacion.ForeColor = m_ColorNoComplementado
                    Me.ComandoVerRechazo.Visible = True
                End If
            End If
            
            
        Else
            Me.ComandoRechazar.Enabled = False
            Me.ComandoPublicar.Enabled = False
            Me.ComandoGenerarInformePublicabilidad.Enabled = False
            Me.ComandoVerRechazo.Visible = False
            Me.ComandoEnviarCorreoTecnicoPropuestaPublicacion.Enabled = False
            Me.ComandoPublicar.Caption = "Publicado"
            If .FechaPublicacion <> "" Then
                Me.lblPublicacion.Visible = True
                If IsDate(.FechaPreparadaParaPublicar) Then
                    Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") PUBLICADO EL " & .FechaPublicacion
                Else
                    Me.lblPublicacion.Caption = "PROPUESTA PUBLICADA EL " & .FechaPublicacion
                End If
                
                Me.lblPublicacion.ForeColor = m_ColorComplementado
                
            Else
                Me.lblPublicacion.Visible = False
            End If
        End If
        
    End With
    

    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Establecer datos ha producido el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub m_FormMotivos_Motivado(p_Motivo As String)
    
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If p_Motivo = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    With m_ObjEdicionActiva
        .PropuestaRechazadaPorCalidadMotivo = p_Motivo
        .RechazoPropuestaParaPublicacion m_Error
        If m_Error <> "" Then
           Err.Raise 1000
        End If
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Propuesta de publicación rechazada. Le ha llegado un correo a los responsables del proyecto", vbInformation, "Rechazo de propuesta")
    If FormularioAbierto("FormPublicacionCalidad") Then
        DoCmd.Close acForm, "FormPublicacionCalidad", acSaveNo
    End If
    If FormularioAbierto("FormRiesgosGestion") Then
        If Forms("FormRiesgosGestion").FormDetalle.SourceObject = "FormRiesgosGestionEdicion" Then
            EstablecerlblRechazado Forms("FormRiesgosGestionEdicion")
        End If
        
    End If
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormMotivos_Motivado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub m_FormPublicar_Publicado()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    DoCmd.Close acForm, Me.Name
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormPublicar_Publicado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

