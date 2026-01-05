VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "Form_FormPublicacionTecnicoPropuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_Error As String


Private Sub ComandoGenerarInforme_Click()
    Dim m_fechaUltimaEdicion As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   m_fechaUltimaEdicion = Date
    
    GenerarInformePublicacion m_ObjEdicionActiva, , m_fechaUltimaEdicion, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGenerarInforme_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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


Private Sub ComandoVerCambios_Click()
    Dim m_ColCambios As Scripting.Dictionary
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Avance "Generando HTML con los cambios"
    
    With m_ObjEdicionActiva
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
            m_Error = "No hay cambios entre Ediciones"
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
    AvanceCerrar
    
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
    AvanceCerrar
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
    AvanceCerrar
    pregunta = MsgBox(m_Error, vbCritical, "Error")

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
            Me.lbl2Requerido.Caption = "No Requerido"
            Me.lbl2Complementado.Caption = "N/A"
            Me.lbl2Complementado.ForeColor = m_ColorComplementado
        End If
        If .EsActivo = EnumSiNo.Sí Then
            If blnPermitidoEscribir = True Then
                Me.ComandoProponerPublicacion.Enabled = True
            Else
                Me.ComandoProponerPublicacion.Enabled = False
            End If
            If .FechaPreparadaParaPublicar = "" Then
                Me.ComandoProponerPublicacion.Caption = "Proponer para publicación"
            Else
                Me.ComandoProponerPublicacion.Caption = "Quitar propuesta para publicación"
            End If
            Me.ComandoGenerarInformePublicabilidad.Enabled = True
            Me.lblPublicacion.Visible = True
            If Not IsDate(.FechaPreparadaParaPublicar) Then
                Me.lblPublicacion.Caption = "SIN PROPUESTA TécnicA"
                Me.lblPublicacion.ForeColor = m_ColorNoComplementado
            Else
                If .PropuestaRechazadaPorCalidadFecha = "" Then
                    Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") ESPERANDO A CALIDAD"
                    Me.lblPublicacion.ForeColor = m_ColorComplementado
                    Me.ComandoVerRechazo.Visible = False
                Else
                    If CDate(.FechaPreparadaParaPublicar) > CDate(.PropuestaRechazadaPorCalidadFecha) Then
                        Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") ESPERANDO A CALIDAD"
                        Me.lblPublicacion.ForeColor = m_ColorComplementado
                        Me.ComandoVerRechazo.Visible = False
                    Else
                        Me.lblPublicacion.Caption = "PROPUESTA (" & .FechaPreparadaParaPublicar & ") RECHAZADA (" & .PropuestaRechazadaPorCalidadFecha & ")"
                        Me.lblPublicacion.ForeColor = m_ColorNoComplementado
                        Me.ComandoVerRechazo.Visible = True
                    End If
                    
                End If
            End If
            
        Else
            Me.ComandoProponerPublicacion.Enabled = False
            Me.ComandoGenerarInformePublicabilidad.Enabled = False
            Me.ComandoVerRechazo.Visible = False
            Me.ComandoProponerPublicacion.Caption = "Publicado"
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



Private Sub ComandoProponerPublicacion_Click()
    
    Dim m_URL As String
    On Error GoTo errores
    
    
    m_Error = ""
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If Me.ComandoProponerPublicacion.Caption <> "Proponer para publicación" And _
        Me.ComandoProponerPublicacion.Caption <> "Quitar propuesta para publicación" Then
        m_Error = "Texto del botón no reconocido"
        Err.Raise 1000
    End If
    
    
    
    If Me.ComandoProponerPublicacion.Caption = "Proponer para publicación" Then
        
        With m_ObjEdicionActiva
            If .Publicable = EnumSiNo.No Then
                m_URL = GenerarInformePublicabilidadEdicionInteractivoHTML(p_Edicion:=m_ObjEdicionActiva, p_hWnd:=Me.hWnd, p_Error:=m_Error)
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                m_URLHTMLActivo = m_URL
            
                
                m_Error = "No ha superado el informe de publicabilidad" & vbNewLine & _
                            "Solvente las no conformidades y repita el proceso"
                Err.Raise 1000
            End If
             VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            pregunta = MsgBox("No olvide lo siguiente: " & vbNewLine & vbTab & _
                                "Riesgos del suministrador" & vbNewLine & vbTab & _
                                "Riesgos de los pedidos" & vbNewLine & _
                                "¿Desea seguir con la propuesta?", vbExclamation + vbYesNo + vbDefaultButton1, "Nota")
            If pregunta <> vbYes Then
                AvanceCerrar
                Exit Sub
            End If
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            .ProponerParaPublicacion m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            pregunta = MsgBox("Edición propuesta para publicar. Se le ha enviado un correo a Calidad para que la publique.", _
                            vbInformation, "Propuesta generada")
            
            
        End With
        AvanceCerrar
        Set m_ObjProyectoActivo = Constructor.getProyecto(p_IDProyecto:=m_ObjProyectoActivo.IDProyecto, p_Error:=m_Error)
        EstablecerTareasCalidad p_Reseteando:=EnumSiNo.Sí, p_Regularizando:=EnumSiNo.No, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If FormularioAbierto("FormPublicacionTecnico") Then
            DoCmd.Close acForm, "FormPublicacionTecnico", acSaveNo
        End If
        If FormularioAbierto("FormRiesgosGestion") Then
            DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
        End If
        
        DoCmd.Close acForm, Me.Name, acSaveNo
        
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea quitar la propuesta de publicación", vbExclamation + vbYesNo, "Quitar propuesta")
    If pregunta <> vbYes Then
        AvanceCerrar
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    With m_ObjEdicionActiva
         .QuitarPropuestaParaPublicacion m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Propuesta de publicación eliminada. Se le ha enviado un correo a Calidad.", _
                    vbInformation, "Propuesta eliminada")
    
    
    AvanceCerrar
    Set m_ObjProyectoActivo = Constructor.getProyecto(p_IDProyecto:=m_ObjProyectoActivo.IDProyecto, p_Error:=m_Error)
    EstablecerTareasCalidad p_Reseteando:=EnumSiNo.Sí, p_Regularizando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormPublicacionTecnico") Then
        DoCmd.Close acForm, "FormPublicacionTecnico", acSaveNo
    End If
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoProponerPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub







