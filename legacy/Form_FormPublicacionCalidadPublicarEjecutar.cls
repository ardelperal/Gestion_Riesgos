Option Compare Database
Option Explicit

Private m_Error As String
Public Event Publicado()


Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




Private Sub ComandoGenerarInforme_Click()
    Dim m_fechaUltimaEdicion As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.ComboUltimaEdicion = "Sí" Then
        m_fechaUltimaEdicion = Date
    End If
    
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

Private Sub ComandoPublicar_Click()
    Dim m_fechaUltimaEdicion As String
    Dim m_FechaRef As String
    Dim m_ConCorreoAlRAC As EnumSiNo
    
    On Error GoTo errores
    
    
    m_Error = ""
   
   
    
    pregunta = MsgBox("¿Desea enviar el correo de publicación también al RAC?", vbExclamation + vbYesNo + vbDefaultButton2, "Con envío al RAC")
    If pregunta = vbYes Then
        If Nz(Me.CorreoRAC, "") = "" Then
            pregunta = MsgBox("No está definido el correo del RAC en esta gestión de riesgos", vbExclamation, _
            "Con envío al RAC")
            Exit Sub
        End If
        m_ConCorreoAlRAC = EnumSiNo.Sí
    Else
        m_ConCorreoAlRAC = EnumSiNo.No
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Nz(Me.ComboUltimaEdicion, "") = "Sí" Then
        m_fechaUltimaEdicion = Date
    End If
    'm_FechaRef = "11/07/2025"
     m_URLInforme = PublicarEdicionTransaccional(p_Edicion:=m_ObjEdicionActiva, _
                            p_FechaCierre:=m_fechaUltimaEdicion, _
                            p_RegenerandoCambiosDesdeElInicio:=EnumSiNo.No, _
                            p_FechaSiguientePublicacion:=Nz(Me.FechaMaxProximaPublicacion, ""), _
                            p_fechaRef:=m_FechaRef, _
                            p_CorreoRAC:=Nz(Me.CorreoRAC, ""), _
                            p_ConCorreoAlRAC:=m_ConCorreoAlRAC, _
                            p_EnviarCorreo:=EnumSiNo.Sí, _
                            p_ActualizarTareasCalidad:=EnumSiNo.Sí, _
                            p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    
    If FSO.FileExists(m_URLInforme) Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("¿Desea ver el informe de la edición generada?", vbExclamation + vbYesNo + vbDefaultButton1, "Informe generado")
        If pregunta = vbYes Then
            m_URLHTMLActivo = m_URLInforme
'            If FormularioAbierto("FormWeb") Then
'                DoCmd.Close acForm, "FormWeb", acSaveNo
'            End If
'            DoCmd.OpenForm "FormWeb"
           Ejecutar Me.hWnd, "open", m_URLInforme, "", "", 1
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
'    If FormularioAbierto("FormRiesgosGestion") Then
'        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
'    End If
'    If FormularioAbierto("FormProyectosGestion") Then
'        Set Form_FormProyectosGestion.m_ObjProyectoSeleccionado = Nothing
'        Form_FormProyectosGestion.ListaFiltrados_Click
'    End If
    RaiseEvent Publicado
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    If FormularioAbierto("FormPublicacionCalidad") Then
        DoCmd.Close acForm, "FormPublicacionCalidad", acSaveNo
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    
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

Private Sub ComandoSalir_Click()
    
    On Error GoTo errores
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSalir_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub EstablecerFechaMax()
    
    Dim m_FechaMax As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_FechaMax = CalcularFechaMaxProximaPublicacion(p_Edicion:=m_ObjEdicionActiva, _
                                                    p_FechaPublicacion:=Date, _
                                                    p_EsUltimaEdicion:=Me.ComboUltimaEdicion, _
                                                    p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If IsDate(m_FechaMax) Then
        Me.FechaMaxProximaPublicacion = m_FechaMax
    Else
        Me.FechaMaxProximaPublicacion = Null
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al EstablecerFechaMax se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComboUltimaEdicion_AfterUpdate()
    EstablecerFechaMax
End Sub

Private Sub Form_Load()
    Dim m_FechaPrevistaCierre As String
    Dim m_CorreoRAC As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    If m_ObjEdicionActiva Is Nothing Then
        m_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    With m_ObjEdicionActiva.Proyecto
        m_FechaPrevistaCierre = .FechaPrevistaCierreCalculada
        If IsDate(m_FechaPrevistaCierre) Then
            If DateDiff("d", CDate(m_FechaPrevistaCierre), Date) <= 15 Then
                Me.ComboUltimaEdicion = "Sí"
            Else
                Me.ComboUltimaEdicion = "No"
            End If
        Else
            Me.ComboUltimaEdicion = "No"
        End If
        Me.ComboUltimaEdicion = "No"
        If IsDate(.Expediente.FechaInicioContrato) Then
            Me.FechaInicioContrato = .Expediente.FechaInicioContrato
        End If
        If IsDate(.Expediente.FechaFinContrato) Then
            Me.FechaFinContrato = .Expediente.FechaFinContrato
        End If
        m_CorreoRAC = .CorreoRACCalculado
        If m_CorreoRAC <> "" Then
            Me.CorreoRAC = m_CorreoRAC
        End If
    End With
    EstablecerFechaMax
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
   Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



