
Option Compare Database
Option Explicit

Public Event Alta()
Public Event Edicion()
Public blnPermitidoEditar As Boolean
Private blnAltaProyecto As Boolean
Private blnEdicionProyecto As Boolean
Private m_Error As String



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

Private Sub ComandoGrabar_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales" Then
        Form_FormGestionRiesgosDatosGenerales.RellenarDatosAlObjeto m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    grabar EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_EsAlta = EnumSiNo.No
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    pregunta = MsgBox("Datos guardados con éxito", vbInformation, "Guardar")
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function grabar(Optional ByRef p_SinEco As EnumSiNo = EnumSiNo.No, Optional ByRef p_Error As String) As String
    
    Dim blnHaHabidocambios As Boolean
    Dim blnAvisarNoInforme As Boolean
    Dim m_Usuario As Usuario
    Dim m_IDProyecto As String
    Dim m_Form As Form
    On Error GoTo errores
    If p_SinEco = Empty Then
        p_SinEco = EnumSiNo.No
    End If
    If m_ObjProyectoActivo Is Nothing Then
        Set m_ObjProyectoActivo = New Proyecto
        If Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales" Then
            With Me.FormDetalle
                m_ObjProyectoActivo.IDExpediente = .Controls("IDExpediente")
                m_ObjProyectoActivo.ParaInformeAvisos = .Controls("ParaInformeAvisos")
                m_ObjProyectoActivo.EnUTE = .Controls("EnUTE")
                m_ObjProyectoActivo.Elaborado = .Controls("Elaborado")
                m_ObjProyectoActivo.Revisado = .Controls("Revisado")
                m_ObjProyectoActivo.Aprobado = .Controls("Aprobado")
            End With
        End If
    End If
    blnHaHabidocambios = HaHabidoCambios(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not blnHaHabidocambios Then
        p_Error = "Ningún cambio que grabar"
        Err.Raise 1000
    End If
    
    If p_SinEco = EnumSiNo.No Then
        If m_ObjProyectoActivo.ParaInformeAvisos = "No" Then
            If m_ObjProyectoAlInicio Is Nothing Then
                blnAvisarNoInforme = True
            Else
                If m_ObjProyectoAlInicio.ParaInformeAvisos <> m_ObjProyectoActivo.ParaInformeAvisos Then
                    blnAvisarNoInforme = True
                End If
            End If
        End If
        If blnAvisarNoInforme = True Then
            
            pregunta = MsgBox("Al poner que no es para informes, implica que este proyecto no va a generar ningún correo electrónico" & vbNewLine & _
                                "¿Está de acuerdo con eso?", vbExclamation + vbYesNo + vbDefaultButton2, "Sin informes")
            If pregunta <> vbYes Then
                AvanceCerrar
                
                Exit Function
            End If
            
        End If
    End If
    
    With m_ObjProyectoActivo
        
        If m_ObjProyectoAlInicio Is Nothing Then
            CrearProyectoTransaccional m_ObjProyectoActivo, EsTecnico, Nothing, p_Error
        Else
            .Registrar m_ObjProyectoAlInicio, p_Error
        End If
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_IDProyecto = .IDProyecto
        If m_ObjProyectoAlInicio Is Nothing Then
            
            'se ha de generar la primera EDICIÓN
           Set m_ObjEdicionActiva = .EdicionActiva
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se ha generado la primera EDICIÓN automáticamente"
                Err.Raise 1000
            End If
            
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            RaiseEvent Alta
            
        Else
            RaiseEvent Edicion
                
        End If
        m_EsAlta = EnumSiNo.No
        Set m_ObjProyectoActivo = Constructor.getProyecto(m_IDProyecto, p_Error)
        Set m_ObjProyectoAlInicio = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    AvanceCerrar
    
    Me.NavResponsables.Enabled = True
    Me.NavRiesgosOferta.Enabled = True
    Me.NavSuministradores.Enabled = True
    Me.lblPendienteGrabacion.Visible = False
    Me.ComandoRegistrarRiesgos.Visible = True
    Me.ComandoRegistrarRiesgos.Enabled = True
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Grabar ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ComandoRegistrarRiesgos_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrarRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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

Private Sub Form_Load()
      
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.InsideHeight = 9195
    Me.InsideWidth = 13440
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.NavDatosGenerales.NavigationTargetName = "FormGestionRiesgosDatosGenerales"
    Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales"
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    AvanceCerrar
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

Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_Titulo As String
    
    
    On Error GoTo errores
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    
    If m_EsAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        Err.Raise 1000
    End If
    If m_ObjProyectoActivo Is Nothing Then
        Set m_ObjProyectoActivo = New Proyecto
        Me.ComandoRegistrarRiesgos.Visible = False
    Else
        Me.ComandoRegistrarRiesgos.Visible = True
    End If
    If m_EsAlta = EnumSiNo.Sí Then
        If EsTecnico = EnumSiNo.Sí Then
            m_Error = "Es un usuario no autorizado"
            Err.Raise 1000
        End If
        Me.NavResponsables.Enabled = False
        Me.NavRiesgosOferta.Enabled = False
        Me.NavSuministradores.Enabled = False
        m_Titulo = "ALTA DE GESTIÓN DE RIESGOS"
    Else
        If m_ObjProyectoAlInicio Is Nothing Then
            Set m_ObjProyectoAlInicio = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjProyectoAlInicio Is Nothing Then
                p_Error = "No se conocen los datos de la GESTIÓN de riesgos"
                Err.Raise 1000
            End If
        End If
        Me.NavResponsables.Enabled = True
        Me.NavRiesgosOferta.Enabled = True
        Me.NavSuministradores.Enabled = True
        m_Titulo = "EDICIÓN DE LA GESTIÓN DE RIESGOS: " & m_ObjProyectoActivo.NombreProyecto
    End If
    Me.NavDatosGenerales.Enabled = True
    Me.lblTitulo.Caption = m_Titulo
    
    If EsTecnico = EnumSiNo.Sí Then
        blnPermitidoEditar = False
    Else
        blnPermitidoEditar = True
    End If
    If blnPermitidoEditar = True Then
        Me.ComandoGrabar.Enabled = True
    Else
        Me.ComandoGrabar.Enabled = False
    End If
    'Me.FormDetalle.SourceObject = "FormGestionRiesgosDatosGenerales"
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    Dim m_Usuario As Usuario
    Dim m_frm As Form
    On Error GoTo errores
    
    If m_ObjProyectoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    
    With m_ObjProyectoAlInicio
        
        If .ParaInformeAvisos <> m_ObjProyectoActivo.ParaInformeAvisos Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .EnUTE <> m_ObjProyectoActivo.EnUTE Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .NombreUsuarioCalidad <> m_ObjProyectoActivo.NombreUsuarioCalidad Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .CorreoRAC <> m_ObjProyectoActivo.CorreoRAC Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Elaborado <> m_ObjProyectoActivo.Elaborado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Revisado <> m_ObjProyectoActivo.Revisado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Aprobado <> m_ObjProyectoActivo.Aprobado Then
            HaHabidoCambios = True
            Exit Function
        End If
        
       
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en el método HaHabidoCambios" & vbNewLine & Err.Description
    End If
End Function




