
Option Compare Database
Option Explicit
Private blnRiesgoActivo As Boolean
Private m_EstadoRiesgo As EnumRiesgoEstado
Private blnPermitidoEditar As Boolean
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String

Private m_Error As String

Private Sub ComandoAceptacionRegistrar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Nz(Me.JustificacionAceptacionRiesgo, "") = "" Then
        m_Error = "Se ha de indicar alguna justificación"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.IDRiesgo = "" Then
        m_Error = "Primero ha de registrar el resto de campos del Riesgo"
        Err.Raise 1000
    End If
    m_ObjRiesgoActivo.AceptacionRegistrar Me.JustificacionAceptacionRiesgo, Date, m_ObjRiesgoActivo.EstadoEnum, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
    If IsDate(m_ObjRiesgoAlInicio.FechaJustificacionAceptacionRiesgo) Then
        Me.FechaJustificacionAceptacionRiesgo = m_ObjRiesgoAlInicio.FechaJustificacionAceptacionRiesgo
        
    End If
    If m_ObjRiesgoAlInicio.FechaRechazoAceptacionPorCalidad = "" Then
        Me.FechaRechazoAceptacionPorCalidad = Null
    Else
        Me.FechaRechazoAceptacionPorCalidad = m_ObjRiesgoAlInicio.FechaRechazoAceptacionPorCalidad
    End If
    ActualizarEtiquetaEstadoRiesgo
    ' Recargar y re-seleccionar
      Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
      If m_Error <> "" Then Err.Raise 1000
      
      Form_FormRiesgosGestion.SeleccionarNodo m_ObjRiesgoActivo, m_Error
      If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAceptacionRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ComandoAceptar_Click()
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    With m_ObjRiesgoActivo
        If Me.ComandoAceptar.Caption = "Aceptar" Then
            .AceptacionAprobar Date, m_EstadoRiesgo, m_Error
        Else
            .AceptacionAprobarQuitar m_EstadoRiesgo, m_Error
        End If
        
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    EstablecerDatosCalidad m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Form_FormRiesgosGestionRiesgo.EstablecerDatos
    Form_FormRiesgo.EstablecerDatos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAceptar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRechazar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    With m_ObjRiesgoActivo
        If Me.ComandoRechazar.Caption = "Rechazar" Then
            .AceptacionRechazo m_EstadoRiesgo, Date, m_Error
        Else
            .AceptacionRechazoQuitar m_EstadoRiesgo, m_Error
        End If
        
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    EstablecerDatosCalidad m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Form_FormRiesgosGestionRiesgo.EstablecerDatos
    Form_FormRiesgo.EstablecerDatos
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

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerDatosCalidad m_Error
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Edicion As Edicion
    Dim lst As ListBox
    Dim m_Mitigacion As Variant
    Dim i As Long
    
    On Error GoTo errores
    Me.AllowEdits = True
    Set lst = Me.ListaMitigacion
    lst.RowSource = "mitigación;Descripción"
    
    Set m_Edicion = m_ObjRiesgoActivo.Edicion
    If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = True
        End If
    End If
    If blnPermitidoEditar = True Then
        Me.AllowEdits = True
        If FormularioAbierto("FormRiesgosGestion") Then
            Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
            m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
        End If
    Else
        Me.AllowEdits = False
    End If
    If m_ObjRiesgoActivo.EsActivo = EnumSiNo.Sí Then
        blnRiesgoActivo = True
    Else
        blnRiesgoActivo = False
    End If
    For Each m_Mitigacion In m_ObjEntorno.ColMitigacionValores.keys
        lst.AddItem m_Mitigacion & ";" & Replace(m_ObjEntorno.ColMitigacionValores(m_Mitigacion), ";", ":")
    Next
    If m_ObjRiesgoActivo.Mitigacion <> "" Then
        For i = 1 To lst.ListCount - 1
            If lst.Column(0, i) = m_ObjRiesgoActivo.Mitigacion Then
                lst.Selected(i) = True
                Me.DescripcionMitigacion = lst.Column(1, i)
                Exit For
            End If
        Next
        EstablecerControlCombo Me.ListaMitigacion, EnumSiNo.Sí
    Else
        EstablecerControlCombo Me.ListaMitigacion, EnumSiNo.No
    End If
    If m_ObjRiesgoActivo.Mitigacion = "Aceptar" Then
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        If blnPermitidoEditar = True Then
            If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Aceptado Then
                Me.JustificacionAceptacionRiesgo.Enabled = False
                Me.ComandoAceptacionRegistrar.Enabled = False
            Else
                Me.JustificacionAceptacionRiesgo.Enabled = True
                Me.ComandoAceptacionRegistrar.Enabled = True
            End If
        Else
            Me.JustificacionAceptacionRiesgo.Enabled = False
            Me.ComandoAceptacionRegistrar.Enabled = False
        End If
        Exit Function
    Else
        HacerInVisiblesCamposOcultos
    End If
    If m_ObjRiesgoAlInicio Is Nothing Then
        HacerInVisiblesCamposOcultos
        
    End If
    
    
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function EstablecerDatosCalidad(Optional ByRef p_Error As String) As String
    
  
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_RiesgoEnAceptacion As EnumSiNo
    
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.ComandoAceptar.Visible = False
        Me.ComandoRechazar.Visible = False
        Exit Function
    End If
    If m_ObjRiesgoAlInicio Is Nothing Then
        Me.ComandoAceptar.Visible = False
        Me.ComandoRechazar.Visible = False
        Exit Function
    End If
    If Not IsDate(m_ObjRiesgoAlInicio.FechaMitigacionAceptar) Then
        Me.ComandoAceptar.Visible = False
        Me.ComandoRechazar.Visible = False
        Exit Function
    End If
    Me.ComandoAceptar.Visible = True
    Me.ComandoRechazar.Visible = True
    
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    
    m_RiesgoEnRetirada = RiesgoEnRetirada(m_EstadoRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RiesgoEnRetirada = EnumSiNo.No Then
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_EstadoRiesgo, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_RiesgoEnAceptacion = EnumSiNo.No Then
            p_Error = "El riesgo no está ni en estado de aceptación ni de retirada"
            Err.Raise 1000
        End If
    End If
    With m_ObjRiesgoActivo
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            
            If m_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinJustificar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoAceptar.Enabled = False
                Me.ComandoRechazar.Enabled = False
                
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.AceptadoSinVisar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                    Me.ComandoRechazar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.Aceptado Then
                Me.ComandoAceptar.Caption = "Quitar Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoRechazar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.AceptadoRechazado Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Quitar Rechazar"
                Me.ComandoAceptar.Enabled = False
                If EsTecnico <> EnumSiNo.Sí And blnRiesgoActivo = True Then
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoRechazar.Enabled = False
                End If
            End If
        End If
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatosCalidad ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Sub EstablecerControlesParaMitigacionAceptar()
    Dim lst As ListBox
    Dim m_Mitigacion As Variant
    Dim i As Long
    On Error Resume Next
    HacerVisiblesCamposOcultos
    RellenarCamposOcultos m_ObjRiesgoAlInicio
    If blnPermitidoEditar = True Then
        If m_ObjRiesgoAlInicio.EstadoEnum = EnumRiesgoEstado.Aceptado Then
            Me.JustificacionAceptacionRiesgo.Enabled = False
            Me.ComandoAceptacionRegistrar.Enabled = False
        Else
            Me.JustificacionAceptacionRiesgo.Enabled = True
            Me.ComandoAceptacionRegistrar.Enabled = True
        End If
    Else
        Me.JustificacionAceptacionRiesgo.Enabled = False
        Me.ComandoAceptacionRegistrar.Enabled = False
    End If
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = "Aceptar" Then
            lst.Selected(i) = True
            Me.DescripcionMitigacion = lst.Column(1, i)
            Exit For
        End If
    Next
End Sub


Private Function QuitarSeleccion(Optional ByRef p_Error As String) As String
    
    Dim i As Long
    On Error Resume Next
    
    
    For i = 1 To Me.ListaMitigacion.ListCount - 1
        Me.ListaMitigacion.Selected(i) = False
    Next
    
    Me.DescripcionMitigacion = Null
End Function
Private Sub HacerVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionAceptacionRiesgo.Visible = True
    Me.lblFechaJustificacionAceptacionRiesgo.Visible = True
    Me.ComandoAceptacionRegistrar.Visible = True
    Me.JustificacionAceptacionRiesgo.Visible = True
    Me.lblFechaAprobacionAceptacionPorCalidad.Visible = True
    Me.lblFechaJustificacionAceptacionRiesgo.Visible = True
    Me.lblFechaRechazoAceptacionPorCalidad.Visible = True
    Me.FechaJustificacionAceptacionRiesgo.Visible = True
    Me.FechaRechazoAceptacionPorCalidad.Visible = True
    Me.FechaAprobacionAceptacionPorCalidad.Visible = True
End Sub
Private Sub HacerInVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionAceptacionRiesgo.Visible = False
    Me.lblFechaJustificacionAceptacionRiesgo.Visible = False
    Me.ComandoAceptacionRegistrar.Visible = False
    Me.JustificacionAceptacionRiesgo.Visible = False
    Me.lblFechaAprobacionAceptacionPorCalidad.Visible = False
    Me.lblFechaJustificacionAceptacionRiesgo.Visible = False
    Me.lblFechaRechazoAceptacionPorCalidad.Visible = False
    Me.FechaJustificacionAceptacionRiesgo.Visible = False
    Me.FechaRechazoAceptacionPorCalidad.Visible = False
    Me.FechaAprobacionAceptacionPorCalidad.Visible = False
End Sub
Private Sub RellenarCamposOcultos(Optional p_Riesgo As riesgo)

    On Error Resume Next
    With p_Riesgo
        If .JustificacionAceptacionRiesgo <> "" Then
            Me.JustificacionAceptacionRiesgo = .JustificacionAceptacionRiesgo
        End If
        If IsDate(.FechaJustificacionAceptacionRiesgo) Then
            Me.FechaJustificacionAceptacionRiesgo = .FechaJustificacionAceptacionRiesgo
        End If
        If IsDate(.FechaAprobacionAceptacionPorCalidad) Then
            Me.FechaAprobacionAceptacionPorCalidad = .FechaAprobacionAceptacionPorCalidad
        End If
        If IsDate(.FechaRechazoAceptacionPorCalidad) Then
            Me.FechaRechazoAceptacionPorCalidad = .FechaRechazoAceptacionPorCalidad
        End If
    End With
    
End Sub

Private Sub ListaMitigacion_Click()
    
    Dim m_DescripcionSeleccionada As String
    Dim m_MitigacionSeleccionada As String
    Dim m_RiesgoEnRetirada As EnumSiNo
    
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_DescripcionSeleccionada = Nz(Me.ListaMitigacion.Column(1), "")
    m_MitigacionSeleccionada = Nz(Me.ListaMitigacion.Column(0), "")
    
   
    
    
    
    If Not m_ObjRiesgoAlInicio Is Nothing Then
        If m_ObjRiesgoAlInicio.Mitigacion = "Aceptar" And m_MitigacionSeleccionada <> "Aceptar" Then
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            pregunta = MsgBox("¿Desea quitar la aceptación del riesgo y elegir otra mitigación?", vbExclamation + vbYesNo + vbDefaultButton2, "Quitar mitigación aceptar")
            If pregunta <> vbYes Then
                Me.Parent.ComandoSalir.SetFocus
                EstablecerControlesParaMitigacionAceptar
                VBA.DoEvents
                DoCmd.Hourglass True
                VBA.DoEvents
                Exit Sub
            End If
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            m_ObjRiesgoActivo.AceptacionRegistrarQuitar m_MitigacionSeleccionada, Date, m_ObjRiesgoActivo.EstadoEnum, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            Me.DescripcionMitigacion = Me.ListaMitigacion.Column(1)
            EstablecerControlCombo Me.ListaMitigacion, EnumSiNo.Sí
            HacerInVisiblesCamposOcultos
            ActualizarEtiquetaEstadoRiesgo
            ' Recargar y re-seleccionar
            Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
            If m_Error <> "" Then Err.Raise 1000
            
            Form_FormRiesgosGestion.SeleccionarNodo m_ObjRiesgoActivo, m_Error
            If m_Error <> "" Then Err.Raise 1000
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
        If m_MitigacionSeleccionada = "Aceptar" Then
            If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Incompleto Then
                m_Error = "Se han de grabar primero los datos antes de aceptar el riesgo"
                Err.Raise 1000
            End If
            m_RiesgoEnRetirada = RiesgoEnRetirada(m_ObjRiesgoActivo.EstadoEnum)
            If m_RiesgoEnRetirada = EnumSiNo.Sí Then
                m_Error = "El riesgo está en fase de ser retirado"
                Err.Raise 1000
            End If
            If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
                m_Error = "El riesgo está materializado"
                Err.Raise 1000
            End If
            Me.DescripcionMitigacion = m_DescripcionSeleccionada
            m_ObjRiesgoActivo.Mitigacion = m_MitigacionSeleccionada
            HacerVisiblesCamposOcultos
            Me.ComandoAceptacionRegistrar.Enabled = True
            Me.JustificacionAceptacionRiesgo.Enabled = True
        Else
             Me.DescripcionMitigacion = m_DescripcionSeleccionada
           
        End If
    Else
        If m_MitigacionSeleccionada = "Aceptar" Then
            If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Incompleto Then
                m_Error = "Se han de grabar primero los datos antes de aceptar el riesgo"
                Err.Raise 1000
            End If
            m_RiesgoEnRetirada = RiesgoEnRetirada(m_ObjRiesgoActivo.EstadoEnum)
            If m_RiesgoEnRetirada = EnumSiNo.Sí Then
                m_Error = "El riesgo está en fase de ser retirado"
                Err.Raise 1000
            End If
            If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
                m_Error = "El riesgo está materializado"
                Err.Raise 1000
            End If
            Me.DescripcionMitigacion = m_DescripcionSeleccionada
            
            HacerVisiblesCamposOcultos
            Me.ComandoAceptacionRegistrar.Enabled = True
            Me.JustificacionAceptacionRiesgo.Enabled = True
        Else
             Me.DescripcionMitigacion = m_DescripcionSeleccionada
             m_ObjRiesgoActivo.JustificacionAceptacionRiesgo = ""
             m_ObjRiesgoActivo.FechaJustificacionAceptacionRiesgo = ""
             Me.JustificacionAceptacionRiesgo = Null
             Me.FechaJustificacionAceptacionRiesgo = Null
            HacerInVisiblesCamposOcultos
            
        End If
        
        
    End If
    
    EstablecerControlCombo Me.ListaMitigacion, EnumSiNo.Sí
   
    EstablecerColorPestañasRiesgo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaMitigacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub








