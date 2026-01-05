
Option Compare Database
Option Explicit



Public m_ObjColEmpresas As Scripting.Dictionary
Public m_ObjColNombres As Scripting.Dictionary

Public Event RiesgoNuevo()
Public Event RiesgoEditado()

Public Event cambioEnPriorizacion(p_Priorizacion As String)
Public m_PermitidoEditar As EnumSiNo
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_Error As String



Private Sub ComandoActualizar_Click()
    Dim m_Form As Form
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    Set m_ObjRiesgoActivo = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Form = Me.FrmDetalle.Form
    With Form
        If .Name = "FormRiesgoDefinicion" Then
            Form_FormRiesgoDefinicion.EstablecerDatos
        ElseIf .Name = "FormRiesgoPlazoCosteCalidad" Then
            Form_FormRiesgoPlazoCosteCalidad.EstablecerDatos
        ElseIf .Name = "FormRiesgoVulnerabilidad" Then
            Form_FormRiesgoVulnerabilidad.EstablecerDatos
        ElseIf .Name = "FormRiesgoMitigacion" Then
            Form_FormRiesgoMitigacion.EstablecerDatos
        ElseIf .Name = "FormRiesgoRetirado" Then
            Form_FormRiesgoRetirado.EstablecerDatos
        ElseIf .Name = "FormRiesgoMaterializado" Then
            Form_FormRiesgoMaterializado.EstablecerDatos
        End If
    End With
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

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


Private Sub ComandoSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoVerEdicion_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    'Set m_ObjEdicionActiva = m_ObjRiesgoActivo.Edicion
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
        m_Error = "Al ComandoVerEdicion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    
    Me.InsideHeight = 9060
    Me.InsideWidth = 13980
    
    
    EstablecerDatos m_Error
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
Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_Titulo As String
    Dim m_Particula As String
    Dim m_EstadoRiesgo As EnumRiesgoEstado
    
    
    On Error GoTo errores
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    
    If Not m_ObjRiesgoActivo Is Nothing Then
        Set m_ObjEdicionActiva = m_ObjRiesgoActivo.Edicion
    Else
        If m_ObjEdicionActiva Is Nothing Then
            If m_ObjRiesgoActivo Is Nothing Then
                m_Error = "No se conoce ni el riesgo ni la edición"
                Err.Raise 1000
            End If
            Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjRiesgoActivo.IDEdicion)
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se puede determinar la edición activa"
                Err.Raise 1000
            End If
        End If
    End If
    
    If m_ObjEdicionActiva Is Nothing Then
        If m_ObjRiesgoActivo Is Nothing Then
            m_Error = "No se conoce ni el riesgo ni la edición"
            Err.Raise 1000
        End If
        Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjRiesgoActivo.IDEdicion)
        If m_ObjEdicionActiva Is Nothing Then
            p_Error = "No se puede determinar la edición activa"
            Err.Raise 1000
        End If
    End If
    If m_PermitidoEditar = Empty Then
        With m_ObjEdicionActiva
            If .EsActivo = EnumSiNo.No Then
                m_PermitidoEditar = EnumSiNo.No
            Else
                If .Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                    m_PermitidoEditar = EnumSiNo.Sí
                
                Else
                    m_PermitidoEditar = EnumSiNo.No
                End If
            End If
        End With
    
    End If
    If m_PermitidoEditar = EnumSiNo.Sí Then
        Me.AllowEdits = True
        Me.ComandoGrabar.Enabled = True
    Else
        Me.AllowEdits = False
        Me.ComandoGrabar.Enabled = False
    End If
    'Debug.Print Me.lblEstadoRiesgo.ForeColor
    
    If m_EsAlta = EnumSiNo.Sí Then
        
        If m_PermitidoEditar = EnumSiNo.No Then
            m_Error = "Es un alta y el usuario no tiene permitida esa operaciï¿½n"
            Err.Raise 1000
        End If
        Set m_ObjRiesgoActivo = New riesgo
        m_ObjRiesgoActivo.IDEdicion = m_ObjEdicionActiva.IDEdicion
        Set m_ObjRiesgoAlInicio = Nothing
        Me.NavGeneral.Enabled = True
        Me.lblEstadoRiesgo.Visible = False
    Else
        If m_ObjRiesgoActivo Is Nothing Then
            m_Error = "Es una edición y no hay datos"
            Err.Raise 1000
        End If
        Me.lblEstadoRiesgo.Visible = True
        Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo, , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjRiesgoAlInicio Is Nothing Then
            m_Error = "Es una edición y no hay datos"
            Err.Raise 1000
        End If

    End If
    m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum
    
    If m_EstadoRiesgo = EnumRiesgoEstado.Incompleto Or m_EstadoRiesgo = Empty Then
        Me.NavRetirado.Enabled = False
        Me.NavMaterializado.Enabled = False
    Else
        Me.NavRetirado.Enabled = True
        Me.NavMaterializado.Enabled = True
    End If
    With m_ObjEdicionActiva
        If .EsActivo = EnumSiNo.Sí Then
            blnEdicionActiva = True
            m_Particula = "(edición activa)"
        Else
            blnEdicionActiva = False
            m_Particula = "(edición no Activa)"
        End If
        If m_EsAlta = EnumSiNo.Sí Then
            m_Titulo = "ALTA DE RIESGO PARA LA edición " & .Edicion & " " & m_Particula
        Else
            m_Titulo = "DETALLE DEL RIESGO: " & m_ObjRiesgoAlInicio.CodigoRiesgo & " edición " & .Edicion & " " & m_Particula
        End If
        
        Me.lblTitulo.Caption = m_Titulo
        If .Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            Me.NavGeneral.NavigationTargetName = "FormRiesgoDefinicion"
        Else
            Me.NavGeneral.NavigationTargetName = "FormRiesgoDefinicionNoBiblioteca"
        End If
    End With
    ActualizarEtiquetaEstadoRiesgo , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerColorPestañasRiesgo p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
        m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
    End If
    
    
    
    Me.FrmDetalle.SourceObject = Me.NavGeneral.NavigationTargetName
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

    
Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_EstadoRiesgo As String
    Dim m_NombreIcono As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    RellenarObjeto
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_HaHabidoCambios = False Then
        m_Error = "Ningï¿½n cambio que guardar"
        Err.Raise 1000
    End If
    
    With m_ObjRiesgoActivo
        '.Estado = m_ObjRiesgoActivo.EstadoCalculadoTexto
        If m_EsAlta = EnumSiNo.Sí Then
            .IDEdicion = m_ObjEdicionActiva.IDEdicion
        End If
        
        .Registrar EnumSiNo.Sí, m_ObjRiesgoAlInicio, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_EstadoRiesgo = .EstadoEnum
        ActualizarEtiquetaEstadoRiesgo
    End With
    If m_EstadoRiesgo = EnumRiesgoEstado.Incompleto Then
        Me.NavRetirado.Enabled = False
        Me.NavMaterializado.Enabled = False
    Else
        Me.NavRetirado.Enabled = True
        Me.NavMaterializado.Enabled = True
    End If
    
    If m_ObjRiesgoAlInicio Is Nothing Then
       RaiseEvent RiesgoNuevo
        
    Else

        ' Recargar y re-seleccionar el riesgo
          Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
          If m_Error <> "" Then Err.Raise 1000
          
          Form_FormRiesgosGestion.SeleccionarNodo m_ObjRiesgoActivo, m_Error
          If m_Error <> "" Then Err.Raise 1000
       
        
        If FormularioAbierto("FormRiesgosGestion") Then
            If Forms("FormRiesgosGestion").FormDetalle.SourceObject = "FormRiesgosGestionRiesgo" Then
                Form_FormRiesgosGestionRiesgo.EstablecerDatos
            End If
            
        End If
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
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



Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    
    If m_ObjRiesgoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    
    
    With m_ObjRiesgoAlInicio
        If .Priorizacion <> m_ObjRiesgoActivo.Priorizacion Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Descripcion <> m_ObjRiesgoActivo.Descripcion Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            If .CausaRaiz <> m_ObjRiesgoActivo.CausaRaiz Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
        If .EntidadDetecta <> m_ObjRiesgoActivo.EntidadDetecta Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DetectadoPor <> m_ObjRiesgoActivo.DetectadoPor Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaDetectado <> m_ObjRiesgoActivo.FechaDetectado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaMaterializado <> m_ObjRiesgoActivo.FechaMaterializado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaRetirado <> m_ObjRiesgoActivo.FechaRetirado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Plazo <> m_ObjRiesgoActivo.Plazo Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Coste <> m_ObjRiesgoActivo.Coste Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Calidad <> m_ObjRiesgoActivo.Calidad Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Vulnerabilidad <> m_ObjRiesgoActivo.Vulnerabilidad Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Mitigacion <> m_ObjRiesgoActivo.Mitigacion Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .JustificacionAceptacionRiesgo <> m_ObjRiesgoActivo.JustificacionAceptacionRiesgo Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .JustificacionRetiroRiesgo <> m_ObjRiesgoActivo.JustificacionRetiroRiesgo Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .CodRiesgoBiblioteca <> m_ObjRiesgoActivo.CodRiesgoBiblioteca Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub NavMaterializado_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   
End Sub

Private Sub NavMitigacion_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   

End Sub

Public Function RellenarObjeto(Optional ByRef p_Error As String) As String
    Dim m_Form As Form
    On Error GoTo errores
    Set m_Form = Me.FrmDetalle.Form
    With m_ObjRiesgoActivo
        If Me.FrmDetalle.SourceObject = "FormRiesgoDefinicion" Then
            .Origen = Nz(m_Form.Origen, "")
            .Priorizacion = Nz(m_Form.Priorizacion, "")
            .Descripcion = Nz(m_Form.Descripcion, "")
            .CausaRaiz = Nz(m_Form.CausaRaiz, "")
            .FechaDetectado = Nz(m_Form.FechaDetectado, "")
            .EntidadDetecta = Nz(m_Form.EntidadDetecta, "")
            .DetectadoPor = Nz(m_Form.DetectadoPor, "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoDefinicionNoBiblioteca" Then
            .Origen = Nz(m_Form.Origen, "")
            .Priorizacion = Nz(m_Form.Priorizacion, "")
            .Descripcion = Nz(m_Form.Descripcion, "")
            .FechaDetectado = Nz(m_Form.FechaDetectado, "")
            .EntidadDetecta = Nz(m_Form.EntidadDetecta, "")
            .DetectadoPor = Nz(m_Form.DetectadoPor, "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoPlazoCosteCalidad" Then
            .Plazo = Nz(m_Form.ListaPlazo.Column(0), "")
            .Calidad = Nz(m_Form.ListaCalidad.Column(0), "")
            .Coste = Nz(m_Form.ListaCoste.Column(0), "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoVulnerabilidad" Then
            .Vulnerabilidad = Nz(m_Form.ListaVulnerabilidad.Column(0), "")
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoMitigacion" Then
            .Mitigacion = Nz(m_Form.ListaMitigacion.Column(0), "")
            If m_Form.JustificacionAceptacionRiesgo.Visible Then
                .JustificacionAceptacionRiesgo = Nz(m_Form.JustificacionAceptacionRiesgo, "")
            Else
                .JustificacionAceptacionRiesgo = ""
            End If
            If m_Form.FechaJustificacionAceptacionRiesgo.Visible Then
                .FechaJustificacionAceptacionRiesgo = Nz(m_Form.FechaJustificacionAceptacionRiesgo, "")
            Else
                .FechaJustificacionAceptacionRiesgo = ""
            End If
        ElseIf Me.FrmDetalle.SourceObject = "FormRiesgoRetirado" Then
            .FechaRetirado = Nz(m_Form.FechaRetirado, "")
            .JustificacionRetiroRiesgo = Nz(m_Form.JustificacionRetiroRiesgo, "")
            .FechaJustificacionRetiroRiesgo = Nz(m_Form.FechaJustificacionRetiroRiesgo, "")
            .FechaRechazoRetiroPorCalidad = Nz(m_Form.FechaRechazoRetiroPorCalidad, "")
            
       
        End If
        
    End With
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha devuelto el error: " & Err.Description
    End If
End Function



Private Sub NavPlazoCosteCalidad_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
   
End Sub

Private Sub NavRetirado_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
    
End Sub

Private Sub NavVulnerabilidad_Enter()
    If Me.AllowEdits = True Then
        RellenarObjeto
    End If
    
End Sub



