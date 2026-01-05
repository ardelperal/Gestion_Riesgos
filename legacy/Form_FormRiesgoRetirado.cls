VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgoRetirado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private blnRiesgoActivo As Boolean
Private m_EstadoRiesgo As EnumRiesgoEstado
Private blnPermitidoEditar As Boolean
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_Error As String



Private Sub ComandoAceptar_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    With m_ObjRiesgoActivo
        
        If Me.ComandoAceptar.Caption = "Aceptar" Then
            .RetiroAprobar Date, m_EstadoRiesgo, m_Error
        Else
            .RetiroAprobacionQuitar m_EstadoRiesgo, m_Error
        End If
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    Set m_ObjRiesgoActivo = Constructor.getRiesgo(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
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
            .RetiroRechazo m_EstadoRiesgo, Date, m_Error
        Else
            .RetiroRechazoQuitar m_EstadoRiesgo, m_Error
        End If
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End With
    Set m_ObjRiesgoActivo = Constructor.getRiesgo(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
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

Private Sub ComandoRetiroRegistrar_Click()
    
    Dim m_NombreIcono As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
   
    If Nz(Me.JustificacionRetiroRiesgo, "") = "" Then
        m_Error = "Se ha de indicar alguna justificación"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.IDRiesgo = "" Then
        m_Error = "Primero ha de registrar el resto de campos del Riesgo"
        Err.Raise 1000
    End If
    If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
        pregunta = MsgBox("El riesgo actualmente estó materializado." & vbNewLine & _
                    "¿Desea proponer la retirada del riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Propuesta de retirada")
        If pregunta <> vbYes Then
            Exit Sub
        End If
        m_ObjRiesgoActivo.MaterializacionQuitarRegistrar m_ObjRiesgoActivo.EstadoEnum, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_ObjRiesgoActivo.RetiroRegistrar Me.JustificacionRetiroRiesgo, Date, m_ObjRiesgoActivo.EstadoEnum, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
    If IsDate(m_ObjRiesgoAlInicio.FechaJustificacionRetiroRiesgo) Then
        Me.FechaJustificacionRetiroRiesgo = m_ObjRiesgoAlInicio.FechaJustificacionRetiroRiesgo
        Me.ComandoEliminar.Visible = True
        Me.ComandoEliminar.Enabled = True
    End If
    ActualizarEtiquetaEstadoRiesgo
    ' Recargar y re-seleccionar
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Form_FormRiesgosGestion.SeleccionarNodo m_ObjRiesgoActivo, m_Error
    If m_Error <> "" Then Err.Raise 1000
    EstablecerDatosCalidad
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Ha sido registrado la propuesta de Retiro del Riesgo y ha salido un correo electrónico a Calidad para su aprobación", vbInformation, "Riesgo propuesto para Retirar")
    
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRetiroRegistrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_Estado As EnumRiesgoEstado
    On Error GoTo errores
    
    
    m_Error = ""
    
    pregunta = MsgBox("¿Desea quitar la retirada del Riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Quitar retirada")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjRiesgoActivo.IDRiesgo = "" Then
        m_Error = "Primero ha de registrar el resto de campos del Riesgo"
        Err.Raise 1000
    End If
    m_Estado = m_ObjRiesgoActivo.ESTADOCalculado
    m_RiesgoEnRetirada = RiesgoEnRetirada(m_Estado)
    If m_RiesgoEnRetirada = EnumSiNo.No Then
        m_Error = "El retiro del riesgo no es sin justificar, sin visar o rechazado"
        Err.Raise 1000
    End If
    
    
    
    m_ObjRiesgoActivo.RetiroRegistrarQuitar m_Estado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Parent.ComandoSalir.SetFocus
    Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(m_ObjRiesgoActivo.IDRiesgo)
    ActualizarEtiquetaEstadoRiesgo
    ' Recargar y re-seleccionar
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Form_FormRiesgosGestion.SeleccionarNodo m_ObjRiesgoActivo, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    EstablecerDatos
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("Ha sido registrado la habilitación del riesgo", vbInformation, "Riesgo habilitado")
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoFRetirado_Click()
    Dim m_Fecha As String
    Dim m_RiesgoEnAceptacion As EnumSiNo
    On Error GoTo errores
    m_Error = ""
    If blnPermitidoEditar = False Then
       Exit Sub
    End If
    m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_ObjRiesgoActivo.EstadoEnum)
    If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
        Me.FechaRetirado = Null
        m_Error = "El riesgo estó aceptado o en fase de aceptación"
        Err.Raise 1000
    End If
    m_Fecha = Nz(InputBox("Introduzca la fecha de retiro del riesgo", "Fecha retirado", Date), "")
    If Not IsDate(m_Fecha) Then
        Exit Sub
    End If
    
    If Not m_ObjRiesgoAlInicio Is Nothing Then
        If m_ObjRiesgoAlInicio.FechaRetirado = m_Fecha Then
            Exit Sub
        End If
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        If blnPermitidoEditar = True Then
            If m_ObjRiesgoAlInicio.EstadoEnum = EnumRiesgoEstado.Retirado Then
                Me.JustificacionRetiroRiesgo.Enabled = False
                Me.ComandoRetiroRegistrar.Enabled = False
                Me.ComandoEliminar.Enabled = True
                Me.ComandoEliminar.Visible = True
            Else
                Me.JustificacionRetiroRiesgo.Enabled = True
                Me.ComandoRetiroRegistrar.Enabled = True
                Me.ComandoEliminar.Visible = False
                
            End If
        Else
            Me.JustificacionRetiroRiesgo.Enabled = False
            Me.ComandoRetiroRegistrar.Enabled = False
            Me.ComandoEliminar.Enabled = False
            Me.ComandoEliminar.Visible = False
        End If
        Me.FechaRetirado = Date
        Exit Sub
    End If
    Me.FechaRetirado = Date
    HacerVisiblesCamposOcultos
    Me.JustificacionRetiroRiesgo.Enabled = True
    Me.ComandoRetiroRegistrar.Enabled = True
    Me.ComandoEliminar.Visible = False
    Me.ComandoEliminar.Enabled = False
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFRetirado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
'    On Error GoTo errores
    On Error Resume Next
    
    If m_ObjRiesgoAlInicio Is Nothing Then
        HacerInVisiblesCamposOcultos
        Me.AllowEdits = False
        Exit Function
    End If
    
    'Form_FormRiesgo.NavRetirado.SetFocus
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
        m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
    End If
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
    If m_ObjRiesgoActivo.EsActivo = EnumSiNo.Sí Then
        blnRiesgoActivo = True
    Else
        blnRiesgoActivo = False
    End If
    If blnPermitidoEditar = True Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Retirado Then
        Me.FechaRetirado = m_ObjRiesgoAlInicio.FechaRetirado
        Me.ComandoFRetirado.Enabled = False
        If blnPermitidoEditar = True Then
            Me.ComandoEliminar.Enabled = True
        End If
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        Exit Function
    End If
    
    If IsDate(m_ObjRiesgoAlInicio.FechaRetirado) Then
        Me.FechaRetirado = m_ObjRiesgoAlInicio.FechaRetirado
        HacerVisiblesCamposOcultos
        RellenarCamposOcultos m_ObjRiesgoAlInicio
        If blnPermitidoEditar = True Then
            Me.JustificacionRetiroRiesgo.Enabled = True
            Me.ComandoRetiroRegistrar.Enabled = True
            Me.ComandoEliminar.Enabled = True
        Else
            
            Me.ComandoFRetirado.Enabled = False
            Me.JustificacionRetiroRiesgo.Enabled = False
            Me.ComandoRetiroRegistrar.Enabled = False
            Me.ComandoEliminar.Enabled = False
            Me.ComandoEliminar.Visible = False
        End If
        Exit Function
    End If
    HacerInVisiblesCamposOcultos
    Me.FechaRetirado = Null
    

    
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
    If Not IsDate(m_ObjRiesgoAlInicio.FechaRetirado) Then
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
            p_Error = "El riesgo no estó ni en estado de aceptación ni de retirada"
            Err.Raise 1000
        End If
    End If
    With m_ObjRiesgoActivo
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            
           If m_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinJustificar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoAceptar.Enabled = False
                Me.ComandoRechazar.Enabled = False
                
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.RetiradoSinVisar Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                If blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                    Me.ComandoRechazar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                    Me.ComandoRechazar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.Retirado Then
                Me.ComandoAceptar.Caption = "Quitar Aceptar"
                Me.ComandoRechazar.Caption = "Rechazar"
                Me.ComandoRechazar.Enabled = False
                If blnRiesgoActivo = True Then
                    Me.ComandoAceptar.Enabled = True
                Else
                    Me.ComandoAceptar.Enabled = False
                End If
            ElseIf m_EstadoRiesgo = EnumRiesgoEstado.RetiradoRechazado Then
                Me.ComandoAceptar.Caption = "Aceptar"
                Me.ComandoRechazar.Caption = "Quitar Rechazar"
                Me.ComandoAceptar.Enabled = False
                If blnRiesgoActivo = True Then
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
Private Sub HacerVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionRetiroRiesgo.Visible = True
    Me.ComandoRetiroRegistrar.Visible = True
    Me.ComandoEliminar.Visible = True
    
    Me.JustificacionRetiroRiesgo.Visible = True
    Me.lblFechaJustificacionRetiroRiesgo.Visible = True
    Me.lblFechaAprobacionRetiroPorCalidad.Visible = True
    Me.lblFechaRechazoRetiroPorCalidad.Visible = True
    Me.FechaJustificacionRetiroRiesgo.Visible = True
    Me.FechaAprobacionRetiroPorCalidad.Visible = True
    Me.FechaRechazoRetiroPorCalidad.Visible = True
End Sub
Private Sub HacerInVisiblesCamposOcultos()

    On Error Resume Next
    Me.lblJustificacionRetiroRiesgo.Visible = False
    Me.ComandoRetiroRegistrar.Visible = False
   
    Me.ComandoEliminar.Visible = False
    
    Me.JustificacionRetiroRiesgo.Visible = False
    Me.lblFechaJustificacionRetiroRiesgo.Visible = False
    Me.lblFechaAprobacionRetiroPorCalidad.Visible = False
    Me.lblFechaRechazoRetiroPorCalidad.Visible = False
    Me.FechaJustificacionRetiroRiesgo.Visible = False
    Me.FechaAprobacionRetiroPorCalidad.Visible = False
    Me.FechaRechazoRetiroPorCalidad.Visible = False
End Sub

Private Sub RellenarCamposOcultos(Optional p_Riesgo As riesgo)

    On Error Resume Next
    With p_Riesgo
        If .JustificacionRetiroRiesgo <> "" Then
            Me.JustificacionRetiroRiesgo = .JustificacionRetiroRiesgo
        End If
        If IsDate(.FechaJustificacionRetiroRiesgo) Then
            Me.FechaJustificacionRetiroRiesgo = .FechaJustificacionRetiroRiesgo
        End If
        If IsDate(.FechaAprobacionRetiroPorCalidad) Then
            Me.FechaAprobacionRetiroPorCalidad = .FechaAprobacionRetiroPorCalidad
        End If
        If IsDate(.FechaRechazoRetiroPorCalidad) Then
            Me.FechaRechazoRetiroPorCalidad = .FechaRechazoRetiroPorCalidad
        End If
    End With
    
End Sub





