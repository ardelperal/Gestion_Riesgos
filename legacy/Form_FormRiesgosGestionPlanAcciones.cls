VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgosGestionPlanAcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private blnPermitidoEditar As Boolean
Private m_ObjPMAccionAlInicio As PMAccion
Private m_ObjPCAccionAlInicio As PCAccion
Private m_EsMitigacion As String
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_NodoPlan As MSComctlLib.Node
Private m_NodoAccion As MSComctlLib.Node
Private m_Riesgo As riesgo
Private m_Error As String
Private Sub ComandoEliminar_Click()
    Dim m_Plan As Object
    
    Dim m_NombreIcono As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMAccionActiva Is Nothing Then
            m_Error = "No se sabe la acción plan de mitigaci¿n"
            Err.Raise 1000
        End If
        Set m_Plan = m_ObjPMAccionActiva.Mitigacion
    Else
        If m_ObjPCAccionActiva Is Nothing Then
            m_Error = "No se sabe la acción del plan de contingencia"
            Err.Raise 1000
        End If
        Set m_Plan = m_ObjPCAccionActiva.Contingencia
    End If
    
    If blnPermitidoEditar = False Then
        m_Error = "Usuario no autorizado a esa operaci¿n"
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente borrar la acción seleccionada?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de una acción")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_EsMitigacion = "Sí" Then
        m_ObjPMAccionActiva.Borrar m_Error
        
    Else
        m_ObjPCAccionActiva.Borrar m_Error
        
    End If
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ' Recargar desde caché y re-seleccionar el plan
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Form_FormRiesgosGestion.SeleccionarNodo m_Plan, m_Error
    If m_Error <> "" Then Err.Raise 1000
    Form_FormRiesgosGestion.m_Arbol.Nodes.Remove m_NodoSeleccionado.Key
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("acción borrada con ¿xito", vbInformation, "acción borrada")
    
    
    
    Form_FormRiesgosGestion.Arbol_NodeClick m_NodoPlan
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
Private Sub ComandoGrabar_Click()
    Dim m_HaHabidoCambios As Boolean
    Dim m_Plan As Object
    Dim m_NombreIcono As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If blnPermitidoEditar = False Then
        Exit Sub
    End If
    If m_EsMitigacion = "Sí" Then
        With m_ObjPMAccionActiva
            
            .Accion = Nz(Me.Accion, "")
            .ResponsableAccion = Nz(Me.ResponsableAccion, "")
            .FechaInicio = Nz(Me.FechaInicio, "")
            .FechaFinPrevista = Nz(Me.FechaFinPrevista, "")
            .FechaFinReal = Nz(Me.FechaFinReal, "")
        End With
        Set m_Plan = m_ObjPMAccionActiva.Mitigacion
    Else
        With m_ObjPCAccionActiva
            
            .Accion = Nz(Me.Accion, "")
            .ResponsableAccion = Nz(Me.ResponsableAccion, "")
            .FechaInicio = Nz(Me.FechaInicio, "")
            .FechaFinPrevista = Nz(Me.FechaFinPrevista, "")
            .FechaFinReal = Nz(Me.FechaFinReal, "")
        End With
        Set m_Plan = m_ObjPCAccionActiva.Contingencia
    End If
    
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No hay nada que grabar"
        Err.Raise 1000
    End If
    If m_EsMitigacion = "Sí" Then
        m_ObjPMAccionActiva.Registrar p_ObjPMAccionAlInicio:=m_ObjPMAccionAlInicio, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjPMAccionAlInicio = Constructor.getPMAccion(m_ObjPMAccionAlInicio.IDAccionMitigacion)
        Set m_ObjRiesgoActivo.ColPMs = Nothing
    Else
        m_ObjPCAccionActiva.Registrar p_ObjPCAccionAlInicio:=m_ObjPCAccionAlInicio, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjPCAccionAlInicio = Constructor.getPCAccion(m_ObjPCAccionAlInicio.IDAccionContingencia)
        Set m_ObjRiesgoActivo.ColPCs = Nothing
    End If
    
    
    EstablecerDatos
    
    ' Recargar y re-seleccionar la acción
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    If m_EsMitigacion = "Sí" Then
        Form_FormRiesgosGestion.SeleccionarNodo m_ObjPMAccionActiva, m_Error
    Else
        Form_FormRiesgosGestion.SeleccionarNodo m_ObjPCAccionActiva, m_Error
    End If
    If m_Error <> "" Then Err.Raise 1000
    
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
Private Sub FechaFinReal_Change()
    EstablecerColorFechaFinReal Me
End Sub
Private Sub FechaInicio_Change()
    EstablecerColorFechaInicio Me
End Sub
Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Form_FormRiesgosGestion.m_EsMitigacion = Empty Then
        m_Error = "No se sabe si es plan de mitigaci¿n o de contingencia"
        Err.Raise 1000
    End If
    If Form_FormRiesgosGestion.m_EsMitigacion = EnumSiNo.Sí Then
        m_EsMitigacion = "Sí"
    Else
        m_EsMitigacion = "No"
    End If
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMAccionActiva Is Nothing Then
            m_Error = "No se conoce el plan de mitigaci¿n de la acción"
            Err.Raise 1000
        End If
        Set m_ObjPMAccionAlInicio = getPMAccion(m_ObjPMAccionActiva.IDAccionMitigacion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjPMAccionAlInicio Is Nothing Then
            m_Error = "Es una edici¿n y no se han podido cargar los datos"
            Err.Raise 1000
        End If
        
    Else
        If m_ObjPCAccionActiva Is Nothing Then
            m_Error = "No se conoce el plan de contingencia de la acción"
            Err.Raise 1000
        End If
        Set m_ObjPCAccionAlInicio = getPCAccion(m_ObjPCAccionActiva.IDAccionContingencia, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjPCAccionAlInicio Is Nothing Then
            m_Error = "Es una edici¿n y no se han podido cargar los datos"
            Err.Raise 1000
        End If
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If blnPermitidoEditar = True Then
        Me.ComandoGrabar.Enabled = True
        Me.ComandoEliminar.Enabled = True
        
        Me.AllowEdits = True
    Else
        Me.ComandoGrabar.Enabled = False
        Me.ComandoEliminar.Enabled = False
        Me.AllowEdits = False
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_RowSource As String
    
    On Error GoTo errores
    
    p_Error = ""
    Me.ComandoEliminar.Enabled = False
    Me.ComandoGrabar.Enabled = False
    Me.FechaFinPrevista.ForeColor = m_ColorOK
    m_RowSource = m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados
    
    If FormularioAbierto("FormRiesgosGestion") Then
        Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
        m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
    End If
    If m_NodoSeleccionado Is Nothing Then
        p_Error = "No se conoce el nodo seleccionado"
        Err.Raise 1000
    End If
    If TypeOf Form_FormRiesgosGestion.m_ObjSeleccionado Is PM Or _
        TypeOf Form_FormRiesgosGestion.m_ObjSeleccionado Is PC Then
        Set m_NodoPlan = m_NodoSeleccionado
        Set m_NodoAccion = Nothing
    ElseIf TypeOf Form_FormRiesgosGestion.m_ObjSeleccionado Is PMAccion Or _
        TypeOf Form_FormRiesgosGestion.m_ObjSeleccionado Is PCAccion Then
        Set m_NodoAccion = m_NodoSeleccionado
        Set m_NodoPlan = m_NodoAccion.Parent
    Else
        p_Error = "Nodo seleccionado desconocido"
        Err.Raise 1000
    End If
    
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_ObjEdicionActiva.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = True
        End If
    End If
    
    Me.ResponsableAccion.RowSource = m_RowSource
    If m_EsMitigacion = "Sí" Then
        With m_ObjPMAccionAlInicio
            Me.lblTitulo.Caption = "Accci¿n " & .CodAccion & " " & " " & "plan de mitigaci¿n " & " " & .Mitigacion.CodMitigacion
            Me.Accion = .Accion
            If .ResponsableAccion <> "" Then
                Me.ResponsableAccion = .ResponsableAccion
            End If
            If IsDate(.FechaInicio) Then
                Me.FechaInicio = .FechaInicio
            End If
            If IsDate(.FechaFinPrevista) Then
                Me.FechaFinPrevista = .FechaFinPrevista
            End If
            If m_ObjPMAccionAlInicio.EsActivo = No Then
                Me.FechaFinPrevista.ForeColor = m_ColorOK
            Else
                If IsDate(.FechaFinPrevista) Then
                    Me.FechaFinPrevista = .FechaFinPrevista
                    If .NecesitaReplanificacion = EnumSiNo.Sí Then
                        Me.FechaFinPrevista.ForeColor = m_ColorNOOK
                    Else
                        Me.FechaFinPrevista.ForeColor = m_ColorOK
                    End If
                End If
            End If
            
            If IsDate(.FechaFinReal) Then
                Me.FechaFinReal = .FechaFinReal
            End If
            Set m_Riesgo = .Mitigacion.riesgo
        End With
        
    Else
        With m_ObjPCAccionAlInicio
            Me.lblTitulo.Caption = "Accci¿n " & .CodAccion & " " & " " & "plan de contingencia " & " " & .Contingencia.CodContingencia
            Me.Accion = .Accion
            If .ResponsableAccion <> "" Then
                Me.ResponsableAccion = .ResponsableAccion
            End If
            If IsDate(.FechaInicio) Then
                Me.FechaInicio = .FechaInicio
            End If
            If IsDate(.FechaFinPrevista) Then
                Me.FechaFinPrevista = .FechaFinPrevista
                If .NecesitaReplanificacion = EnumSiNo.Sí Then
                    Me.FechaFinPrevista.ForeColor = m_ColorNOOK
                Else
                    Me.FechaFinPrevista.ForeColor = m_ColorOK
                End If
            End If
            If IsDate(.FechaFinReal) Then
                Me.FechaFinReal = .FechaFinReal
            End If
            Set m_Riesgo = .Contingencia.riesgo
        End With
    End If
    
    RellenarBordeCamposObligatorios Me
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    
    
    On Error GoTo errores
    
    If m_EsMitigacion = "Sí" Then
        
        With m_ObjPMAccionAlInicio
            If .Accion <> Nz(Me.Accion, "") Then
                HaHabidoCambios = True
                Exit Function
            End If
            If .ResponsableAccion <> Nz(Me.ResponsableAccion, "") Then
                HaHabidoCambios = True
                Exit Function
            End If
            If IsDate(.FechaInicio) Then
                If Not IsDate(Nz(Me.FechaInicio, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaInicio, "dd/mm/yyyy") <> Format(Me.FechaInicio, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaInicio, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
            If IsDate(.FechaFinPrevista) Then
                If Not IsDate(Nz(Me.FechaFinPrevista, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaFinPrevista, "dd/mm/yyyy") <> Format(Me.FechaFinPrevista, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaFinPrevista, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
            If IsDate(.FechaFinReal) Then
                If Not IsDate(Nz(Me.FechaFinReal, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaFinReal, "dd/mm/yyyy") <> Format(Me.FechaFinReal, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaFinReal, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
        End With
    Else
        With m_ObjPCAccionAlInicio
            If .Accion <> Nz(Me.Accion, "") Then
                HaHabidoCambios = True
                Exit Function
            End If
            If .ResponsableAccion <> Nz(Me.ResponsableAccion, "") Then
                HaHabidoCambios = True
                Exit Function
            End If
            If IsDate(.FechaInicio) Then
                If Not IsDate(Nz(Me.FechaInicio, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaInicio, "dd/mm/yyyy") <> Format(Me.FechaInicio, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaInicio, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
            If IsDate(.FechaFinPrevista) Then
                If Not IsDate(Nz(Me.FechaFinPrevista, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaFinPrevista, "dd/mm/yyyy") <> Format(Me.FechaFinPrevista, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaFinPrevista, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
            If IsDate(.FechaFinReal) Then
                If Not IsDate(Nz(Me.FechaFinReal, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
                If Format(.FechaFinReal, "dd/mm/yyyy") <> Format(Me.FechaFinReal, "dd/mm/yyyy") Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            Else
                If IsDate(Nz(Me.FechaFinReal, "")) Then
                    HaHabidoCambios = True
                    Exit Function
                End If
            End If
        End With
    End If
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Sub Accion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub
Private Sub Accion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
End Sub
Private Sub FechaFinPrevista_BeforeUpdate(Cancel As Integer)
    
   EstablecerColorFechaFinPrevista Me
End Sub
Private Sub FechaFinPrevista_Change()
    EstablecerColorFechaFinPrevista Me
End Sub
Private Sub FechaFinPrevista_Exit(Cancel As Integer)
    
    EstablecerColorFechaFinPrevista Me
End Sub
Private Sub FechaFinReal_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub
Private Sub FechaFinReal_Exit(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub
Private Sub FechaInicio_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub
Private Sub FechaInicio_Exit(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub


