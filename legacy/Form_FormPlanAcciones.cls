Option Compare Database
Option Explicit

Private m_EsMitigacion As String
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_NodoPlan As MSComctlLib.Node
Public Event AccionNueva(m_PMAccion As Object)
Private Const ALTO_DESCRIPCION_SIN_CRAIZ As Long = 2712
Private Const ALTO_DESCRIPCION_CON_CRAIZ As Long = 1017
Private m_Error As String



Private Sub Accion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

Private Sub Accion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
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


Private Sub ComandoGrabar_Click()
    Dim m_Plan As Object
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    If m_EsMitigacion = "Sí" Then
        Set m_ObjPMAccionActiva = New PMAccion
        With m_ObjPMAccionActiva
            .IDMitigacion = m_ObjPMActivo.IDMitigacion
            .Accion = Nz(Me.Accion, "")
            .ResponsableAccion = Nz(Me.ResponsableAccion, "")
            .FechaInicio = Nz(Me.FechaInicio, "")
            .FechaFinPrevista = Nz(Me.FechaFinPrevista, "")
            .FechaFinReal = Nz(Me.FechaFinReal, "")
            .Registrar p_Error:=m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            RaiseEvent AccionNueva(m_ObjPMAccionActiva)
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Set m_ObjRiesgoActivo.ColPMs = Nothing
            ' Recargar y re-seleccionar
            Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
            If m_Error <> "" Then Err.Raise 1000
            
            Form_FormRiesgosGestion.SeleccionarNodo m_ObjPMAccionActiva, m_Error
            If m_Error <> "" Then Err.Raise 1000
            pregunta = MsgBox("¿Desea dar de alta otra acción", vbExclamation + vbYesNo + vbDefaultButton1, "Nueva acción")
            If pregunta <> vbYes Then
                DoCmd.Close acForm, Me.Name, acSaveNo
                Exit Sub
            End If
            Set m_ObjPMAccionActiva = Nothing
            Me.Accion = Null
            Me.ResponsableAccion = Null
            Me.FechaInicio = Null
            Me.FechaFinPrevista = Null
            Me.FechaFinReal = Null
            
            Me.Accion.SetFocus
            Exit Sub
            
        End With
        
        
    Else
        Set m_ObjPCAccionActiva = New PCAccion
        With m_ObjPCAccionActiva
            
            .IDContingencia = m_ObjPCActivo.IDContingencia
            .Accion = Nz(Me.Accion, "")
            .ResponsableAccion = Nz(Me.ResponsableAccion, "")
            .FechaInicio = Nz(Me.FechaInicio, "")
            .FechaFinPrevista = Nz(Me.FechaFinPrevista, "")
            .FechaFinReal = Nz(Me.FechaFinReal, "")
            .Registrar p_Error:=m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            RaiseEvent AccionNueva(m_ObjPCAccionActiva)
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Set m_ObjRiesgoActivo.ColPCs = Nothing
            ' Recargar y re-seleccionar
            Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.Sí, p_Error:=m_Error
            If m_Error <> "" Then Err.Raise 1000
            
            Form_FormRiesgosGestion.SeleccionarNodo m_ObjPCAccionActiva, m_Error
            If m_Error <> "" Then Err.Raise 1000
            pregunta = MsgBox("¿Desea dar de alta otra acción", vbExclamation + vbYesNo + vbDefaultButton1, "Nueva acción")
            If pregunta <> vbYes Then
                DoCmd.Close acForm, Me.Name, acSaveNo
                Exit Sub
            End If
            Set m_ObjPCAccionActiva = Nothing
            Me.Accion = Null
            Me.ResponsableAccion = Null
            Me.FechaInicio = Null
            Me.FechaFinPrevista = Null
            Me.FechaFinReal = Null
            
            Me.Accion.SetFocus
            Exit Sub
            
        End With
    End If
    
    
    
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
Private Sub ComandoSalir_Click()
    
    
    DoCmd.Close acForm, Me.Name, acSaveNo
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

Private Sub FechaFinReal_Change()
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaFinReal_Exit(Cancel As Integer)
    EstablecerColorFechaFinReal Me
End Sub

Private Sub FechaInicio_BeforeUpdate(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub

Private Sub FechaInicio_Change()
    EstablecerColorFechaInicio Me
End Sub

Private Sub FechaInicio_Exit(Cancel As Integer)
    EstablecerColorFechaInicio Me
End Sub

Private Sub Form_Load()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    m_EsMitigacion = Nz(Me.openArgs(), "")
    
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim m_DescRiesgo As String
    Dim m_Riesgo As riesgo
    Dim m_CausaRaiz As String
    Dim m_Titulo As String
    Dim m_RowSource As String
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion
    Dim m_EnumTipoNodo As EnumTipoNodo
    On Error GoTo errores
    
    p_Error = ""
    
    Me.ComandoGrabar.Enabled = True
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    If m_EsMitigacion = "Sí" Then
        If m_ObjPMActivo Is Nothing Then
            p_Error = "No se conoce el plan de mitigaci¿n"
            Err.Raise 1000
        End If
        Set m_Edicion = m_ObjPMActivo.riesgo.Edicion
        Set m_Riesgo = m_ObjPMActivo.riesgo
        m_Titulo = "ALTA acción PARA EL PLAN " & m_ObjPMActivo.CodMitigacion
        Me.lblDisparador.Caption = "Denominaci¿n"
        Me.DisparadorDelPlan = m_ObjPMActivo.DisparadorDelPlan
    Else
        If m_ObjPCActivo Is Nothing Then
            p_Error = "No se conoce el plan de contingencia"
            Err.Raise 1000
        End If
        Set m_Edicion = m_ObjPCActivo.riesgo.Edicion
        Set m_Riesgo = m_ObjPCActivo.riesgo
        m_Titulo = "ALTA acción PARA EL PLAN " & m_ObjPCActivo.CodContingencia
        Me.lblDisparador.Caption = "Disparador"
        Me.DisparadorDelPlan = m_ObjPCActivo.DisparadorDelPlan
    End If
    If m_Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        m_DescRiesgo = m_Riesgo.Descripcion
        m_CausaRaiz = m_Riesgo.CausaRaiz
    Else
        m_DescRiesgo = m_Riesgo.Descripcion
    End If
    Me.lblTitulo1.Caption = m_Titulo
    If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = True
        End If
        If FormularioAbierto("FormRiesgosGestion") Then
            Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
            m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
            m_EnumTipoNodo = getTipoNodo(m_NodoSeleccionado)
            If m_EnumTipoNodo = EnumTipoNodo.PM Then
                Set m_NodoPlan = m_NodoSeleccionado
            ElseIf m_EnumTipoNodo = EnumTipoNodo.PC Then
                Set m_NodoPlan = m_NodoSeleccionado
            ElseIf m_EnumTipoNodo = EnumTipoNodo.PMA Then
                Set m_NodoPlan = m_NodoSeleccionado.Parent
            ElseIf m_EnumTipoNodo = EnumTipoNodo.PCA Then
                Set m_NodoPlan = m_NodoSeleccionado.Parent
            End If
        End If
    End If
    
    If blnPermitidoEditar = True Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    
    If m_Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        Me.lblCRaiz.Visible = False
        Me.CausaRaiz.Visible = False
        Me.Descripcion.Height = ALTO_DESCRIPCION_SIN_CRAIZ
        Me.Descripcion.FontSize = 11
    Else
        Me.lblCRaiz.Visible = True
        Me.CausaRaiz.Visible = True
        Me.Descripcion.Height = ALTO_DESCRIPCION_CON_CRAIZ
        If m_Riesgo.CausaRaiz <> "" Then
            Me.CausaRaiz = m_Riesgo.CausaRaiz
        End If
        Me.Descripcion.FontSize = 9
        Me.CausaRaiz.FontSize = 9
    End If
    
    Me.Descripcion = m_DescRiesgo
    
    m_RowSource = m_Riesgo.Edicion.Proyecto.CadenaNombreAutorizados
    m_Error = m_Riesgo.Edicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.ResponsableAccion.RowSource = m_RowSource
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Sub ResponsableAccion_BeforeUpdate(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub

Private Sub ResponsableAccion_Exit(Cancel As Integer)
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub




