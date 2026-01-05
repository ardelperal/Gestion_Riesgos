Option Compare Database
Option Explicit

Public Event RiesgoAceptado()
Public Event RiesgoRetirado()
Public Event RiesgoMaterializado()

Private m_URLInformeRiesgo As String
Private WithEvents m_FormPlan As Form_FormPlanPrincipal


' Constantes de diseño (en Twips: 1cm = 567 twips)
' Ajustados para que no solapen con la nueva sección de Histórico
Private Const ALTO_DESCRIPCION_SIN_CRAIZ As Long = 2000 ' ~3.5cm (reducido para dar espacio)
Private Const ALTO_DESCRIPCION_CON_CRAIZ As Long = 1137 ' ~2.0cm

Private m_Error As String

' ==========================================
' EVENTOS DE BOTONES
' ==========================================

Private Sub ComandoAltaPC_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormPlanPrincipal") Then DoCmd.Close acForm, "FormPlanPrincipal", acSaveNo
    DoCmd.OpenForm "FormPlanPrincipal", , , , , , "No"
    If FormularioAbierto("FormPlanPrincipal") Then Set m_FormPlan = Forms("FormPlanPrincipal")
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoAltaPC_Click", Err
End Sub

Private Sub ComandoAltaPM_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormPlanPrincipal") Then DoCmd.Close acForm, "FormPlanPrincipal", acSaveNo
    DoCmd.OpenForm "FormPlanPrincipal", , , , , , "Sí"
    If FormularioAbierto("FormPlanPrincipal") Then Set m_FormPlan = Forms("FormPlanPrincipal")
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoAltaPM_Click", Err
End Sub

Private Sub ComandoDetalle_Click()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormRiesgo") Then DoCmd.Close acForm, "FormRiesgo", acSaveNo
    m_EsAlta = EnumSiNo.No
    Set m_ObjRiesgoAlInicio = Nothing
    m_EsAlta = EnumSiNo.No
    DoCmd.OpenForm "FormRiesgo"
    Forms("FormRiesgo").ComandoVerEdicion.Visible = False
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoDetalle_Click", Err
End Sub

Private Sub ComandoEliminar_Click()
    Dim m_IdRiesgo As String
    Dim m_NodoPadre As MSComctlLib.Node
    
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    
    If m_ObjRiesgoActivo Is Nothing Then Err.Raise 1000, , "No se sabe el riesgo activo"
    If m_ObjRiesgoActivo.Borrable = EnumSiNo.No Then Err.Raise 1000, , "El riesgo no permite ser borrado"
    If m_ObjRiesgoActivo.Edicion.UsuarioConectadoAutorizado = EnumSiNo.No Then Err.Raise 1000, , "Usuario no autorizado"
    
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    If MsgBox("¿Desea realmente borrar el riesgo seleccionado?", vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de un riesgo") <> vbYes Then Exit Sub
    
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_IdRiesgo = m_ObjRiesgoActivo.IDRiesgo
    m_ObjRiesgoActivo.Borrar m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    
    Set m_NodoPadre = Form_FormRiesgosGestion.m_Arbol.SelectedItem.Parent
    If Not Form_FormRiesgosGestion.m_ColRiesgosAplicados Is Nothing Then
        If Form_FormRiesgosGestion.m_ColRiesgosAplicados.Exists(CStr(m_IdRiesgo)) Then
            Form_FormRiesgosGestion.m_ColRiesgosAplicados.Remove (CStr(m_IdRiesgo))
        End If
    End If
    Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    Form_FormRiesgosGestion.SeleccionarNodo m_ObjEdicionActiva
    MsgBox "Riesgo borrado con éxito", vbInformation, "Riesgo borrado"
    Exit Sub
errores:
    GestionarError "ComandoEliminar_Click", Err
End Sub

Private Sub ComandoVerInformeRiesgo_Click()
    Dim m_URLInforme As String
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    m_URLInforme = GenerarInformeRiesgoHTML(p_IDRiesgo:=m_ObjRiesgoActivo.IDRiesgo, p_hWnd:=Application.hWndAccessApp, p_Error:=m_Error)
    If m_Error <> "" Then Err.Raise 1000
    AvanceCerrar
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "ComandoVerInformeRiesgo_Click", Err
End Sub

' ==========================================
' CARGA DEL FORMULARIO
' ==========================================

Private Sub Form_Load()
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    If m_ObjRiesgoActivo Is Nothing Then
        VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
        Exit Sub
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "Form_Load", Err
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim m_ErroresRiesgo As String
    Dim m_EstadoTexto As String
    Dim m_RiesgoAltoOMuyAltoTexto As String
    Dim blnPermitidoEditar As Boolean
    Dim m_Edicion As Edicion
    Dim m_DatosPublicabilidad As tPublicabilidadRiesgoDatos
    Dim m_ChecksPublicabilidad As Scripting.Dictionary
    Dim m_VeredictoPublicabilidad As EnumPublicabilidadVeredicto
    Dim m_Publicable As EnumSiNo
    
    ' Variables para Histórico
    Dim m_ObjColRiesgosEstados As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Resultado As String
    Dim dato As Variant
    Dim sFechaEstado As String, sEstadoRiesgo As String
    
    On Error GoTo errores
    
    Me.ComandoDetalle.Enabled = True
    Me.ComandoEliminar.Visible = False
    Me.ComandoEliminar.Enabled = False
    Me.ComandoAltaPC.Enabled = False
    Me.ComandoAltaPM.Enabled = False
    
    With m_ObjRiesgoActivo
        ' Título
        m_RiesgoAltoOMuyAltoTexto = .RiesgoAltoOMuyAltoTexto
        If m_RiesgoAltoOMuyAltoTexto = "Sí" Then
            Me.lblTitulo.Caption = "Riesgo " & .CodigoRiesgo & " RIESGO ALTO O MUY ALTO****"
            Me.lblTitulo.ForeColor = RGB(255, 0, 0)
        Else
            Me.lblTitulo.Caption = "Riesgo " & .CodigoRiesgo
            Me.lblTitulo.ForeColor = RGB(0, 102, 255)
        End If
        
        ' Permisos
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
            Me.ComandoAltaPM.Enabled = True
            Me.ComandoAltaPC.Enabled = True
            If .Borrable = EnumSiNo.Sí Then
                Me.ComandoEliminar.Visible = True
                Me.ComandoEliminar.Enabled = True
            End If
        End If
        
        ' Campos básicos
        Me.CodigoRiesgo = .CodigoRiesgo
        If .Priorizacion <> "" Then Me.Priorizacion = .Priorizacion
        If .Origen <> "" Then Me.Origen = .Origen
        If .Descripcion <> "" Then Me.Descripcion = .Descripcion
        
        ' Causa Raíz
        If m_ObjProyectoActivo.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            Me.lblCRaiz.Visible = False
            Me.CausaRaiz.Visible = False
            Me.Descripcion.Height = ALTO_DESCRIPCION_SIN_CRAIZ
            Me.Descripcion.FontSize = 11
        Else
            Me.lblCRaiz.Visible = True
            Me.CausaRaiz.Visible = True
            Me.Descripcion.Height = ALTO_DESCRIPCION_CON_CRAIZ
            If .CausaRaiz <> "" Then Me.CausaRaiz = .CausaRaiz
            Me.Descripcion.FontSize = 9
            Me.CausaRaiz.FontSize = 9
        End If
        
        ' Resto de campos
        If .EntidadDetecta <> "" Then Me.EntidadDetecta = .EntidadDetecta
        If .DetectadoPor <> "" Then Me.DetectadoPor = .DetectadoPor
        If .FechaDetectado <> "" Then Me.FechaDetectado = .FechaDetectado
        
        If .FechaMaterializado <> "" Then
            Me.FechaMaterializado = .FechaMaterializado
            Me.FechaMaterializado.ForeColor = m_ColorNOOK
        Else
            Me.FechaMaterializado.ForeColor = m_ColorOK
        End If
        
        If .FechaRetirado <> "" Then
            Me.FechaRetirado = .FechaRetirado
            Me.FechaRetirado.ForeColor = m_ColorNOOK
        Else
            Me.FechaRetirado.ForeColor = m_ColorOK
        End If
        
        If .ImpactoGlobal <> "" Then Me.ImpactoGlobal = .ImpactoGlobal
        Me.Estado = .ESTADOCalculadoTexto
        Me.FechaEstado = .FechaEstado
        Me.RiesgoAltoOMuyAlto = m_RiesgoAltoOMuyAltoTexto
        
       
        m_ErroresRiesgo = .ErroresRiesgoTexto
        p_Error = .Error
        If p_Error <> "" Then Err.Raise 1000
        
       
        
        ' Informe / Carencias
        If m_ObjRiesgoActivo.Edicion.EsActivo = EnumSiNo.Sí Then
            Me.ComandoVerInformeRiesgo.Visible = True
            If ConstruirDatosPublicabilidadRiesgo(m_ObjRiesgoActivo, m_DatosPublicabilidad, , p_Error) = EnumSiNo.No Then Err.Raise 1000
            m_Publicable = EvaluarPublicabilidadRiesgo(m_DatosPublicabilidad, m_ChecksPublicabilidad, m_VeredictoPublicabilidad, p_Error)
            If p_Error <> "" Then Err.Raise 1000

            If m_VeredictoPublicabilidad = EnumPublicabilidadVeredicto.NoPublicable Then
                Me.ComandoVerInformeRiesgo.ForeColor = 2366701
                Me.ComandoVerInformeRiesgo.Caption = "Carencias"
            Else
                Me.ComandoVerInformeRiesgo.ForeColor = 16737792
                Me.ComandoVerInformeRiesgo.Caption = "Sin Carencias"
            End If
        Else
            Me.ComandoVerInformeRiesgo.Visible = False
        End If
        
        ' ----------------------------------------------------------------------
        ' CARGA DE HISTÓRICO DE ESTADOS (Nueva Funcionalidad)
        ' ----------------------------------------------------------------------
        ' Requiere control: lstEstadosHistoricos (ListBox)
        On Error Resume Next ' Evitar fallo si el control no existe aún
        Me.lstEstadosHistoricos.RowSource = ""
        Me.lstEstadosHistoricos.AddItem "Estado;Fecha"
        
        Set m_ObjColRiesgosEstados = getEstadosDiferentesHastaEdicion( _
                                        m_Edicion, _
                                        .CodigoRiesgo, _
                                        CStr(m_Edicion.FechaPublicacion), _
                                        "", _
                                        p_Error)
                                        
        If p_Error <> "" Then Err.Raise 1000
        
        If Not m_ObjColRiesgosEstados Is Nothing Then
            For Each m_Id In m_ObjColRiesgosEstados
                m_Resultado = m_ObjColRiesgosEstados(m_Id)
                dato = Split(m_Resultado, "|")
                sEstadoRiesgo = dato(0)
                
                If UBound(dato) >= 1 Then
                    sFechaEstado = dato(1)
                    If IsDate(sFechaEstado) Then
                        sFechaEstado = Format(sFechaEstado, "dd/mm/yyyy")
                    End If
                Else
                    sFechaEstado = ""
                End If
                
                Me.lstEstadosHistoricos.AddItem sEstadoRiesgo & ";" & sFechaEstado
            Next
        End If
        On Error GoTo errores
        
    End With

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub GestionarError(Metodo As String, objErr As ErrObject)
    DoCmd.Hourglass False
    If objErr.Number <> 1000 Then
        m_Error = "Al " & Metodo & " se ha producido el error n: " & objErr.Number & vbNewLine & "Detalle: " & objErr.Description
        CorreoAlAdministrador m_Error
        MsgBox m_Error, vbCritical, "Error"
    Else
        MsgBox m_Error, vbExclamation, "Advertencia"
    End If
End Sub

Private Sub m_FormPlan_PlanNuevo(p_Plan As Object)
    On Error GoTo errores
    VBA.DoEvents: DoCmd.Hourglass True: VBA.DoEvents
    m_Error = ""
    If TypeOf p_Plan Is PM Then
        Form_FormRiesgosGestion.CargarArbolPM P_NodoRiesgo:=Form_FormRiesgosGestion.m_NodoSeleccionado, _
                                            p_Riesgo:=m_ObjRiesgoActivo, p_PM:=p_Plan, _
                                            p_borrarNodoSeleccionado:=EnumSiNo.Sí, p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    Else
        Form_FormRiesgosGestion.CargarArbolPC P_NodoRiesgo:=Form_FormRiesgosGestion.m_NodoSeleccionado, _
                                            p_Riesgo:=m_ObjRiesgoActivo, p_PC:=p_Plan, _
                                            p_borrarNodoSeleccionado:=EnumSiNo.Sí, p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    End If
    If m_Error <> "" Then Err.Raise 1000
    Form_FormRiesgosGestion.SeleccionarNodo p_Plan, m_Error
    If m_Error <> "" Then Err.Raise 1000
    VBA.DoEvents: DoCmd.Hourglass False: VBA.DoEvents
    Exit Sub
errores:
    GestionarError "m_FormPlan_PlanNuevo", Err
End Sub

