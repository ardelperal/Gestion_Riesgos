VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgoMaterializado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private blnPermitidoEditar As Boolean
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_VerDescripcion As String
Private m_RiesgoMatSeleccionado As RiesgoMaterializacion
Private m_NC As NC
Private m_Error As String

Private Sub ComandoFechaMaterializado_Click()
    
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_FechaMaterializado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    With m_ObjRiesgoActivo
        m_RiesgoEnAceptacion = RiesgoEnAceptacion(m_ObjRiesgoActivo.EstadoEnum)
        If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
            m_Error = "El riesgo est¿ en camino de ser aceptado"
            Err.Raise 1000
        End If
        m_RiesgoEnRetirada = RiesgoEnRetirada(.EstadoEnum)
        If m_RiesgoEnRetirada = EnumSiNo.Sí Then
            m_Error = "El riesgo est¿ en camino de ser retirado"
            Err.Raise 1000
        End If
        If .EstadoEnum = EnumRiesgoEstado.Materializado Then
            m_Error = "El riesgo aparece como materializado"
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("Al materializarse un riesgo, recuerde que ha de ejecutar el plan de contingencia y " & _
                "publicar la edici¿n para que quede constancia", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Materialización")
        If pregunta <> vbYes Then
            
            Exit Sub
        End If
        m_FechaMaterializado = Nz(InputBox("Introduzca la fecha de Materialización", "Fecha Materialización", Date), "")
        If Not IsDate(m_FechaMaterializado) Then
            Exit Sub
        End If
         VBA.DoEvents
        DoCmd.Hourglass True
        VBA.DoEvents
        Me.FechaMaterializado = m_FechaMaterializado
        
        m_ObjRiesgoActivo.FechaMaterializado = Me.FechaMaterializado
        .MaterializacionRegistrar Me.FechaMaterializado, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(.IDRiesgo)
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
        If m_Error <> "" Then
           Err.Raise 1000
        End If
        
    End With
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
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoFechaMaterializado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
   
End Sub


Private Sub ComandoQuitarMaterializado_Click()
    
    On Error GoTo errores
    m_Error = ""
    pregunta = MsgBox("¿Desea habilitar el riesgo?", vbExclamation + vbYesNo + vbDefaultButton2, "Habilitar el riesgo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    With m_ObjRiesgoActivo
         .MaterializacionQuitarRegistrar .EstadoEnum, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjRiesgoAlInicio = Constructor.getRiesgo(.IDRiesgo)
        Me.ComandoFechaMaterializado.SetFocus
        Me.ComandoQuitarMaterializado.Visible = False
        If IsDate(.FechaMaterializado) Then
            Me.FechaMaterializado = .FechaMaterializado
        Else
            Me.FechaMaterializado = Null
        End If
        
    End With
    ActualizarEtiquetaEstadoRiesgo EnumSiNo.Sí
    EstablecerDatos
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
        m_Error = "Al ComandoQuitarMaterializado_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoVerEnTarea_Click()
   
    On Error GoTo errores

    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If FormularioAbierto("FormCalidadRiesgoMaterializaciones") Then
        DoCmd.Close acForm, "FormCalidadRiesgoMaterializaciones", acSaveNo
    End If
    DoCmd.OpenForm "FormCalidadRiesgoMaterializaciones"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents


    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerEnTarea_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    On Error GoTo errores
    EstablecerLista p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
    Me.FechaMaterializado.Enabled = False
    Set m_Edicion = m_ObjRiesgoActivo.Edicion
    If m_Edicion.EsActivo = EnumSiNo.No Then
        blnPermitidoEditar = False
    Else
        If m_Edicion.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            blnPermitidoEditar = True
        
        Else
            blnPermitidoEditar = False
        End If
    End If
    If blnPermitidoEditar = True Then
        If FormularioAbierto("FormRiesgosGestion") Then
            Set m_NodoSeleccionado = Form_FormRiesgosGestion.m_NodoSeleccionado
            m_VerDescripcion = Nz(Forms("FormRiesgosGestion").Controls("comboVerDescripcion"), "")
        End If
        Me.ComandoFechaMaterializado.Enabled = True
        If m_ObjRiesgoActivo.EstadoEnum = EnumRiesgoEstado.Materializado Then
            Me.ComandoQuitarMaterializado.Visible = True
            Me.ComandoQuitarMaterializado.Enabled = True
            Me.FechaMaterializado = m_ObjRiesgoActivo.FechaMaterializado
        Else
            Me.ComandoQuitarMaterializado.Visible = False
            Me.FechaMaterializado = Null
        End If
    Else
        Me.ComandoFechaMaterializado.Enabled = False
        Me.ComandoQuitarMaterializado.Visible = False
        Me.FechaMaterializado = m_ObjRiesgoActivo.FechaMaterializado
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
   
End Function
Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Id As Variant
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_Decidido As String
    Dim m_Edicion As String
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    Me.ComandoVerEnTarea.Visible = False
    Set lst = Me.ListasMaterializaciones
    lst.RowSource = "ID;Edici¿n;FMaterializacion;Mat.;Decidido"
    
    Set m_ObjRiesgoActivo.ColMaterializaciones = Nothing
    Set m_Col = m_ObjRiesgoActivo.ColMaterializaciones
    p_Error = m_ObjRiesgoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    If EsTecnico = EnumSiNo.No Then
        Me.ComandoVerEnTarea.Visible = True
    End If
    For Each m_Id In m_Col
        Set m_RiesgoMat = m_Col(m_Id)
        If m_RiesgoMat.Decidido = EnumSiNo.Sí Then
            m_Decidido = "Sí"
        Else
            m_Decidido = "No"
        End If
        m_Edicion = m_RiesgoMat.Edicion.Edicion
        
        lst.AddItem m_RiesgoMat.ID & ";" & m_Edicion & ";" & m_RiesgoMat.Fecha & ";" & m_RiesgoMat.EsMaterializacion & ";" & m_Decidido
        Set m_RiesgoMat = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function





