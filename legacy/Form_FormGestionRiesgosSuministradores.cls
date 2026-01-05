VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormGestionRiesgosSuministradores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private blnPermitidoEscribir As Boolean
Private m_ProyectoSuministradorSeleccionado As ProyectoSuministrador


Private m_Error As String

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjProyectoAlInicio.ColSuministradores = Nothing
    EstablecerLista m_Error
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
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoActualizarSuministradores_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not m_ObjProyectoActivo.Expediente Is Nothing Then
        Set m_ObjProyectoActivo.Expediente = Nothing
    End If
    
    EstablecerCombo m_Error
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
        m_Error = "Al ComandoActualizarSuministradores_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoAlta_Click()

    Dim m_Suministrador As Suministrador
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_Suministrador = Constructor.getSuministrador(p_Id:=Nz(Me.Suministrador.Column(0), ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Suministrador Is Nothing Then
        m_Error = "Seleccione un suministrador del desplegable"
        Err.Raise 1000
    End If
    m_ObjProyectoActivo.RegistrarSuministrador m_Suministrador.IDSuministrador, Nz(Me.GestionCalidad, ""), m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo.ColSuministradores = Nothing
    If Not m_ObjProyectoAlInicio Is Nothing Then
        Set m_ObjProyectoAlInicio.ColSuministradores = Nothing
    End If
    Set m_ObjProyectoActivo.EdicionActiva.ColSuministradores = Nothing
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Suministrador = Null
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
'-------------------------------------------
' Nombre: EstablecerCombo
' Propósito: Poblar el combo de suministradores usando la lógica jerárquica (retrocompatible)
' Parámetros:
'   - p_Error As String (ByRef, opcional): detalle si ocurre error
' Retorno: String (vacío si OK)
'-------------------------------------------

Public Function EstablecerCombo( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim cmb As ComboBox
    Dim m_Id As Variant
    Dim m_Suministrador As Suministrador
    Dim m_Col As Scripting.Dictionary
    Dim m_Nombre As String
    Dim sErr As String
    Dim sIDExp As String
    On Error GoTo errores
    p_Error = ""
    Set cmb = Me.Suministrador
    cmb.RowSource = ""
    
    ' Validaciones de contexto
    If m_ObjProyectoActivo Is Nothing Then
        Exit Function
    End If
    If m_ObjProyectoActivo.Expediente Is Nothing Then
        Exit Function
    End If
    
    sIDExp = Nz(m_ObjProyectoActivo.Expediente.IDExpediente, "")
    If sIDExp = "" Then
        Exit Function
    End If
    
    ' Lógica retrocompatible: selecciona viejo/nuevo según TempVars("CadenaJerarquicaModelo")
    Set m_Col = Constructor.getExpedienteSuministradores_RC(sIDExp, sErr)
    If sErr <> "" Then
        p_Error = sErr
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    
    ' Poblar combo: IDSuministrador;Nombre/Nemotécnico (limpia ';')
    For Each m_Id In m_Col
        Set m_Suministrador = m_Col(m_Id)
        m_Nombre = Nz(m_Suministrador.Nemotecnico, "")
        If m_Nombre = "" Then
            m_Nombre = Replace(Nz(m_Suministrador.Nombre, ""), ";", ":")
        End If
        cmb.AddItem m_Suministrador.IDSuministrador & ";" & m_Nombre
        Set m_Suministrador = Nothing
    Next

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstalbecerCombo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
'-------------------------------------------
' Nombre: RefrescarSuministradoresExpediente
' Propósito: Utilidad para refrescar el combo según el estado actual del proyecto/expediente
' Parámetros:
'   - p_Error As String (ByRef, opcional)
' Retorno: String (vacío si OK)
'-------------------------------------------
Public Function RefrescarSuministradoresExpediente( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    On Error GoTo ManejoError
    p_Error = ""
    ' Limpia cachés relevantes y vuelve a poblar combo/lista
    If Not m_ObjProyectoActivo Is Nothing Then
        If Not m_ObjProyectoActivo.Expediente Is Nothing Then
            Set m_ObjProyectoActivo.Expediente = Nothing
        End If
        Set m_ObjProyectoActivo.ColSuministradores = Nothing
        If Not m_ObjProyectoAlInicio Is Nothing Then
            Set m_ObjProyectoAlInicio.ColSuministradores = Nothing
        End If
    End If
    
    EstablecerCombo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Function
ManejoError:
    If Err.Number <> 1000 Then
        p_Error = "El método RefrescarSuministradoresExpediente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Sub ComandoCalidad_Click()

    Dim m_lineaAntes As String
    Dim m_LineaAhora As String
    Dim m_valorFinal As String
    Dim m_ValorActual As String
    Dim lst As ListBox
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ProyectoSuministradorSeleccionado Is Nothing Then
        Set m_ProyectoSuministradorSeleccionado = Constructor.getSuministradorEnProyecto(Nz(Me.LIstaSuministradores.Column(0), ""), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ProyectoSuministradorSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    Set lst = Me.LIstaSuministradores
    m_lineaAntes = lst.Column(0) & ";" & lst.Column(1) & ";" & lst.Column(2)
    m_ValorActual = m_ProyectoSuministradorSeleccionado.GestionCalidad
    If m_ValorActual = "Sí" Then
        m_valorFinal = "No"
    ElseIf m_ValorActual = "No" Then
        m_valorFinal = "Sí"
    Else
        m_Error = "No se sabe el valor actual de la gestión de calidad por parte del suministrador"
        Err.Raise 1000
    End If
    m_ProyectoSuministradorSeleccionado.CambiarCalidad m_valorFinal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_valorFinal = m_ProyectoSuministradorSeleccionado.GestionCalidad
    m_LineaAhora = lst.Column(0) & ";" & lst.Column(1) & ";" & m_valorFinal
    
    lst.RowSource = Replace(lst.RowSource, m_lineaAntes, m_LineaAhora)
    lst.Requery
    Set m_ObjProyectoActivo = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, m_Error)
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
        m_Error = "Al ComandoCalidad_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_ProyectoSuministradorSeleccionado Is Nothing Then
        Set m_ProyectoSuministradorSeleccionado = Constructor.getSuministradorEnProyecto(Nz(Me.LIstaSuministradores.Column(0), ""), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ProyectoSuministradorSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar el suministrador de la gestión de riesgos?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Suministrador de gestión de riesgos")
    If pregunta <> vbYes Then
        Exit Sub
    End If
     VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_ProyectoSuministradorSeleccionado.Eliminar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ProyectoSuministradorSeleccionado = Nothing
    Set m_ObjProyectoActivo.ColSuministradores = Nothing
    If Not m_ObjProyectoAlInicio Is Nothing Then
        Set m_ObjProyectoAlInicio.ColSuministradores = Nothing
    End If
    EstablecerLista m_Error
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
        m_Error = "Al ComandoEliminar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoVerSuministrador_Click()
    
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_ObjSuministradorActivo = Constructor.getSuministrador(p_Id:=Nz(Me.Suministrador.Column(0), ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjSuministradorActivo Is Nothing Then
        m_Error = "Seleccione un suministrador del desplegable"
        Err.Raise 1000
    End If
    If FormularioAbierto("FormSuministrador") Then
        DoCmd.Close acForm, "FormSuministrador"
    End If
    
    DoCmd.OpenForm "FormSuministrador"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerSuministrador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub LIstaSuministradores_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.ComandoCalidad.Enabled = False
    Me.ComandoEliminar.Enabled = False
    If Nz(Me.LIstaSuministradores.Column(0), "") = "" Then
        Set m_ProyectoSuministradorSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_ProyectoSuministradorSeleccionado Is Nothing Then
        Set m_ProyectoSuministradorSeleccionado = Constructor.getSuministradorEnProyecto(Me.LIstaSuministradores.Column(0), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ProyectoSuministradorSeleccionado.ID <> Me.LIstaSuministradores.Column(0) Then
            Set m_ProyectoSuministradorSeleccionado = Constructor.getSuministradorEnProyecto(Me.LIstaSuministradores.Column(0), , , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_ProyectoSuministradorSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoCalidad.Enabled = True
        Me.ComandoEliminar.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al LIstaSuministradores_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    
    On Error GoTo errores
    If m_ObjProyectoActivo Is Nothing Or m_ObjProyectoAlInicio Is Nothing Then
        m_Error = "Es una edición y no se conoce el proyecto activo"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        blnPermitidoEscribir = False
    Else
        If m_ObjProyectoActivo.EsActivo = EnumSiNo.Sí Then
            blnPermitidoEscribir = True
        Else
            blnPermitidoEscribir = False
        End If
    End If
    EstablecerCombo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerLista p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoAlta.Enabled = True
       
    Else
        Me.ComandoAlta.Enabled = False
        
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Function EstablecerLista( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_ProyectoSuministrador As ProyectoSuministrador
    On Error GoTo errores
    
    Set m_Col = m_ObjProyectoActivo.ColSuministradores
    p_Error = m_ObjProyectoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.LIstaSuministradores.RowSource = "IDSuministrador;Nombre;Gest."
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Id In m_Col
        Set m_ProyectoSuministrador = m_Col(m_Id)
        Me.LIstaSuministradores.AddItem m_ProyectoSuministrador.ID & ";" & _
                            Replace(m_ProyectoSuministrador.Suministrador.Nombre, ";", ":") & ";" & _
                            m_ProyectoSuministrador.GestionCalidad
        Set m_ProyectoSuministrador = Nothing
    Next
    Me.LIstaSuministradores.Selected(1) = True
    LIstaSuministradores_Click
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub Suministrador_Change()
    On Error Resume Next
    If Nz(Me.Suministrador, "") = "" Then
        Me.ComandoVerSuministrador.Enabled = False
    Else
        Me.ComandoVerSuministrador.Enabled = True
    End If
End Sub
