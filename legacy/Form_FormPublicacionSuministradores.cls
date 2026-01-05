VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormPublicacionSuministradores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private m_SuministadorEdicionSeleccionado As EdicionSuministrador

Private m_Error As String

Private Function EstablecerLista( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_EdicionSuministrador As EdicionSuministrador
    Dim m_Anexo As String
    
    On Error GoTo errores
    Set m_ObjEdicionActiva.ColSuministradores = Nothing
    Set m_Col = m_ObjEdicionActiva.ColSuministradores
    p_Error = m_ObjEdicionActiva.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.LIstaSuministradores.RowSource = "ID;Nombre;Anexo"
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Id In m_Col
        Set m_EdicionSuministrador = m_Col(m_Id)
        If m_EdicionSuministrador.IDAnexo = "" Then
            m_Anexo = "No"
        Else
            m_Anexo = "Sí"
        End If
        Me.LIstaSuministradores.AddItem m_EdicionSuministrador.ID & ";" & _
                            Replace(m_EdicionSuministrador.Suministrador.Nombre, ";", ":") & ";" & _
                            m_Anexo
        Set m_EdicionSuministrador = Nothing
    Next
    Me.LIstaSuministradores.Selected(1) = True
    LIstaSuministradores_Click
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo EstablecerLista ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ComandoAnexar_Click()
    
    Dim m_URL As String
    
    
    Dim m_lineaAntes As String
    Dim m_LineaAhora As String
    Dim m_valorFinal As String
    Dim lst As ListBox
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_SuministadorEdicionSeleccionado Is Nothing Then
        Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(Nz(Me.LIstaSuministradores.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_SuministadorEdicionSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    m_URL = Seleccionar(True, "Seleccione el archivo a anexar", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set lst = Me.LIstaSuministradores
    m_lineaAntes = lst.Column(0) & ";" & lst.Column(1) & ";" & lst.Column(2)
    m_SuministadorEdicionSeleccionado.AnexarPorTecnico m_URL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(m_SuministadorEdicionSeleccionado.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_SuministadorEdicionSeleccionado.Anexo Is Nothing Then
        m_valorFinal = "Sí"
    Else
        m_valorFinal = "No"
    End If
    
    m_LineaAhora = lst.Column(0) & ";" & lst.Column(1) & ";" & m_valorFinal
    
    lst.RowSource = Replace(lst.RowSource, m_lineaAntes, m_LineaAhora)
    lst.Requery
    
    CachePublicabilidad_RecalcularEdicionYResetear m_ObjEdicionActiva, Nothing, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    LIstaSuministradores_Click
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoEliminarEvidenciaAcuerdoSuministrador_Click()
    
    Dim m_lineaAntes As String
    Dim m_LineaAhora As String
    Dim m_valorFinal As String
    Dim lst As ListBox
    On Error GoTo errores
    
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_SuministadorEdicionSeleccionado Is Nothing Then
        Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(Nz(Me.LIstaSuministradores.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_SuministadorEdicionSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    pregunta = MsgBox("¿Desea eliminar la evidencia de acuerdo de este suministrador?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar evidencia")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set lst = Me.LIstaSuministradores
    m_lineaAntes = lst.Column(0) & ";" & lst.Column(1) & ";" & lst.Column(2)
    m_SuministadorEdicionSeleccionado.EliminarEvidencia m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(m_SuministadorEdicionSeleccionado.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_SuministadorEdicionSeleccionado.Anexo Is Nothing Then
        m_valorFinal = "Sí"
    Else
        m_valorFinal = "No"
    End If
    Me.LIstaSuministradores.SetFocus
    m_LineaAhora = lst.Column(0) & ";" & lst.Column(1) & ";" & m_valorFinal
    
    lst.RowSource = Replace(lst.RowSource, m_lineaAntes, m_LineaAhora)
    lst.Requery
    LIstaSuministradores_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEliminarEvidenciaAcuerdoSuministrador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    
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
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    Me.ImagenAnexo.Visible = False
    
    
    EstablecerLista p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El mëtodo Establecerdatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Private Sub ImagenAnexo_Click()
        
   On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_SuministadorEdicionSeleccionado Is Nothing Then
        Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(Nz(Me.LIstaSuministradores.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_SuministadorEdicionSeleccionado Is Nothing Then
            m_Error = "Seleccione un elemento de la lista"
            Err.Raise 1000
        End If
    End If
    If m_SuministadorEdicionSeleccionado.Anexo Is Nothing Then
        m_Error = "El elemento seleccionado no parece tener ningún anexo asociado"
        Err.Raise 1000
    End If
    
    AbrirEnLocal m_SuministadorEdicionSeleccionado.Anexo.URLAnexo, Me.hWnd, m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenAnexo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub LIstaSuministradores_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.ImagenAnexo.Visible = False
    Me.ComandoEliminarEvidenciaAcuerdoSuministrador.Visible = False
    If Nz(Me.LIstaSuministradores.Column(0), "") = "" Then
        Set m_SuministadorEdicionSeleccionado = Nothing
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If m_SuministadorEdicionSeleccionado Is Nothing Then
        Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(Me.LIstaSuministradores.Column(0), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_SuministadorEdicionSeleccionado.ID <> Me.LIstaSuministradores.Column(0) Then
            Set m_SuministadorEdicionSeleccionado = Constructor.getSuministradorEnEdicion(Me.LIstaSuministradores.Column(0), m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_SuministadorEdicionSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoAnexar.Enabled = True
    Else
        Me.ComandoAnexar.Enabled = False
    End If
    If Not m_SuministadorEdicionSeleccionado.Anexo Is Nothing Then
        Me.ImagenAnexo.Visible = True
        Me.ComandoEliminarEvidenciaAcuerdoSuministrador.Visible = True
        If blnPermitidoEscribir = True Then
            Me.ComandoEliminarEvidenciaAcuerdoSuministrador.Enabled = True
        Else
            Me.ComandoEliminarEvidenciaAcuerdoSuministrador.Enabled = False
        End If
    Else
        Me.ImagenAnexo.Visible = False
        Me.ComandoEliminarEvidenciaAcuerdoSuministrador.Visible = False
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



