VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgoPlazoCosteCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Error As String



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

Private Function SeleccionarValorEnLIsta(lst As ListBox, p_Valor As String, Optional ByRef p_Error As String) As String
    
   
    Dim i As Long
    On Error GoTo errores
    
    If p_Valor = "" Then
        Exit Function
    End If
    
    For i = 0 To lst.ListCount - 1
        If Nz(lst.Column(0, i), "") = p_Valor Then
            lst.Selected(i) = True
            If lst.Name = Me.ListaCalidad.Name Then
                ListaCalidad_Click
            End If
            
            
            Exit For
        End If
    Next
    
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarValorEnLIsta ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
   
'    On Error GoTo errores
On Error Resume Next
    Me.AllowEdits = True
    EstablecerListas p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
    With m_ObjRiesgoActivo
        
        SeleccionarValorEnLIsta Me.ListaPlazo, .Plazo, p_Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
        If .Plazo = "" Then
            EstablecerControlCombo Me.ListaPlazo, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaPlazo, EnumSiNo.Sí
        End If
    
        SeleccionarValorEnLIsta Me.ListaCoste, .Coste, p_Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
        If .Coste = "" Then
            EstablecerControlCombo Me.ListaCoste, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaCoste, EnumSiNo.Sí
        End If
        SeleccionarValorEnLIsta Me.ListaCalidad, .Calidad, p_Error
'        If p_Error <> "" Then
'            Err.Raise 1000
'        End If
        If .Calidad = "" Then
            EstablecerControlCombo Me.ListaCalidad, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaCalidad, EnumSiNo.Sí
        End If
    End With
    
    
'    Exit Function
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
'        vbNewLine & "Detalle: " & Err.Description
'    End If

End Function
Private Sub ListaPlazo_Click()
    On Error Resume Next
    
    If Me.AllowEdits = True Then
        
        
        m_ObjRiesgoActivo.Plazo = Nz(Me.ListaPlazo.Column(0), "")
        If m_ObjRiesgoActivo.Plazo = "" Then
            EstablecerControlCombo Me.ListaPlazo, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaPlazo, EnumSiNo.Sí
        End If
        EstablecerColorPestañasRiesgo
    End If
End Sub
Private Sub ListaCalidad_Click()
    On Error Resume Next
    Me.DescripcionCalidad = Nz(Me.ListaCalidad.Column(1), "")
    If Me.AllowEdits = True Then
        
        m_ObjRiesgoActivo.Calidad = Nz(Me.ListaCalidad.Column(0), "")
        If m_ObjRiesgoActivo.Calidad = "" Then
            EstablecerControlCombo Me.ListaCalidad, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaCalidad, EnumSiNo.Sí
        End If
        EstablecerColorPestañasRiesgo
    End If
End Sub

Private Sub ListaCoste_Click()
    On Error Resume Next
    
    If Me.AllowEdits = True Then
        m_ObjRiesgoActivo.Coste = Nz(Me.ListaCoste.Column(0), "")
        If m_ObjRiesgoActivo.Coste = "" Then
            EstablecerControlCombo Me.ListaCoste, EnumSiNo.No
        Else
            EstablecerControlCombo Me.ListaCoste, EnumSiNo.Sí
        End If
        EstablecerColorPestañasRiesgo
        
    End If
End Sub

Public Function EstablecerListas(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_indice As Variant
    Dim m_Descripcion As String
    Dim m_ObjAreaImpacto As AreaImpacto
    
    On Error GoTo errores
    
    Set lst = Me.ListaCalidad
    lst.RowSource = "Impacto;Descripción"
    If Not m_ObjEntorno.ColAreaImpactoCalidad Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_indice In m_ObjEntorno.ColAreaImpactoCalidad
            Set m_ObjAreaImpacto = m_ObjEntorno.ColAreaImpactoCalidad(m_indice)
            lst.AddItem m_ObjAreaImpacto.Indice & ";" & m_ObjAreaImpacto.Descripcion
            Set m_ObjAreaImpacto = Nothing
        Next
    End If
    Set lst = Me.ListaCoste
    lst.RowSource = "Impacto;Descripción"
    If Not m_ObjEntorno.ColAreaImpactoCoste Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_indice In m_ObjEntorno.ColAreaImpactoCoste
            Set m_ObjAreaImpacto = m_ObjEntorno.ColAreaImpactoCoste(m_indice)
            lst.AddItem m_ObjAreaImpacto.Indice & ";" & m_ObjAreaImpacto.Descripcion
            Set m_ObjAreaImpacto = Nothing
        Next
    End If
    Set lst = Me.ListaPlazo
    lst.RowSource = "Impacto;Descripción"
    If Not m_ObjEntorno.ColAreaImpactoPlazo Is Nothing Then
        p_Error = m_ObjEntorno.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        For Each m_indice In m_ObjEntorno.ColAreaImpactoPlazo
            Set m_ObjAreaImpacto = m_ObjEntorno.ColAreaImpactoPlazo(m_indice)
            lst.AddItem m_ObjAreaImpacto.Indice & ";" & m_ObjAreaImpacto.Descripcion
            Set m_ObjAreaImpacto = Nothing
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListas ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

