VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgoDefinicionNoBiblioteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Error As String

Private Sub Descripcion_Exit(Cancel As Integer)
    On Error Resume Next
    
   
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo

End Sub

Private Sub DetectadoPor_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub



Private Sub EntidadDetecta_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    RellenarComboDetectadoPor Nz(Me.EntidadDetecta, "")
    
    
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub


Private Sub FechaDetectado_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
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

Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    
    
    
    Dim m_PriorizacionCorrecta As EnumSiNo
    On Error GoTo errores
    
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se puede determinar la edición activa"
        Err.Raise 1000
    End If
    
    If Form_FormRiesgo.m_PermitidoEditar = Empty Then
        With m_ObjEdicionActiva
            If .EsActivo = EnumSiNo.No Then
               Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.No
            Else
                If .Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                    Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.Sí
                
                Else
                    Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.No
                End If
            End If
        End With
    
    End If
    
    EstablecerComboOrigen Me.Origen, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If Form_FormRiesgo.m_PermitidoEditar = EnumSiNo.Sí Then
        Me.AllowEdits = True
    Else
        Me.AllowEdits = False
    End If
    With m_ObjRiesgoActivo
        If .Descripcion <> "" Then
            Me.Descripcion = .Descripcion
        End If
        
        If IsDate(.FechaDetectado) Then
            Me.FechaDetectado = .FechaDetectado
        End If
        If .EntidadDetecta <> "" Then
            Me.EntidadDetecta = .EntidadDetecta
        End If
        RellenarComboDetectadoPor .EntidadDetecta, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If .DetectadoPor <> "" Then
            Me.DetectadoPor = .DetectadoPor
        End If
        If .Origen <> "" Then
            Me.Origen = .Origen
        End If
        If IsNumeric(.Priorizacion) Then
            Me.Priorizacion = .Priorizacion
        End If
        If .IDRiesgo = "" Then
            m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                Nz(Me.Priorizacion, ""))
        Else
            m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                Nz(Me.Priorizacion, ""), .IDRiesgo)
        End If
    End With
    If m_PriorizacionCorrecta = EnumSiNo.Sí Then
        EstablecerControlCombo Me.Priorizacion, EnumSiNo.Sí
    Else
        EstablecerControlCombo Me.Priorizacion, EnumSiNo.No
    End If
    
    
    RellenarBordeCamposObligatorios Me
    
   
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function RellenarComboDetectadoPor(Optional p_EntidadDetecta As String, Optional ByRef p_Error As String) As String

    Dim cmb As ComboBox
    Dim ColSuministradores As Scripting.Dictionary
    Dim m_Suministrador As Suministrador
    Dim m_SuministradorEdicion As EdicionSuministrador
    Dim m_Id As Variant
    
    
    
    Dim m_ValorActual As String
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_objColRiesgos As Scripting.Dictionary
    Dim m_RowSource As String
    On Error GoTo errores
    
    Set cmb = Me.DetectadoPor
    m_ValorActual = Nz(cmb.Value, "")
    cmb.Value = Null
    cmb.RowSource = ""
    
    If p_EntidadDetecta <> "Persona" And p_EntidadDetecta <> "Empresa" Then
        Exit Function
    End If
    If p_EntidadDetecta = "Persona" Then
        
        m_RowSource = m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados
        p_Error = m_ObjEdicionActiva.Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        cmb.RowSource = m_RowSource

        If m_ValorActual <> "" Then
            cmb.Value = m_ValorActual
            
        End If
        
        Exit Function
    End If
    
    Set ColSuministradores = m_ObjRiesgoActivo.Edicion.ColSuministradores
    If ColSuministradores Is Nothing Then
        Exit Function
    End If
    For Each m_Id In ColSuministradores
        Set m_SuministradorEdicion = ColSuministradores(m_Id)
        Set m_Suministrador = m_SuministradorEdicion.Suministrador
        cmb.AddItem m_Suministrador.Nombre
        Set m_Suministrador = Nothing
        Set m_SuministradorEdicion = Nothing
    Next
    cmb.Value = m_ValorActual
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboDetectadoPor ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Private Sub Origen_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
    
    
    RellenarBordeCamposObligatorios Me
    EstablecerColorPestañasRiesgo
End Sub




Private Sub Priorizacion_Change()
    Dim m_PriorizacionCorrecta As EnumSiNo
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.FechaDetectado.SetFocus
    Me.Priorizacion.SetFocus
   
    If m_EsAlta = EnumSiNo.Sí Then
    
        m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                    Nz(Me.Priorizacion, ""))
    Else
        m_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjEdicionActiva.IDEdicion, _
                                    Nz(Me.Priorizacion, ""), m_ObjRiesgoActivo.IDRiesgo)
    End If
    
    EstablecerColorPestañasRiesgo p_PriorizacionCorrecta:=m_PriorizacionCorrecta
    If Nz(Me.Priorizacion, "") = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    'm_PriorizacionCorrecta = Constructor.EsPriorizacionCorrecta(m_ObjRiesgoActivo.IDEdicion, Me.Priorizacion, m_ObjRiesgoActivo.IDRiesgo)
    If m_PriorizacionCorrecta = EnumSiNo.No Then
        If Not m_ObjRiesgoAlInicio Is Nothing Then
            If IsNumeric(m_ObjRiesgoAlInicio.Priorizacion) Then
                Me.Priorizacion = m_ObjRiesgoAlInicio.Priorizacion
                m_ObjRiesgoActivo.Priorizacion = m_ObjRiesgoAlInicio.Priorizacion
                
                EstablecerColorPestañasRiesgo
            End If
        Else
            
            Me.Priorizacion = Null
        End If
        
        
        m_Error = "Esa priorización no es correcta. Puede que sea superior a los riesgos no retirados, o no sea un número válido, o esté repetida"
        Err.Raise 1000
    Else
       
        EstablecerControlCombo Me.Priorizacion, EnumSiNo.Sí
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Priorizacion_Change se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

