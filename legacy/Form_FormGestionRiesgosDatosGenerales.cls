VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormGestionRiesgosDatosGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda
Attribute m_FormExpedientes.VB_VarHelpID = -1
Private m_UltimoValorJuridica As String
Private m_Expediente As Expediente
Private m_Error As String


Private Sub Aprobado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Aprobado = Nz(Me.Aprobado, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set m_Expediente = Constructor.getExpediente(p_IDExpediente:=Nz(Me.IDExpediente, ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarDatosExpediente m_Expediente, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarBordeCamposObligatorios Me
    Me.Elaborado.SetFocus
    AvanceCerrar
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

Private Sub ComandoIrABusquedaExpedientes_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        If Nz(Me.PalabraClave, "") <> "" Then
            m_FormExpedientes.PalabraClave = Nz(Me.PalabraClave, "")
            Form_FormExpedientesBusqueda.Filtrar
        End If
        
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrABusquedaExpedientes_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
    
End Sub

Private Sub Elaborado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Elaborado = Nz(Me.Elaborado, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    RellenarDatosAlObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Unload se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
   
End Sub

Private Sub m_FormExpedientes_Seleccionado(m_Exp As Expediente)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    RellenarDatosExpediente m_Exp
    If Nz(Me.IDExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    Me.PalabraClave = Null
    RellenarBordeCamposObligatorios Me
    Me.Elaborado.SetFocus
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormExpedientes_Seleccionado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    AvanceCerrar
   
End Sub

Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Me.ComandoBuscarExpediente.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        Me.ComandoBuscarExpediente.SetFocus
        Me.PalabraClave.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoBuscarExpediente_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub


Private Sub ParaInformeAvisos_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.ParaInformeAvisos = Nz(Me.ParaInformeAvisos, "")
    RellenarBordeCamposObligatorios Me
End Sub

Private Sub Revisado_BeforeUpdate(Cancel As Integer)
    m_ObjProyectoActivo.Revisado = Nz(Me.Revisado, "")
    RellenarBordeCamposObligatorios Me
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
   ' Debug.Print Me.IDExpediente.BackColor
    AvanceCerrar
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
    AvanceCerrar
    
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




Private Sub ComandoBuscarExpediente_Click()
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Expediente As Expediente
    
    Dim m_PalabraClave As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    m_PalabraClave = Nz(Me.PalabraClave, "")
    If m_PalabraClave = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    Set m_Col = Constructor.getExpedientesBusqueda(p_PalabraClave:=m_PalabraClave, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Me.IDExpediente = Null
        Me.Juridica = Null
        Me.CodExp = Null
        Me.Nemotecnico = Null
        Me.FechaInicioContrato = Null
        Me.FechaFinContrato = Null
        Me.OrganoContratacion = Null
        Me.ResponsableCalidad = Null
        Me.CodigoDocumento = Null
        Me.CorreoRAC = Null
        If Nz(Me.IDExpediente, "") = "" Then
            Me.ComandoActualizar.Enabled = False
        Else
            Me.ComandoActualizar.Enabled = True
        End If
        RellenarBordeCamposObligatorios Me
        ComandoIrABusquedaExpedientes_Click
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    If m_Col.Count = 1 Then
        For Each m_Id In m_Col
            Set m_Expediente = m_Col(m_Id)
            RellenarDatosExpediente m_Expediente, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If Nz(Me.IDExpediente, "") = "" Then
                Me.ComandoActualizar.Enabled = False
            Else
                Me.ComandoActualizar.Enabled = True
            End If
        Next
        Me.PalabraClave = Null
        RellenarBordeCamposObligatorios Me
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        m_FormExpedientes.PalabraClave = Nz(Me.PalabraClave, "")
        Form_FormExpedientesBusqueda.Filtrar
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscarExpediente_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Function RellenarDatosExpediente(Optional p_Expediente As Expediente, Optional ByRef p_Error As String) As String
    
    Dim m_EnUTE As EnumSiNo
    On Error GoTo errores
    
    If p_Expediente Is Nothing Then
        Me.IDExpediente = Null
        Me.Juridica = Null
        Me.CodExp = Null
        Me.Nemotecnico = Null
        Me.Titulo = Null
        Me.FechaInicioContrato = Null
        Me.FechaFinContrato = Null
        Me.OrganoContratacion = Null
        Me.ResponsableCalidad = Null
        Me.CodigoDocumento = Null
        Me.CorreoRAC = Null
        Me.EnUTE = Null
        Exit Function
    End If
    With p_Expediente
        Me.IDExpediente = .IDExpediente
        Me.Juridica = .CadenaJuridicas
        Me.CodExp = .CodExp
        If Not m_ObjProyectoAlInicio Is Nothing Then
            If m_ObjProyectoAlInicio.NombreProyecto <> "" Then
                Me.Nemotecnico = m_ObjProyectoAlInicio.NombreProyecto
            Else
                Me.Nemotecnico = Null
            End If
        Else
            If .Nemotecnico <> "" Then
                Me.Nemotecnico = .Nemotecnico
            Else
                Me.Nemotecnico = .Titulo
            End If
        End If
        Me.Titulo = .Titulo
        
        If IsDate(.FechaInicioContrato) Then
            Me.FechaInicioContrato = .FechaInicioContrato
        End If
        
        If IsDate(.FechaFinContrato) Then
            Me.FechaFinContrato = .FechaFinContrato
        End If
        If Not .OrganoContratacion Is Nothing Then
            Me.OrganoContratacion = .OrganoContratacion.OrganoContratacion
        End If
        If Not .ResponsableCalidad Is Nothing Then
            Me.ResponsableCalidad = .ResponsableCalidad.Nombre
        End If
        If Not m_ObjProyectoAlInicio Is Nothing Then
            If m_ObjProyectoAlInicio.EnUTE <> "" Then
                Me.EnUTE = m_ObjProyectoAlInicio.EnUTE
            Else
                Me.EnUTE = Null
            End If
        Else
            If .EnUTE = EnumSiNo.Sí Then
                Me.EnUTE = "Sí"
            Else
                Me.EnUTE = "No"
            End If
            
        End If
        
        If Not m_ObjProyectoAlInicio Is Nothing Then
            If m_ObjProyectoAlInicio.CodigoDocumento <> "" Then
                Me.CodigoDocumento = m_ObjProyectoAlInicio.CodigoDocumento
            Else
                Me.CodigoDocumento = Null
            End If
        Else
            If .CodigoDocParteExpediente <> "" Then
                Me.CodigoDocumento = .CodigoDocParteExpediente
            Else
                Me.CodigoDocumento = Null
            End If
        End If
        Me.CorreoRAC = .CadenaCorreoRACs
        
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosExpediente ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function
Public Function RellenarDatosAlObjeto(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    If m_ObjProyectoActivo Is Nothing Then
        Exit Function
    End If
    With m_ObjProyectoActivo
        .IDExpediente = Nz(Me.IDExpediente, "")
        .CodigoDocumento = Nz(Me.CodigoDocumento, "")
        .ParaInformeAvisos = Nz(Me.ParaInformeAvisos, "")
        .EnUTE = Nz(Me.EnUTE, "")
        .CorreoRAC = Nz(Me.CorreoRAC, "")
        .Elaborado = Nz(Me.Elaborado, "")
        .Revisado = Nz(Me.Revisado, "")
        .Aprobado = Nz(Me.Aprobado, "")
        
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosAlObjeto ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    Me.PalabraClave.Enabled = True
    Me.PalabraClave.Locked = False
    Me.ComandoIrABusquedaExpedientes.Enabled = True
    Me.ComandoBuscarExpediente.Enabled = True
    If m_EsAlta = EnumSiNo.No Then
        If m_ObjProyectoAlInicio Is Nothing Then
            Set m_ObjProyectoAlInicio = Constructor.getProyecto(m_ObjProyectoActivo.IDProyecto, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_ObjProyectoAlInicio Is Nothing Then
                p_Error = "No se conoce la gestión de riesgos"
                Err.Raise 1000
            End If
        End If
    End If
    
    
    
    If Not m_ObjProyectoAlInicio Is Nothing Then
        If Not m_ObjProyectoAlInicio.EdicionActiva Is Nothing Then
            If Not m_ObjProyectoAlInicio.EdicionActiva.FechaPublicacion <> "" Then
                Me.PalabraClave.Enabled = False
                Me.PalabraClave.Locked = True
                Me.ComandoIrABusquedaExpedientes.Enabled = False
                Me.ComandoBuscarExpediente.Enabled = False
            End If
        End If
    End If
    If Nz(Me.IDExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    If m_EsAlta = Empty Then
        p_Error = "No se sabe si es para alta o no"
        Err.Raise 1000
    End If
    Me.Elaborado = ""
    Me.Revisado = ""
    Me.Aprobado = ""
    EstablecerCombos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Form_FormGestionRiesgos.blnPermitidoEditar = False Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    If m_ObjProyectoActivo Is Nothing Then
        RellenarBordeCamposObligatorios Me
        Exit Function
    End If
    With m_ObjProyectoActivo
        
        RellenarDatosExpediente .Expediente, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If

        If IsDate(.FechaCierre) Then
            Me.FechaCierre = .FechaCierre
        End If

        
        If .ParaInformeAvisos = "" Then
            .ParaInformeAvisos = "Sí"
        End If
        
        Me.ParaInformeAvisos = .ParaInformeAvisos
        

        
        If .Elaborado <> "" Then
            Me.Elaborado = .Elaborado
        Else
            Me.Elaborado = Null
        End If
        If .Revisado <> "" Then
            Me.Revisado = .Revisado
        Else
            Me.Revisado = Null
        End If
        If .Aprobado <> "" Then
            Me.Aprobado = .Aprobado
        Else
            Me.Aprobado = Null
        End If
        
    End With
    
    RellenarBordeCamposObligatorios Me
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
End Function

Private Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
    Dim m_UsuarioRed As Variant
    Dim m_ObjUsuario As Usuario
    Dim m_ColUsuariosTecnicos As Scripting.Dictionary
   
    Dim m_SoloActivos As EnumSiNo
    On Error GoTo errores
    
    
    
    
    
    Me.ParaInformeAvisos.RowSource = "Sí;No"
    Me.ParaInformeAvisos.Requery
    
   
    
    
    Me.Elaborado.RowSource = ""
    Me.Elaborado.Requery
    Me.Revisado.RowSource = ""
    Me.Revisado.Requery
    Me.Aprobado.RowSource = ""
    Me.Aprobado.Requery
    If m_ObjProyectoAlInicio Is Nothing Then
        m_SoloActivos = EnumSiNo.Sí
    Else
        m_SoloActivos = EnumSiNo.No
    End If
    
    EstablecerComboTecnicos Me.Elaborado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboTecnicos Me.Revisado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerComboTecnicos Me.Aprobado, m_SoloActivos, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
        
    End If
    
    
End Function



Private Function EstablecerComboTecnicos( _
                                            ByRef cmb As ComboBox, _
                                            Optional ByRef p_SoloActivos As EnumSiNo = EnumSiNo.Sí, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_objColUsuarios As Scripting.Dictionary
    Dim m_UsuarioRed  As Variant
    Dim m_ObjUsuario As Usuario
    
    On Error GoTo errores
    cmb.RowSource = ""
    If p_SoloActivos = Empty Then
        p_SoloActivos = EnumSiNo.Sí
    End If
    If p_SoloActivos = EnumSiNo.Sí Then
        Set m_objColUsuarios = m_ObjEntorno.ColUsuariosTecnicosActivos
    Else
        Set m_objColUsuarios = m_ObjEntorno.ColUsuariosTecnicos
    End If
    If m_objColUsuarios Is Nothing Then
        Exit Function
    End If
    
    For Each m_UsuarioRed In m_objColUsuarios
        Set m_ObjUsuario = m_objColUsuarios(m_UsuarioRed)
        cmb.AddItem m_ObjUsuario.Nombre
        Set m_ObjUsuario = Nothing
siguiente:
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboTecnicos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function





