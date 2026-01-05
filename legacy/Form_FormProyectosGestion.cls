Option Compare Database
Option Explicit



Public WithEvents m_ObjProyectoSeleccionado As Proyecto
Private m_ObjColProyectosFiltrados As Scripting.Dictionary

Private m_Error As String


Private m_ObjEdicionUltimaNoActiva As Edicion
Private WithEvents m_FormAnexos As Form_FormAnexos
Private WithEvents m_FormGestionRiesgos As Form_FormGestionRiesgos
Private m_URLInformeUltimaEdicionProyectoSeleccionado As String

Private Sub EstablecerBotoneraAlInicio(Optional ByRef p_Error As String)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoAnexar.Enabled = False
    Me.ComandoVerEvolucionRiesgos.Enabled = False
    Me.ImagenDocumentoUltimPublicacion.Visible = False
    Me.ComandoEditarProyecto.Enabled = False
    
    Me.ComandoEliminar.Enabled = False
   
    
    
    Me.ComandoRiesgos.Enabled = False
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.ComandoNuevoProyecto.Enabled = False
    Else
        Me.ComandoNuevoProyecto.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al EstablecerBotoneraAlInicio se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub


Private Function Filtrar( _
                        Optional p_RellenandoColeccion As EnumSiNo = EnumSiNo.Sí, _
                        Optional p_Error As String _
                        ) As String
    
    
    
    Dim m_Col As Scripting.Dictionary
    Dim m_fechaProx As String
    Dim m_IDProyecto As Variant
    Dim m_Proyecto As Proyecto
    Dim m_FechaUltimaPublicacion As String
    Dim lst As ListBox
    
    On Error GoTo errores
    
    m_Error = ""
    
   
    
    Set m_ObjProyectoSeleccionado = Nothing
    Set m_ObjProyectoActivo = Nothing
    EstablecerBotoneraAlInicio p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set lst = Me.ListaFiltrados
    lst.RowSourceType = "Lista de Valores"
    lst.RowSource = "ID;Cod.Exp;Nemotécnico;Fecha prev. cierre;Fecha Cierre;F.Ultima Pub.;Prox Pub.;Rep. Calidad"
    'lst.Requery
    Set m_Col = Constructor.getProyectosBusqueda1(Nz(Me.comboActivos, "Todos"), _
                            Nz(Me.ComboJefeProyecto.Column(1), ""), _
                            Nz(Me.PalabraClave, ""), _
                            Nz(Me.ComboResponsableCalidad, ""), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_IDProyecto In m_Col
        Set m_Proyecto = m_Col(m_IDProyecto)
        With m_Proyecto
            If Me.comboActivos = "No" Then
                m_fechaProx = "-----"
            Else
                m_fechaProx = .FechaMaxProximaPublicacion
                If Not IsDate(m_fechaProx) Then
                    m_fechaProx = "-----"
                End If
            End If
            m_FechaUltimaPublicacion = .FechaUltimaPublicacion
            If m_FechaUltimaPublicacion = "" Then
                m_FechaUltimaPublicacion = "-----"
            End If
              
            
            lst.AddItem .IDProyecto & ";" & .Proyecto & ";" & .NombreProyecto & ";" & _
                    .FechaPrevistaCierre & ";" & .FechaCierre & ";" & m_FechaUltimaPublicacion & ";" & m_fechaProx & ";" & .NombreUsuarioCalidad
        End With
        Set m_Proyecto = Nothing
    
siguienteProyecto:
    Next
    lst.Selected(1) = True
    ListaFiltrados_Click
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Filtrar ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ComandoActualizar_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjProyectoActivo = Nothing
    
    Me.ComandoEliminar.Enabled = False
    Me.ComandoEditarProyecto.Caption = "Editar datos principales"
    
    Me.ComandoEditarProyecto.Enabled = False
    Set m_ObjEntorno.ColProyectosTotales = Nothing
    Filtrar EnumSiNo.Sí, m_Error
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

Private Sub ComandoEditarProyecto_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjProyectoActivo = m_ObjProyectoSeleccionado
    Set m_ObjProyectoAlInicio = Nothing
    m_EsAlta = EnumSiNo.No
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditarProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente borrar completamente la gestión de riesgos seleccionada?", _
                vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar gestión de riesgos")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set m_ObjProyectoActivo = m_ObjProyectoSeleccionado
    m_ObjProyectoActivo.Borrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo = Nothing
    Set m_ObjProyectoSeleccionado = Nothing
    EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Filtrar , m_Error
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

Private Sub m_FormGestionRiesgos_Alta()
    ComandoFiltrar_Click
    
End Sub

Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.ComandoLimpiarPalabraClave.SetFocus
        Me.PalabraClave.SetFocus
        ComandoFiltrar_Click
        KeyCode = 0
    End If
End Sub
Public Sub ComandoFiltrar_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Filtrar , m_Error
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
        m_Error = "Al ComandoFiltrar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoLimpiarActivos_Click()
    If Nz(Me.comboActivos, "") = "" Then
        Exit Sub
    End If
    Me.comboActivos = Null
    ComandoFiltrar_Click
End Sub

Private Sub ComandoLimpiarAutorizado_Click()
    If Nz(Me.ComboJefeProyecto, "") = "" Then
        Exit Sub
    End If
    Me.ComboJefeProyecto = Null
    ComandoFiltrar_Click
End Sub

Private Sub ComandoLimpiarPalabraClave_Click()
    If Nz(Me.PalabraClave, "") = "" Then
        Exit Sub
    End If
    Me.PalabraClave = Null
    ComandoFiltrar_Click
End Sub

Private Sub ComandoLimpiarResponsableCalidad_Click()
     If Nz(Me.ComboResponsableCalidad, "") = "" Then
        Exit Sub
    End If
    Me.ComboResponsableCalidad = Null
    ComandoFiltrar_Click
End Sub

Private Sub ComandoNuevoProyecto_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    If EsTecnico = EnumSiNo.Sí Then
        m_Error = "Sólo los miembros de Calidad pueden generar el alta de una gestión de riesgos" & vbNewLine & _
            "Mire a ver si ya la ha dado de alta"
        Err.Raise 1000
    End If
    Set m_ObjProyectoActivo = Nothing
    Set m_ObjProyectoAlInicio = Nothing
    m_EsAlta = EnumSiNo.Sí
    
    
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    If FormularioAbierto("FormGestionRiesgos") Then
        Set m_FormGestionRiesgos = Forms("FormGestionRiesgos")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevoProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoRiesgos_Click()
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    End If
    Set m_ObjProyectoActivo = m_ObjProyectoSeleccionado
    
    If m_ObjProyectoSeleccionado.EdicionActiva Is Nothing Then
        Set m_ObjEdicionActiva = m_ObjProyectoSeleccionado.EdicionUltima
    Else
        Set m_ObjEdicionActiva = m_ObjProyectoSeleccionado.EdicionActiva
    End If
    
    If FormularioAbierto("FormRiesgosGestion") Then
        DoCmd.Close acForm, "FormRiesgosGestion", acSaveNo
    End If
    
    DoCmd.OpenForm "FormRiesgosGestion"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoVerEvolucionRiesgos_Click()
     Dim m_URLExcel As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_URLExcel = m_ObjProyectoSeleccionado.ObtenerGraficasRiesgos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    If FSO.FileExists(m_URLExcel) Then
        Ejecutar Me.hWnd, "open", m_URLExcel, "", "", 1
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerEvolucionRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Function RellenaComboJps( _
                                    ByRef cmb As ComboBox, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    Dim m_ColUsuarios As Scripting.Dictionary
    Dim m_Usuario As Usuario
    Dim m_UsuarioRed As Variant
    Dim m_UsuarioLargo As String
    On Error GoTo errores
    
    cmb.RowSource = ""
    
    Set m_ColUsuarios = m_ObjEntorno.ColJPs
    
    If Not m_ColUsuarios Is Nothing Then
        For Each m_UsuarioRed In m_ColUsuarios.keys
           Set m_Usuario = m_ColUsuarios(m_UsuarioRed)
            m_UsuarioLargo = m_Usuario.Nombre
            If cmb.ColumnCount = 1 Then
                cmb.AddItem m_UsuarioLargo
            Else
                cmb.AddItem m_UsuarioRed & ";" & m_UsuarioLargo
            End If
            Set m_Usuario = Nothing
        Next
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenaComboJps ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub comboActivos_BeforeUpdate(Cancel As Integer)
    ComandoFiltrar_Click
End Sub

Private Sub ComboJefeProyecto_BeforeUpdate(Cancel As Integer)
    ComandoFiltrar_Click
End Sub

Private Sub ComboResponsableCalidad_BeforeUpdate(Cancel As Integer)
    ComandoFiltrar_Click
End Sub

Private Sub Form_Load()
    Dim lst As ListBox
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    AjustarTamaño Me

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.Caption = Me.lblTitulo1.Caption
    Set lst = Me.ListaFiltrados
    lst.RowSourceType = "Lista de Valores"
    lst.RowSource = "ID;Cod.Exp;Nemotécnico;Jurídica;Fecha Cierre;F Prevista Cierre;Prox Pub."
    VBA.DoEvents
    Avance "Estableciendo datos de usuarios...."
    
    VBA.DoEvents
    m_Error = ""
    t1 = Timer
    RellenaComboJps Me.ComboJefeProyecto, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    t2 = Timer
    VBA.DoEvents
    Debug.Print "RellenaComboJps", t2 - t1
    VBA.DoEvents
    t1 = Timer
    EstablecerComboUsuariosCalidad Me.ComboResponsableCalidad, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    t2 = Timer
    Debug.Print "EstablecerComboUsuariosCalidad", t2 - t1
    Me.comboActivos = "Sí"
    
    If EsTecnico = EnumSiNo.Sí Then
        Me.ComboJefeProyecto = m_ObjUsuarioConectado.UsuarioRed
        Me.ComandoEditarProyecto.Caption = "Ver Datos Principales"
        Me.ComandoEliminar.Visible = False
    Else
        Me.ComandoEditarProyecto.Caption = "Editar Datos Principales"
        Me.ComandoEliminar.Visible = True
    End If
    If EsCalidad = EnumSiNo.Sí Then
        Me.ComboResponsableCalidad = m_ObjUsuarioConectado.Nombre
    End If
    Filtrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
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
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ImagenDocumentoUltimPublicacion_Click()
     On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_URLInformeUltimaEdicionProyectoSeleccionado <> "" Then
        AbrirEnLocal m_URLInformeUltimaEdicionProyectoSeleccionado, Me.hWnd, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenDocumentoUltimPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Sub ListaFiltrados_Click()
    Dim m_IDProyectoSeleccionado As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Application.Echo False
    Me.lblUltimoArchivoPublicado.Visible = False
    Me.ImagenDocumentoUltimPublicacion.Visible = False
    m_URLInformeUltimaEdicionProyectoSeleccionado = ""
    If Me.ListaFiltrados.ListCount = 2 Then
        Me.ListaFiltrados.Selected(1) = True
     
        
    End If
    m_IDProyectoSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    If m_IDProyectoSeleccionado = "" Then
        Set m_ObjProyectoSeleccionado = Nothing
        Set m_ObjEdicionActiva = Nothing
        Set m_ObjRiesgoActivo = Nothing
        Set m_ObjPMActivo = Nothing
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjPCActivo = Nothing
        Set m_ObjPCAccionActiva = Nothing
        EstablecerBotoneraAlInicio m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Application.Echo True
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(m_IDProyectoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjProyectoSeleccionado.IDProyecto <> m_IDProyectoSeleccionado Then
            Set m_ObjProyectoSeleccionado = Constructor.getProyecto(m_IDProyectoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Application.Echo True
            Exit Sub
        End If
    End If
    EstablecerBotonera m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyectoSeleccionado Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Application.Echo True
        Exit Sub
    End If
    With m_ObjProyectoSeleccionado
        If Not .EdicionUltimaPublicada Is Nothing Then
            m_Error = .Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_URLInformeUltimaEdicionProyectoSeleccionado = .EdicionUltimaPublicada.URLArchivoInforme
            m_Error = .EdicionUltimaPublicada.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End With
    
    
    If m_URLInformeUltimaEdicionProyectoSeleccionado <> "" Then
        Me.lblUltimoArchivoPublicado.Visible = True
        Me.ImagenDocumentoUltimPublicacion.Visible = True
    Else
        Me.lblUltimoArchivoPublicado.Visible = False
        Me.ImagenDocumentoUltimPublicacion.Visible = False
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Application.Echo True
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaFiltrados_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    Application.Echo True
End Sub

Private Sub ListaFiltrados_DblClick(Cancel As Integer)
    On Error Resume Next
    Application.Echo False
    If Me.ComandoRiesgos.Enabled = True Then
        ComandoRiesgos_Click
    End If
    Application.Echo True
End Sub

Public Function EstablecerBotonera(Optional ByRef p_Error As String) As String
   
    
    On Error GoTo errores
    
    p_Error = ""
    
    EstablecerBotoneraAlInicio p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            Exit Function
        End If
    End If
    Me.ComandoAnexar.Enabled = True
   
   
    Me.ComandoRiesgos.Enabled = True
    
    Me.ComandoEditarProyecto.Enabled = True
    
    
     
    Me.ComandoVerEvolucionRiesgos.Enabled = True
    If EsTecnico <> EnumSiNo.Sí Then
        Me.ComandoEliminar.Enabled = True
    Else
        Me.ComandoEliminar.Enabled = False
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerBotonera ha devuelto el error: " & Err.Description
    End If
End Function
Private Sub m_FormAnexos_AnexoAñadido(ByVal p_ObjAnexo As Anexo)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    Else
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(m_ObjProyectoSeleccionado.IDProyecto, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormAnexos_AnexoAñadido se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub m_FormAnexos_AnexoEliminado(ByVal IDAnexo As String)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjProyectoSeleccionado Is Nothing Then
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(Nz(Me.ListaFiltrados.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjProyectoSeleccionado Is Nothing Then
            m_Error = "Se ha de seleccionar una gestión de riesgos de la lista"
            Err.Raise 1000
        End If
    Else
        Set m_ObjProyectoSeleccionado = Constructor.getProyecto(m_ObjProyectoSeleccionado.IDProyecto, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormAnexos_AnexoEliminado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function ActualizarProyecto( _
                                    Optional p_ObjProyecto As Proyecto, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim lst As ListBox
    Dim i As Integer
    Dim m_LineaActual As String
    Dim m_LineaFinal As String
    
    On Error GoTo errores
    
    p_Error = ""
    m_Error = ""
    If p_ObjProyecto Is Nothing Then
        Exit Function
    End If
    Set p_ObjProyecto = Constructor.getProyecto(p_ObjProyecto.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set lst = Me.ListaFiltrados
    For i = 1 To lst.ListCount - 1
        If Nz(lst.Column(0, i), "") = p_ObjProyecto.IDProyecto Then
            m_LineaActual = Nz(lst.Column(0, i), "") & ";" & Nz(lst.Column(1, i), "") & ";" & Nz(lst.Column(2, i), "") & ";" & _
                        Nz(lst.Column(3, i), "") & ";" & Nz(lst.Column(4, i), "") & ";" & _
                        Nz(lst.Column(5, i), "")
            lst.Selected(i) = True
        End If
    Next
    With p_ObjProyecto
        m_LineaFinal = .IDProyecto & ";" & .Proyecto & ";" & .NombreProyecto & ";" & .Juridica & ";" & _
                            .FechaCierre & ";" & .FechaPrevistaCierreCalculada
    End With
    If m_LineaActual <> m_LineaFinal Then
        lst.RowSource = Replace(lst.RowSource, m_LineaActual, m_LineaFinal)
        lst.Requery
        ListaFiltrados_Click
    Else
        ListaFiltrados_Click
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarProyecto ha producido el error: " & vbNewLine & Err.Description
    End If
End Function

