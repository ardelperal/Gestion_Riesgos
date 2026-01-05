VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormRiesgosEstablecerPrioridades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_RiesgoSeleccionado As riesgo
Private blnHaHabidoPriorizaciones As Boolean
Private blnPermitidoEscribir As Boolean
Private lngNumeroRiesgosActivos As Long
Private m_ColPriorizacionAlInicio As Scripting.Dictionary
Private m_ColPriorizacionAlFinal As Scripting.Dictionary
Private m_ProyConRiesgosBiblioteca As EnumSiNo
Private m_Error As String

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    EstablecerLista EnumSiNo.Sí, m_Error
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

Private Function SeleccionarSiguiente(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    If lst.ListCount < 1 Then
        Exit Function
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        lst.Selected(1) = True
        ListaRiesgos_Click
        Exit Function
    End If
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = m_RiesgoSeleccionado.IDRiesgo Then
            If i = lst.ListCount - 1 Then
                lst.Selected(1) = True
            Else
                lst.Selected(i + 1) = True
            End If
            
            ListaRiesgos_Click
            Exit Function
        End If
        
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarSiguiente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function SeleccionarAnterior(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    If lst.ListCount < 1 Then
        Exit Function
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        lst.Selected(1) = True
        ListaRiesgos_Click
        Exit Function
    End If
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = m_RiesgoSeleccionado.IDRiesgo Then
            If i = 1 Then
                lst.Selected(lst.ListCount - 1) = True
            Else
                lst.Selected(i - 1) = True
            End If
            
            ListaRiesgos_Click
            Exit Function
        End If
        
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ComandoAsignar_Click()
    Dim lst As ListBox
    Dim m_SQL As String
    Dim m_Priorizacion As String
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set lst = Me.ListaRiesgos
    If lst.ItemsSelected.Count = 0 Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    m_Priorizacion = Nz(Me.Priorizacion, "")
    If Not IsNumeric(m_Priorizacion) Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    If CLng(m_Priorizacion) < 1 Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    If CLng(m_Priorizacion) > lngNumeroRiesgosActivos Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbAuxPriorizacion " & _
            "SET Pri = '" & m_Priorizacion & "' " & _
            "WHERE IDRiesgo=" & lst.Column(0) & ";"
    CurrentDb().Execute m_SQL
    lst.Requery
    On Error Resume Next
    SeleccionarSiguiente
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAsignar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

Private Sub ComandoGrabarPriorizacion_Click()
   
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    EstablecerPriorizaciones m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabarPriorizacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComandoSalir_Click()
    On Error Resume Next
    If blnHaHabidoPriorizaciones = True Then
        'RegistrarActualizacionRiesgo p_Edicion:=m_ObjEdicionActiva
        RegistrarActualizacionEdicion p_Edicion:=m_ObjEdicionActiva
        
        Form_FormRiesgosGestion.CargarArbol p_Refrescando:=EnumSiNo.No
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
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


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
  

    On Error GoTo errores
    Me.ComandoGrabarPriorizacion.Enabled = False
    Me.ComandoAsignar.Enabled = False
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se conoce la edición activa"
        Err.Raise 1000
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        If m_ObjEdicionActiva.UsuarioConectadoAutorizado = EnumSiNo.Sí Then
            blnPermitidoEscribir = True
            Me.ComandoGrabarPriorizacion.Enabled = True
        Else
            blnPermitidoEscribir = False
        End If
    Else
        blnPermitidoEscribir = False
    End If
    If m_ObjEdicionActiva.colRiesgosActivos Is Nothing Then
        p_Error = "No hay riesgos no retirados"
        Err.Raise 1000
    End If
    lngNumeroRiesgosActivos = m_ObjEdicionActiva.colRiesgosActivos.Count
    m_ProyConRiesgosBiblioteca = m_ObjEdicionActiva.Proyecto.RequiereRiesgoDeBibliotecaCalculado
    Me.lblInfoOrdinal.Caption = "Introduzca un valor entre " & "1" & " " & "y" & " " & lngNumeroRiesgosActivos & " " & "para la priorización de este riesgo"
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    EstablecerLista EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColPriorizacionAlInicio = getColDeTabla(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function EstablecerLista( _
                                Optional ByRef p_Actualizando As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim lst As ListBox
    
    Dim m_SQL As String
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    
    If m_ProyConRiesgosBiblioteca = EnumSiNo.Sí Then
        lst.ColumnCount = 6
        lst.ColumnWidths = "0cm;1,503cm;3cm;2cm;10cm;16cm"
        m_SQL = "SELECT IDRiesgo, Pri, PriEdAnterior, Codigo, Descripcion, CausaRaiz " & _
                "FROM TbAuxPriorizacion;"
        lst.RowSource = "IDRiesgo;Pri.;Pri.Ed.Anterior;Código;Descripción;CausaRaiz"
    Else
        lst.ColumnCount = 5
        lst.ColumnWidths = "0cm;1,503cm;3cm;2cm;25cm"
        m_SQL = "SELECT IDRiesgo, Pri, PriEdAnterior , Codigo, Descripcion " & _
                "FROM TbAuxPriorizacion;"
        
    End If
    If p_Actualizando = EnumSiNo.Sí Then
        Set m_ObjEdicionActiva = Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion)
    End If
    RellenarTbAux p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    lst.RowSource = m_SQL
    lst.Requery
    
    lst.Selected(1) = True
    ListaRiesgos_Click
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Function getColDeTabla( _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Id As String
    Dim m_Pri As String
    
    On Error GoTo errores
    
    m_SQL = "TbAuxPriorizacion"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If Not rcdDatos.EOF Then
        rcdDatos.MoveFirst
        Do While Not rcdDatos.EOF
            m_Id = Nz(rcdDatos.Fields("IDRiesgo"), "")
            m_Pri = Nz(rcdDatos.Fields("Pri"), "")
            If getColDeTabla Is Nothing Then
                Set getColDeTabla = New Scripting.Dictionary
                getColDeTabla.CompareMode = TextCompare
            End If
            getColDeTabla.Add CStr(m_Id), m_Pri
            rcdDatos.MoveNext
        Loop
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColDeTabla ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function



Private Function SeleccionarRiesgo(p_IDRiesgo As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    Set lst = Me.ListaRiesgos
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = p_IDRiesgo Then
            lst.Selected(i) = True
            ListaRiesgos_Click
            Exit Function
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgo ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Public Function VaciarTbAux(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    
    
    On Error GoTo errores
    m_SQL = "DELETE * " & _
            "FROM TbAuxPriorizacion;"
    CurrentDb().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método VaciarTbAux ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function RellenarTbAux(Optional ByRef p_Error As String) As String
    
    
    Dim m_SQL As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Riesgo As riesgo
    Dim m_Pri As String
    Dim m_PriEdAnterior As String
    Dim m_Descripcion As String
    Dim m_CausaRaiz As String
    
    On Error GoTo errores
    VaciarTbAux p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Col = m_ObjEdicionActiva.colRiesgosActivos
    If m_Col Is Nothing Then
        Exit Function
    End If
    
    For Each m_Id In m_Col
        Set m_Riesgo = m_Col(m_Id)
        m_Descripcion = m_Riesgo.Descripcion
        
        m_Pri = m_Riesgo.Priorizacion
        If Not IsNumeric(m_Pri) Then m_Pri = "-"
        m_PriEdAnterior = m_Riesgo.PriorizacionEdAnterior
        
        
        If m_ProyConRiesgosBiblioteca = EnumSiNo.Sí Then
            m_CausaRaiz = m_Riesgo.CausaRaiz
            m_SQL = "INSERT INTO TbAuxPriorizacion " & _
                "(IDRiesgo, Pri,PriEdAnterior,Codigo,Descripcion,CausaRaiz) " & _
                "VALUES (" & m_Id & ",'" & m_Pri & "','" & m_PriEdAnterior & "','" & _
                        m_Riesgo.CodigoRiesgo & "','" & m_Descripcion & "','" & m_CausaRaiz & "')"
        Else
            m_SQL = "INSERT INTO TbAuxPriorizacion " & _
                "(IDRiesgo, Pri,PriEdAnterior,Codigo,Descripcion) " & _
                "VALUES (" & m_Id & ",'" & m_Pri & "','" & m_PriEdAnterior & "','" & _
                        m_Riesgo.CodigoRiesgo & "','" & m_Descripcion & "')"
        End If
        CurrentDb.Execute m_SQL
        

        Set m_Riesgo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método VaciarTbAux ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Sub ListaRiesgos_Click()
    
    On Error GoTo errores

    m_Error = ""
    DoCmd.Hourglass True
    Me.Priorizacion = Null
    Me.ComandoAsignar.Enabled = False
    
    Me.Priorizacion = Null
    If Nz(Me.ListaRiesgos.Column(0), "") = "" Then
        Set m_RiesgoSeleccionado = Nothing
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        Set m_RiesgoSeleccionado = Constructor.getRiesgo(Me.ListaRiesgos.Column(0), , , m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_RiesgoSeleccionado.IDRiesgo <> Me.ListaRiesgos.Column(0) Then
            Set m_RiesgoSeleccionado = Constructor.getRiesgo(Me.ListaRiesgos.Column(0), , , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    If IsNumeric(m_RiesgoSeleccionado.Priorizacion) Then
        Me.Priorizacion = m_RiesgoSeleccionado.Priorizacion
    Else
        If blnPermitidoEscribir = True Then
            If IsNumeric(m_RiesgoSeleccionado.PriorizacionEdAnterior) Then
                Me.Priorizacion = m_RiesgoSeleccionado.PriorizacionEdAnterior
            End If
        End If
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoAsignar.Enabled = True
        
    End If
   
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Function EstablecerPriorizaciones(Optional ByRef p_Error As String) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Riesgo As riesgo
    Dim m_PriorizacionFinal As String
    Dim m_PriorizacionInicial As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim blnEnTrans As Boolean
    On Error GoTo errores
    If m_ColPriorizacionAlInicio Is Nothing Then
        p_Error = "No se estableció el listado inicial"
        Err.Raise 1000
    End If
    Set m_ColPriorizacionAlFinal = getColDeTabla(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    For Each m_Id In m_ColPriorizacionAlFinal
        m_PriorizacionFinal = m_ColPriorizacionAlFinal(m_Id)
        If Not IsNumeric(m_PriorizacionFinal) Then
            m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
            Err.Raise 1000
        End If
        If CLng(m_PriorizacionFinal) < 1 Then
            m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
            Err.Raise 1000
        End If
        If CLng(m_PriorizacionFinal) > lngNumeroRiesgosActivos Then
            m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
            Err.Raise 1000
        End If
    Next
    'verificar si hay alguna priorización repetida
    For Each m_Id In m_ColPriorizacionAlFinal
        m_PriorizacionFinal = m_ColPriorizacionAlFinal(m_Id)
        If m_Col Is Nothing Then
            Set m_Col = New Scripting.Dictionary
            m_Col.CompareMode = TextCompare
        End If
        If m_Col.Exists(CStr(m_PriorizacionFinal)) Then
            m_Error = "La priorización " & m_PriorizacionFinal & " está repetida"
            Err.Raise 1000
        End If
        m_Col.Add CStr(m_PriorizacionFinal), m_PriorizacionFinal
    Next
    
    Set ws = DBEngine.Workspaces(0)
    Set db = getdb(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ws.BeginTrans
    blnEnTrans = True
    'grabar priorizaciones
    For Each m_Id In m_ColPriorizacionAlFinal
        m_PriorizacionFinal = m_ColPriorizacionAlFinal(m_Id)
        m_PriorizacionInicial = m_ColPriorizacionAlInicio(CStr(m_Id))
        If m_PriorizacionFinal = m_PriorizacionInicial Then
            GoTo siguiente
        End If
        Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(m_Id), p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Riesgo Is Nothing Then
            GoTo siguiente
        End If
        m_Riesgo.GrabarPrioridad m_PriorizacionFinal, p_db:=db, p_ActualizarOrdenCache:=EnumSiNo.No, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
siguiente:
    Next
    CacheArbolRiesgosTx_ActualizarOrdenRiesgos m_ObjEdicionActiva, m_ColPriorizacionAlFinal, p_db:=db, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000

    End If
    ws.CommitTrans
    blnEnTrans = False
    Form_FormRiesgosGestion.ComandoActualizarContador_Click
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerPriorizaciones ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Sub Priorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
     If Me.ComandoAsignar.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        Me.ComandoAsignar.SetFocus
        Me.Priorizacion.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoAsignar_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    ElseIf KeyCode = 40 Then
        SeleccionarSiguiente
    
    
    ElseIf KeyCode = 38 Then
        SeleccionarAnterior
    End If
End Sub
