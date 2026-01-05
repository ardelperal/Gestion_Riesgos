VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_formIndicadorProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Error As String
Private flag As String
Private dato

Private m_Anio As Long
Private m_Semestre As String  ' "1" / "2" / "" (si anual lo decides)

Private Sub Form_Close()
    On Error Resume Next
    IndicadorState_MarcarFormAbierto False
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    IndicadorState_Init
    IndicadorState_MarcarFormAbierto True
    
   

    ParseOpenArgs
    CargarLista_RespetandoExcluidos
    Me.cmdQuitar.Enabled = False
    Exit Sub
errores:
    MsgBox "Form_Open (formIndicadorProyectos): " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub CargarLista_RespetandoExcluidos()
    Dim lst As ListBox
    Dim strIDProyecto As String
    Dim strProyecto As String
    Dim strNombreProyecto As String
    Dim strFechaCierre As String
    Dim strFechaAlta As String
    Dim flag As String
    Dim dato
    Dim varItem As Variant
    Dim ColProyectosAbiertos As Collection
    On Error GoTo errores

    Set lst = Me.lstProyectos

    ' Encabezado (fila 0)
    lst.RowSource = "IDProyecto;Código;Proyecto;Cierre;Alta"
    lst.Requery

    ' Si no llegó año, fallo claro
    If m_Anio <= 0 Then Err.Raise 1000, , "Falta el año (OpenArgs ANIO=...)."

    ' Relleno ColProyectosAbiertos (tu función)
   
    Set ColProyectosAbiertos = getProyectosAbiertos(CStr(m_Anio), m_Semestre, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    If ColProyectosAbiertos Is Nothing Then
        Exit Sub
    End If
    ' Añadir SOLO los NO excluidos
    If ColProyectosAbiertos.Count > 0 Then
        For Each varItem In ColProyectosAbiertos
            dato = Split(varItem, "|")
            strIDProyecto = Nz(dato(0), "")
            strProyecto = Replace(Nz(dato(1), ""), ";", ":")
            strNombreProyecto = Replace(Nz(dato(2), ""), ";", ":")
            strFechaCierre = Nz(dato(3), "")
            strFechaAlta = Nz(dato(4), "") ' FechaAltaRegistro

            If IsNumeric(strIDProyecto) Then
                If Not IndicadorState_EsExcluido(CLng(strIDProyecto)) Then
                    If IsDate(strFechaCierre) Then strFechaCierre = Format(strFechaCierre, "dd/mm/yyyy")
                    If IsDate(strFechaAlta) Then strFechaAlta = Format(strFechaAlta, "dd/mm/yyyy")
                    lst.AddItem strIDProyecto & ";" & strProyecto & ";" & strNombreProyecto & ";" & strFechaCierre & ";" & strFechaAlta
                End If
            End If
        Next
    End If

    Exit Sub

errores:
    MsgBox "CargarLista_RespetandoExcluidos: " & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub ParseOpenArgs()
    ' Espera OpenArgs: "2025|S1" o "2025|S2" o "2025|Anual"
    Dim s As String, a() As String, sem As String

    s = Nz(Me.openArgs, "")
    If InStr(1, s, "|") > 0 Then
        a = Split(s, "|")
        If UBound(a) >= 1 Then
            If IsNumeric(a(0)) Then m_Anio = CLng(a(0))
            sem = UCase$(Trim$(a(1)))
        End If
    End If

    If m_Anio = 0 Then m_Anio = CLng(Year(Date))

    ' Adaptación: tu función legacy espera "1" / "2" (o vacío)
    Select Case sem
        Case "S1", "1"
            m_Semestre = "1"
        Case "S2", "2"
            m_Semestre = "2"
        Case Else
            m_Semestre = ""  ' anual (si lo usas)
    End Select
End Sub



Private Sub lstProyectos_Click()
    On Error Resume Next
    Me.cmdQuitar.Enabled = (SelectedIDProyecto(Me.lstProyectos) > 0)
    
End Sub

Private Function SelectedIDProyecto(ByVal lst As ListBox) As Long
    Dim v As Variant

    SelectedIDProyecto = 0
    If lst.ListIndex < 0 Then Exit Function

    ' Columna 0 = IDProyecto (da igual si hay cabecera falsa o ColumnHeads)
    v = Nz(lst.Column(0, lst.ListIndex), "")

    If IsNumeric(v) Then
        SelectedIDProyecto = CLng(v)
    End If
End Function


Private Sub cmdQuitar_Click()
    Dim lst As ListBox
    Dim idp As Long

    On Error GoTo errores
    Set lst = Me.lstProyectos

    ' Si está en la cabecera (fila 0) o no hay selección, no se quita nada
    If lst.ListIndex <= 0 Then
        MsgBox "Selecciona un proyecto (no la cabecera).", vbExclamation
        Me.cmdQuitar.Enabled = False
        Exit Sub
    End If

    ' ID real del proyecto (columna 0 de la fila seleccionada)
    idp = CLng(lst.Column(0, lst.ListIndex))

    IndicadorState_SetExcluido idp
    lst.RemoveItem lst.ListIndex

    Me.cmdQuitar.Enabled = False
    Exit Sub

errores:
    MsgBox "Quitar: " & vbCrLf & Err.Description, vbCritical
End Sub



Private Sub cmdRecargarTodos_Click()
    On Error GoTo errores
    IndicadorState_ResetExcluidos
    Call CargarLista_RespetandoExcluidos
    Me.cmdQuitar.Enabled = False
    Exit Sub

errores:
    MsgBox "Recargar todos: " & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub cmdAceptar_Click()
    On Error GoTo errores

    IndicadorState_Init
    IndicadorState_SetSeleccionadosDesdeLista Me.lstProyectos
    IndicadorState_MarcarConfirmado True

    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    MsgBox "Aceptar: " & vbCrLf & Err.Description, vbCritical
End Sub



Private Sub cmdCancelar_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function getProyectosAbiertos( _
                                        p_Anio As String, _
                                        Optional ByVal p_Semestre As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Collection
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Dim dIni As Date
    Dim dFinExcl As Date
    
    Dim idp As Long
    Dim Proyecto As String
    Dim NombreProyecto As String
    Dim FechaCierre As Variant
    Dim fechaRegistroInicial As Variant
    
    On Error GoTo errores
    p_Error = ""

    If Not IsNumeric(p_Anio) Then
        p_Error = "El año es obligatorio"
        Err.Raise 1000
    End If
    
    Select Case Nz(p_Semestre, "")
        Case "1"
            dIni = DateSerial(CLng(p_Anio), 1, 1)
            dFinExcl = DateSerial(CLng(p_Anio), 7, 1)          ' 01/07 (exclusivo)
        Case "2"
            dIni = DateSerial(CLng(p_Anio), 7, 1)
            dFinExcl = DateSerial(CLng(p_Anio) + 1, 1, 1)      ' 01/01 del año siguiente (exclusivo)
        Case ""   ' anual
            dIni = DateSerial(CLng(p_Anio), 1, 1)
            dFinExcl = DateSerial(CLng(p_Anio) + 1, 1, 1)
        Case Else
            p_Error = "Sólo hay dos semestres en un año"
            Err.Raise 1000
    End Select

    ' Proyectos “abiertos”/activos en el periodo por solape de intervalos:
    ' (InicioProyecto < FinPeriodoExcl) AND (FinProyecto Is Null OR FinProyecto >= InicioPeriodo)
    sql = _
        "PARAMETERS pIni DateTime, pFinExcl DateTime;" & vbCrLf & _
        "SELECT IDProyecto, Proyecto, NombreProyecto, FechaCierre, FechaRegistroInicial" & vbCrLf & _
        "FROM TbProyectos" & vbCrLf & _
        "WHERE (FechaRegistroInicial Is Null OR FechaRegistroInicial < [pFinExcl])" & vbCrLf & _
        "  AND (FechaCierre Is Null OR FechaCierre >= [pIni])" & vbCrLf & _
        "ORDER BY IDProyecto;"

    Set db = getdb()
    Set qd = db.CreateQueryDef("", sql)
    qd.Parameters("pIni").Value = dIni
    qd.Parameters("pFinExcl").Value = dFinExcl
    
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        Set getProyectosAbiertos = New Collection
        Do While Not rs.EOF
            idp = CLng(Nz(rs!IDProyecto, 0))
            Proyecto = Nz(rs!Proyecto, "")
            NombreProyecto = Nz(rs!NombreProyecto, "")
            FechaCierre = rs!FechaCierre
            fechaRegistroInicial = rs!fechaRegistroInicial
            
            getProyectosAbiertos.Add _
                CStr(idp) & "|" & Proyecto & "|" & NombreProyecto & "|" & Nz(FechaCierre, "") & "|" & Nz(fechaRegistroInicial, "")
            
            rs.MoveNext
        Loop
    End If

salir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qd = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "getProyectosAbiertos: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Resume salir
End Function

