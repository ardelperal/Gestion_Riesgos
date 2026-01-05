Attribute VB_Name = "modIndicadorDashboard"
Option Compare Database
Option Explicit

'========================
' CONFIG
'========================
Public Const TVAR_IDSCSV As String = "Indicador_IdsCsv"
Public Const TVAR_COUNT As String = "Indicador_ProyectosCount"
Public Const TVAR_OPENED As String = "Indicador_ProyectosDialogOpened"


Public Enum IndicadorTile
    itIdentificados = 1
    itRetirados = 2
    itEnOferta = 3
    itMaterializados = 4
    itOfertaTrasladar = 5
End Enum

'========================
' RANGO FECHAS
'========================
Public Function CalcularRangoSemestre( _
    ByVal p_Semestre As String, _
    ByVal p_Anio As Long, _
    ByRef p_dIni As Date, _
    ByRef p_dFin As Date, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo errores
    p_Error = ""
    CalcularRangoSemestre = False

    If p_Anio < 1900 Or p_Anio > 2100 Then
        p_Error = "Año inválido."
        Exit Function
    End If

    Select Case UCase$(Trim$(p_Semestre))
        Case "S1", "1"
            p_dIni = DateSerial(p_Anio, 1, 1)
            p_dFin = DateSerial(p_Anio, 6, 30)

        Case "S2", "2"
            p_dIni = DateSerial(p_Anio, 7, 1)
            p_dFin = DateSerial(p_Anio, 12, 31)

        Case "ANUAL", ""
            p_dIni = DateSerial(p_Anio, 1, 1)
            p_dFin = DateSerial(p_Anio, 12, 31)

        Case Else
            p_Error = "Semestre inválido (S1/S2/Anual)."
            Exit Function
    End Select

    CalcularRangoSemestre = True
    Exit Function

errores:
    p_Error = "CalcularRangoSemestre: " & Err.Number & vbCrLf & Err.Description
End Function

Public Function SqlDateUS(ByVal d As Date) As String
    SqlDateUS = Format$(d, "mm\/dd\/yyyy")
End Function

'========================
' TEMPVARS PROYECTOS
'========================
Public Sub LimpiarTempVarsIndicador()
    On Error Resume Next
    TempVars.Remove TVAR_IDSCSV
    TempVars.Remove TVAR_COUNT
    TempVars.Remove TVAR_OPENED
End Sub


Public Function GetIdsCsvSeleccionados(ByRef p_Error As String) As String
    p_Error = ""
    If IsNull(TempVars(TVAR_IDSCSV)) Then
        p_Error = "No hay proyectos seleccionados."
        GetIdsCsvSeleccionados = ""
        Exit Function
    End If
    GetIdsCsvSeleccionados = Nz(TempVars(TVAR_IDSCSV), "")
    If Len(GetIdsCsvSeleccionados) = 0 Then p_Error = "No hay proyectos seleccionados."
End Function

Public Function GetCountSeleccionados() As Long
    On Error Resume Next
    If Not IsNull(TempVars(TVAR_COUNT)) Then
        GetCountSeleccionados = CLng(Nz(TempVars(TVAR_COUNT), 0))
    Else
        GetCountSeleccionados = 0
    End If
End Function

'========================
' EXCEL EXPORT
'========================
Public Function ExportarSQLaExcel( _
                                    ByVal p_SQL As String, _
                                    ByVal p_Titulo As String, _
                                    ByRef p_Error As String _
                                ) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim xl As Object, wb As Object, ws As Object
    Dim i As Long

    On Error GoTo errores
    p_Error = ""
    ExportarSQLaExcel = False

    Set db = getdb()
    Set rs = db.OpenRecordset(p_SQL, dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        p_Error = "No hay registros para exportar."
        GoTo salir
    End If

    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ws.Name = "Datos"

    ' Cabeceras
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
        ws.Cells(1, i + 1).Font.Bold = True
    Next

    ' Datos
    ws.Range("A2").CopyFromRecordset rs

    ws.Columns.AutoFit
    xl.Visible = True

    ExportarSQLaExcel = True

salir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Exit Function

errores:
    p_Error = "ExportarSQLaExcel: " & Err.Number & vbCrLf & Err.Description
    Resume salir
End Function
Public Function HaPasadoPorDialogo() As Boolean
    On Error Resume Next
    If Not (IsNull(TempVars(TVAR_IDSCSV)) Or IsNull(TempVars(TVAR_COUNT))) Then
        HaPasadoPorDialogo = True
        
    End If
    
   
End Function


