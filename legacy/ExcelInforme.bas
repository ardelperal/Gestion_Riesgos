Attribute VB_Name = "ExcelInforme"
Option Compare Database
Option Explicit
Private m_NumeroHojaPreviaARiesgos As Long
Private m_Alto As Long

Public Function getURLExcel( _
                            p_Edicion As Edicion, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim m_Cod As String
    Dim m_Proyecto As Proyecto
    
    On Error GoTo errores
   
    If p_Edicion Is Nothing Then
       p_Error = "La edición hay que introducirla"
       Err.Raise 1000
    End If
    Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Cod = m_Proyecto.CodigoDocumento
    p_Error = m_Proyecto.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Cod = "" Then
        p_Error = "No se sabe el código del documento del proyecto al que pertenece el informe"
        Err.Raise 1000
    End If
    m_Cod = Replace(m_Cod, "/", "_")
    m_Cod = Replace(m_Cod, "\", "_")
    m_Cod = Replace(m_Cod, vbNewLine, "")
    If m_Proyecto.Juridica = "TdE" Then
            getURLExcel = m_ObjEntorno.URLDirectorioLocal & m_Cod & "V" & _
                    Format(p_Edicion.Edicion, "00") & ".xlsx"
    Else
        getURLExcel = m_ObjEntorno.URLDirectorioLocal & m_Cod & "-" & p_Edicion.Edicion & ".xlsx"
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getURLExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
   
End Function

Public Function GenerarInforme( _
                                p_Edicion As Edicion, _
                                Optional ByRef p_EnExcel As EnumSiNo = EnumSiNo.No, _
                                Optional p_FechaCierre As String, _
                                Optional p_FechaPublicacion As String, _
                                Optional ByRef p_Error As String, _
                                Optional p_db As DAO.Database = Nothing _
                                ) As String
    
    
    
    
    Dim m_ObjProyecto As Proyecto
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    
    Dim wbLibro As Object
    Dim wbHoja As Object
    Dim appExcel As Object
    Dim m_URLPDF As String
    Dim m_URLExcel As String
    Dim m_Col As Scripting.Dictionary
    Dim intHojaRiesgo As Integer
    
    On Error GoTo errores
    
    If p_EnExcel = Empty Then
        p_EnExcel = EnumSiNo.No
    End If
    
    m_URLExcel = getURLExcel(p_Edicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not p_Edicion.EdicionAnterior Is Nothing Then
        If Not IsDate(p_Edicion.FechaPublicacion) Then
            p_Edicion.GrabarCambiosEnEdicion p_Error, p_db
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
    End If
    
    Set m_ObjProyecto = p_Edicion.Proyecto
    p_Edicion.OrdenarColeRiesgosAscendentemente = EnumSiNo.Sí
    If Not FSO.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Tiene abierto un informe anterior"
            Err.Raise 1000
        End If
    End If
    If p_EnExcel = EnumSiNo.No Then
        m_URLPDF = FSO.GetParentFolderName(m_URLExcel) & "\" & FSO.GetBaseName(m_URLExcel) & ".pdf"
        If FSO.FileExists(m_URLPDF) Then
            If FicheroAbierto(m_URLPDF) Then
                p_Error = "Tiene abierto un informe anterior"
                Err.Raise 1000
            End If
        End If
    End If
    FSO.CopyFile m_ObjEntorno.URLPlantillaExcel, m_URLExcel, True
    Set appExcel = CreateObject("Excel.Application")
    appExcel.Visible = False
    appExcel.DisplayAlerts = False
    Set wbLibro = appExcel.Workbooks.Open(m_URLExcel)
    Set wbHoja = wbLibro.Worksheets(1)
    'WbHoja.Name = "Inicio"
    'WbHoja.Move before:=wbLibro.Worksheets(1)
    Avance "Generando Hoja Inicial"
    GeneraHojaInicial p_Edicion, wbHoja, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set wbHoja = wbLibro.Worksheets.Add
    wbHoja.Name = "Portada"
    wbHoja.Move After:=wbLibro.Worksheets("Inicio")
    'WbHoja.Move before:=wbLibro.Worksheets(2)
    Avance "Generando Portada"
    GeneraHojaPortada p_Edicion, wbHoja, p_FechaPublicacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set wbHoja = wbLibro.Worksheets.Add
    wbHoja.Name = "Inventario Riesgos Detectados"
    wbHoja.Move After:=wbLibro.Worksheets("Portada")
    Avance "Generando Inventario"
    If m_ObjProyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        GeneraHojaInventarioConCausaRaiz p_Edicion, wbHoja, p_FechaCierre, p_FechaPublicacion, p_Error
    Else
        GeneraHojaInventarioSinCausaRaiz p_Edicion, wbHoja, p_FechaCierre, p_FechaPublicacion, p_Error
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Col = p_Edicion.ColRiesgosPorPrioridadTodos
    If m_Col Is Nothing Then
        Set m_Col = p_Edicion.colRiesgos
        If m_Col Is Nothing Then
            p_Error = "No se han podido cargar los errores antes de ir generando las hojas de riesgo"
            Err.Raise 1000
        End If
    End If
    intHojaRiesgo = 4
    For Each m_IdRiesgo In m_Col
        Set m_ObjRiesgo = m_Col(m_IdRiesgo)
        If m_ObjRiesgo Is Nothing Then
            GoTo siguiente
        End If
        'If m_ObjRiesgo.CodigoRiesgo = "R015" Then Stop
        Set wbHoja = wbLibro.Worksheets.Add
        wbHoja.Name = m_ObjRiesgo.CodigoRiesgo
        wbHoja.Move After:=wbLibro.Worksheets(intHojaRiesgo)
        Avance "Generando Hoja del Riesgo: " & m_ObjRiesgo.CodigoRiesgo
        GeneraHojaRiesgo m_ObjRiesgo, wbHoja, p_FechaCierre, p_FechaPublicacion, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        intHojaRiesgo = intHojaRiesgo + 1
siguiente:
    Next
    Avance "Dando formato al informe"
    FormatearParaImpresion appExcel, p_Edicion, wbLibro, p_FechaPublicacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If p_EnExcel = EnumSiNo.No Then
        Avance "Exportando a PDF"
        wbLibro.Save
        wbLibro.Close
        On Error Resume Next
        Set wbLibro = appExcel.Workbooks.Open(m_URLExcel)
        If Err.Number <> 0 Then
            Err.Raise 1000
        End If
        wbLibro.ExportAsFixedFormat xlTypePDF, m_URLPDF
        m_URLExcel = m_URLPDF
    Else
        m_URLExcel = m_URLExcel
    End If
    wbLibro.Close True
    Set wbLibro = Nothing
    appExcel.Quit
    Set appExcel = Nothing
    GenerarInforme = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarInforme ha producido el error nº: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    If Not wbLibro Is Nothing Then
        wbLibro.Close False
        Set wbLibro = Nothing
    End If
    If Not appExcel Is Nothing Then
        appExcel.Quit
        Set appExcel = Nothing
    End If
    
End Function
Public Function FormatearParaImpresion( _
                                        appExcel As Object, _
                                        p_ObjEdicion As Edicion, _
                                        wbLibro As Object, _
                                        Optional p_fechaRef As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim wbHoja As Object
    Dim i As Integer
    Dim m_Pie As String
    Dim m_PieIzda As String
    Dim m_PieCentro As String
    Dim m_PieDerecho As String
    Dim intNumeroRiesgos As Integer
    Dim m_Fecha As String
    Dim m_Proyecto As Proyecto
    
    
    
    On Error GoTo errores
    '-------------------
    ' PORTADA
    '-----------------------------
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se ha indicado la Edición"
        Err.Raise 1000
    End If
    If Not IsDate(p_fechaRef) Then p_fechaRef = Date
    m_Fecha = CStr(p_fechaRef)
    Set m_Proyecto = Constructor.getProyecto(p_ObjEdicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_PieDerecho = "Edición: " & p_ObjEdicion.Edicion
    intNumeroRiesgos = p_ObjEdicion.colRiesgos.Count
    
    
    m_PieIzda = "&""-,Negrita""&9" & m_Proyecto.CodigoDocumento & "&"
    m_PieCentro = "&""-,Negrita""&9&P de &N" & "&"
    m_PieDerecho = "Edición: " & p_ObjEdicion.Edicion & "&"
    'm_fecha = Format(m_fecha, "dd/mm/yyyy")
    m_Pie = m_Proyecto.CodigoDocumento & Chr(10) & "Fecha: " & m_Fecha & Chr(10) & m_PieDerecho
    
    'On Error Resume Next
    
    Set wbHoja = wbLibro.Worksheets(1)
        With wbHoja.PageSetup
        wbLibro.Application.PrintCommunication = False
        .LeftFooter = m_Pie
        .BottomMargin = wbLibro.Application.CentimetersToPoints(2.6)
        .FooterMargin = wbLibro.Application.CentimetersToPoints(0.6)
        '.CenterFooter = m_PieCentro
        '.RightFooter = m_PieDerecho
      
       ' .LeftFooter = m_Pie
       ' .CenterFooter = m_PieDerecho
        '.LeftMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.RightMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.TopMargin = wbLibro.Application.InchesToPoints(0.748031496062992)
        
        '.HeaderMargin = wbLibro.Application.InchesToPoints(0.31496062992126)
        
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .BlackAndWhite = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    wbLibro.Application.PrintCommunication = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
   
    Set wbHoja = wbLibro.Worksheets(2)
    With wbHoja.PageSetup
        wbLibro.Application.PrintCommunication = False
        .LeftFooter = m_Pie
        .BottomMargin = wbLibro.Application.CentimetersToPoints(2.6)
        .FooterMargin = wbLibro.Application.CentimetersToPoints(0.6)
        '.CenterFooter = m_PieCentro
        '.RightFooter = m_PieDerecho
        '.LeftFooter = m_Pie
        '.CenterFooter = "2 de " & intNumeroRiesgos + 4
        '.CenterFooter = "Página &P de &N"
        '.RightFooter = m_PieDerecho
        '.LeftMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.RightMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.TopMargin = wbLibro.Application.InchesToPoints(0.94488188976378)
        '.BottomMargin = wbLibro.Application.InchesToPoints(0.94488188976378)
        '.HeaderMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.FooterMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 10
        
        wbLibro.Application.PrintCommunication = True
        If Err.Number <> 0 Then
            Err.Clear
        End If
        
    End With
    Set wbHoja = wbLibro.Worksheets(3)
    With wbHoja.PageSetup
        wbLibro.Application.PrintCommunication = False
        .LeftFooter = m_Pie
        .BottomMargin = wbLibro.Application.CentimetersToPoints(2.6)
        .FooterMargin = wbLibro.Application.CentimetersToPoints(0.6)
        '.CenterFooter = m_PieCentro
        '.RightFooter = m_PieDerecho
'        .LeftFooter = m_Pie
'        .CenterFooter = "Página &P de &N"
'        .RightFooter = m_PieDerecho
        '.LeftMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.RightMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
        '.TopMargin = wbLibro.Application.InchesToPoints(0.748031496062992)
        '.BottomMargin = wbLibro.Application.InchesToPoints(0.748031496062992)
        '.HeaderMargin = wbLibro.Application.InchesToPoints(0.31496062992126)
        '.FooterMargin = wbLibro.Application.InchesToPoints(0.31496062992126)
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .BlackAndWhite = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        
        wbLibro.Application.PrintCommunication = True
        If Err.Number <> 0 Then
            Err.Clear
        End If
        
    End With
    If p_ObjEdicion.colRiesgos.Count > 0 Then
        For i = 4 To intNumeroRiesgos + 3
            Set wbHoja = wbLibro.Worksheets(i)
            With wbHoja.PageSetup
                wbLibro.Application.PrintCommunication = False
                .LeftFooter = m_Pie
                .BottomMargin = wbLibro.Application.CentimetersToPoints(2.6)
                .FooterMargin = wbLibro.Application.CentimetersToPoints(0.6)
                '.CenterFooter = m_PieCentro
                '.RightFooter = m_PieDerecho
                '.LeftMargin = wbLibro.Application.InchesToPoints(0.708661417322835)
                '.RightMargin = wbLibro.Application.InchesToPoints(0.708661417322835)
                '.TopMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
                '.BottomMargin = wbLibro.Application.InchesToPoints(0.393700787401575)
                '.HeaderMargin = wbLibro.Application.InchesToPoints(0.31496062992126)
                '.FooterMargin = wbLibro.Application.InchesToPoints(0.31496062992126)
                .Orientation = xlPortrait
                .PaperSize = xlPaperA4
                .FitToPagesWide = 1
                .FitToPagesTall = 10
            End With
            wbLibro.Application.PrintCommunication = True
            If Err.Number <> 0 Then
                Err.Clear
            End If
            
        Next
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.FormatearParaImpresion ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    appExcel.Echo True
End Function
Private Function GeneraHojaInicial( _
                                    p_ObjEdicion As Edicion, _
                                    ByRef wbHoja As Object, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 19/06/2018
    '   -Modificaciones:
    
    '
    '   -Funcionamiento:
   
    '   -Llamada desde:
    
    '   -Devuelve:
    '       GeneraHojaInicial = Descriptivo
    '       GeneraHojaInicial = "#ERR" & "|" & p_Error
    '-------------------------------------------------------------------
   
    Dim intFilaProyecto As Integer
    Dim intFilaNombreProyecto As Integer
    Dim m_ObjProyecto As Proyecto
    
    On Error GoTo errores
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
        
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se conoce la Edición del Riesgo"
        Err.Raise 1000
    End If
    Set m_ObjProyecto = p_ObjEdicion.Proyecto
    p_Error = p_ObjEdicion.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se conoce el proyecto del Riesgo"
        Err.Raise 1000
    End If
    If m_ObjProyecto.NombreProyecto = "" Then
        p_Error = "Falta el Código del Expediente o el de la Actividad"
        Err.Raise 1000
    End If
    If m_ObjProyecto.Proyecto = "" Then
        p_Error = "Falta el Código del Expediente o el de la Actividad"
        Err.Raise 1000
    End If
    intFilaNombreProyecto = 20
    intFilaProyecto = 21
   
    With wbHoja
        
        .Cells(intFilaNombreProyecto - 3, 1).Value = "INFORME DE gestión de riesgos"
        .Cells(intFilaNombreProyecto, 1).Value = UCase(m_ObjProyecto.NombreProyecto)
        .Cells(intFilaProyecto, 1).Value = UCase(m_ObjProyecto.Proyecto)
        
    End With
    With wbHoja
        .Rows("" & CStr(intFilaNombreProyecto - 3) & ":" & CStr(intFilaNombreProyecto - 3) & "").RowHeight = 27.75
        .Rows("" & CStr(intFilaNombreProyecto) & ":" & CStr(intFilaNombreProyecto) & "").RowHeight = 27.75
        .Rows("" & CStr(intFilaProyecto) & ":" & CStr(intFilaProyecto) & "").RowHeight = 27.75
        .Range(.Cells(intFilaNombreProyecto - 3, 1), .Cells(intFilaNombreProyecto - 3, 16)).MergeCells = True
        .Range(.Cells(intFilaNombreProyecto, 1), .Cells(intFilaNombreProyecto, 16)).MergeCells = True
        .Range(.Cells(intFilaNombreProyecto, 1), .Cells(intFilaNombreProyecto, 16)).ShrinkToFit = True
        .Range(.Cells(intFilaProyecto, 1), .Cells(intFilaProyecto, 16)).MergeCells = True
        .Range(.Cells(intFilaProyecto, 1), .Cells(intFilaProyecto, 16)).ShrinkToFit = True
        
        .Range(.Cells(intFilaNombreProyecto - 3, 1), .Cells(intFilaNombreProyecto - 3, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFilaNombreProyecto, 1), .Cells(intFilaNombreProyecto, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFilaProyecto, 1), .Cells(intFilaProyecto, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(intFilaNombreProyecto - 3, 1), .Cells(intFilaProyecto, 1)).Font.Name = "TheSansCorrespondence"
        .Range(.Cells(intFilaNombreProyecto - 3, 1), .Cells(intFilaProyecto, 1)).Font.Size = 22
        .Range(.Cells(intFilaNombreProyecto - 3, 1), .Cells(intFilaProyecto, 1)).Font.Bold = True
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.GeneraHojaInicial ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GeneraHojaPortada( _
                                    p_ObjEdicion As Edicion, _
                                    ByRef wbHoja As Object, _
                                    Optional p_FechaPublicacion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    
    Dim MiRango As Object
    Dim intFilaInicioEdicionCambio As Integer
    Dim intFilaFinalEdicionCambio As Integer
    Dim intFilaEdicionActual As Integer
    Dim intFila As Integer
    Dim i As Integer
    Dim intFilaTitulares As Integer
    Dim intFilaFinTabla As Integer
    Dim intFilaInicioCambios As Integer
    Dim intFilaFinCambios As Integer
    Dim intOrdinalEdicion As Integer
    Dim m_objEdicion As Edicion
    Dim m_IDEdicion As Variant
    Dim m_ObjProyecto As Proyecto
    Dim m_objColEdiciones As Scripting.Dictionary
    Dim m_EsParaPublicar As EnumSiNo
    Dim m_Elaborado As String
    Dim m_Revisado As String
    Dim m_Aprobado As String
    
    Dim dato
    On Error GoTo errores
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se conoce la edición del informe"
        Err.Raise 1000
    End If
    
    Avance "Portada: Obteniendo datos generales del Proyecto"
    Set m_ObjProyecto = p_ObjEdicion.Proyecto
    
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se sabe para qué proyecto es la portada"
        Err.Raise 1000
    End If
    
    Avance "Portada: Obteniendo colección de Ediciones"
    Set m_objColEdiciones = getEdicionesInvolucradasParaInforme(p_ObjEdicion.IDEdicion, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If m_objColEdiciones Is Nothing Then
        p_Error = "No se conocen las ediciones del informe"
        Err.Raise 1000
    End If
        
    
    '---------------------------------
    ' ANCHO DE COLUMNAS
    '-------------------------------
        With wbHoja
            .Columns("A:A").ColumnWidth = 1.71
            .Columns("B:B").ColumnWidth = 10.71
            .Columns("C:C").ColumnWidth = 13.29
            .Columns("D:D").ColumnWidth = 30.29
            .Columns("E:E").ColumnWidth = 30.29
            .Columns("F:F").ColumnWidth = 30.29
        End With
        With wbHoja
            .Rows("16:16").RowHeight = 15
            .Rows("18:18").RowHeight = 15
            .Range("A16:I16").MergeCells = False
            .Range("A18:I18").MergeCells = False
        End With
    '-------------------------------
    ' INSERTAR IMAGEN
    '-------------------------------
        'wbHoja.Pictures.Insert (strURLLogoTelefonica)
    '---------------------------------
    ' PORTADA
    '-------------------------------
        intFila = 1
        With wbHoja
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 35
            .Range(wbHoja.Cells(intFila, 2), .Cells(intFila, 6)).MergeCells = True
        End With
        intFila = intFila + 1
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "PORTADA"
                .Font.Bold = True
                .Font.Size = 16
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, p_Error:=p_Error
            
        End With
        intFila = intFila + 2
    '---------------------------------
    ' Proyecto
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Value = "Proyecto"
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434828
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila, 4))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "Expediente Nº: " & m_ObjProyecto.Proyecto
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
        End With
        intFila = intFila + 1
    '---------------------------------
    ' Nombre del proyecto
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Value = "Nombre del proyecto"
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434828
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila, 4))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = m_ObjProyecto.NombreProyecto
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
        End With
        intFila = intFila + 1
    '---------------------------------
    ' Jefe del Proyecto
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Value = "Jefe del Proyecto"
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434828
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila, 4))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = p_ObjEdicion.Elaborado
                
                
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
        End With
     intFila = intFila + 1
    '---------------------------------
    ' cliente
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Value = "Cliente"
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434828
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila, 4))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = m_ObjProyecto.Cliente
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            Set MiRango = .Range(.Cells(intFila - 3, 2), .Cells(intFila, 6))
            
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            
        End With
        intFila = intFila + 1
    '---------------------------------
    '  CUADRO DE CONTROL
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = " CUADRO DE CONTROL"
                .Font.Bold = True
                .Font.Size = 16
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16764057
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            
        End With
        intFila = intFila + 2
        intFilaTitulares = intFila
        
    
    '---------------------------------
    ' Edición FECHA   ELABORADO   REVISADO    APROBADO
    '-------------------------------
        With wbHoja
            .Cells(intFila, 2) = "EDICIÓN"
            .Cells(intFila, 3) = "FECHA"
            .Cells(intFila, 4) = "ELABORADO"
            .Cells(intFila, 5) = "REVISADO"
            .Cells(intFila, 6) = "APROBADO"
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 6))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            For Each m_IDEdicion In m_objColEdiciones
                If Nz(m_IDEdicion, "") = "" Then
                    GoTo siguienteEdicion
                End If
                Set m_objEdicion = m_objColEdiciones(m_IDEdicion)
                If Not IsDate(p_FechaPublicacion) Then
                    If InStr(1, m_ObjProyecto.Elaborado, vbNewLine) <> 0 Then
                        dato = Split(m_ObjProyecto.Elaborado, vbNewLine)
                        m_Elaborado = dato(0)
                    Else
                        m_Elaborado = m_ObjProyecto.Elaborado
                    End If
                    If InStr(1, m_ObjProyecto.Revisado, vbNewLine) <> 0 Then
                        dato = Split(m_ObjProyecto.Revisado, vbNewLine)
                        m_Revisado = dato(0)
                    Else
                        m_Revisado = m_ObjProyecto.Revisado
                    End If
                    If InStr(1, m_ObjProyecto.Aprobado, vbNewLine) <> 0 Then
                        dato = Split(m_ObjProyecto.Aprobado, vbNewLine)
                        m_Aprobado = dato(0)
                    Else
                        m_Aprobado = m_ObjProyecto.Aprobado
                    End If
                Else
                    If InStr(1, m_objEdicion.Elaborado, vbNewLine) <> 0 Then
                        dato = Split(m_objEdicion.Elaborado, vbNewLine)
                        m_Elaborado = dato(0)
                    Else
                        m_Elaborado = m_objEdicion.Elaborado
                    End If
                    If InStr(1, m_objEdicion.Revisado, vbNewLine) <> 0 Then
                        dato = Split(m_objEdicion.Revisado, vbNewLine)
                        m_Revisado = dato(0)
                    Else
                        m_Revisado = m_objEdicion.Revisado
                    End If
                    If InStr(1, m_objEdicion.Aprobado, vbNewLine) <> 0 Then
                        dato = Split(m_objEdicion.Aprobado, vbNewLine)
                        m_Aprobado = dato(0)
                    Else
                        m_Aprobado = m_objEdicion.Aprobado
                    End If
                End If
                
                Avance "Portada: Edición: " & m_objEdicion.Edicion
                intFila = intFila + 1
                
                
'                If Not IsDate(m_FechaPublicacion) Then
'                    m_FechaPublicacion = Date
'                End If
                'm_FechaPublicacion = m_ObjProyecto.Publicacion.FechaPublicacion
                
                .Cells(intFila, 2).Value = m_objEdicion.Edicion
                If m_objEdicion.Edicion = p_ObjEdicion.Edicion Then
                    
                    If Not IsDate(p_FechaPublicacion) Then
                        .Range(.Cells(intFila, 3), .Cells(intFila, 6)).MergeCells = True
                        .Cells(intFila, 3).Value = "AÚN SIN PUBLICAR"
                    Else
                        .Cells(intFila, 3).Value = Format(p_FechaPublicacion, "mm/dd/yyyy")
                        .Cells(intFila, 4).Value = m_Elaborado
                        .Cells(intFila, 5).Value = m_Revisado
                        .Cells(intFila, 6).Value = m_Aprobado
                    End If
                    Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 6))
                    With MiRango
                        If Not IsDate(p_FechaPublicacion) Then
                            .Font.Color = RGB(255, 0, 0)
                        End If
                        .Font.Bold = True
                        With MiRango.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 15263976
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End With
                    Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 6))
                    MiRango.HorizontalAlignment = xlLeft
                Else
                    If IsDate(m_objEdicion.FechaEdicion) Then
                        .Cells(intFila, 3).Value = Format(m_objEdicion.FechaEdicion, "mm/dd/yyyy")
                    End If
                    
                    .Cells(intFila, 4).Value = m_objEdicion.Elaborado
                    .Cells(intFila, 5).Value = m_objEdicion.Revisado
                    .Cells(intFila, 6).Value = m_objEdicion.Aprobado
                End If
siguienteEdicion:
            Next
FueraEdicion:
            intFila = intFila + 1
            .Rows(CStr(intFila) & ":" & CStr(intFila)).RowHeight = 100
            'aquí va la firma
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 3))
            With MiRango
                .WrapText = True
            End With
            
            
            Set MiRango = .Range(.Cells(intFilaTitulares + 1, 2), .Cells(intFilaTitulares + m_objColEdiciones.Count + 1, 6))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            Set MiRango = .Range(.Cells(intFilaTitulares, 2), .Cells(intFilaTitulares + m_objColEdiciones.Count + 1, 6))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            
        End With
        intFila = intFilaTitulares + m_objColEdiciones.Count + 2
    '---------------------------------
    ' CONTROL DE CAMBIOS
    '-------------------------------
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Value = "CONTROL DE CAMBIOS"
                .Font.Bold = True
                .Font.Size = 16
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16764057
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 6))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            
        End With
        intFila = intFila + 2
        intFilaTitulares = intFila
    '---------------------------------
    ' Edición APARTADOS   DESCRIPCIÓN del CAMBIO
    '-------------------------------
        With wbHoja
            .Cells(intFila, 2).Value = "EDICIÓN"
            .Cells(intFila, 3).Value = "APARTADOS"
            .Cells(intFila, 4).Value = "DESCRIPCIÓN DEL CAMBIO"
            .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 6))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            intFila = intFila + 1
            Avance "Portada: Escribiendo los cambios----"
            intFilaFinCambios = ParteCambios(m_ObjProyecto, m_objColEdiciones, wbHoja, intFila, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            Set MiRango = .Range(.Cells(intFilaTitulares, 2), .Cells(intFilaFinCambios - 1, 6))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            
            
        End With
    GeneraHojaPortada = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Inform_ObjProyecto.GeneraHojaPortada ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function ParteCambios( _
                                p_Proyecto As Proyecto, _
                                p_ColEdiciones As Scripting.Dictionary, _
                                ByRef wbHoja As Object, _
                                intFilaInicial As Integer, _
                                Optional ByRef p_Error As String _
                                ) As Integer
                                
    '----------------------------------------
    ' RESUMEN
    '----------------------------------------
    
    Dim intFila As Integer
    Dim m_ObjColCambios As Scripting.Dictionary
    Dim m_Cambio As Cambio
    Dim m_IDCambio As Variant
    Dim m_IDEdicion As Variant
    Dim m_Edicion As Edicion
    Dim m_Descripcion As String
    
    On Error GoTo errores
    
    intFila = intFilaInicial
    If p_ColEdiciones Is Nothing Then
        Exit Function
    End If
    For Each m_IDEdicion In p_ColEdiciones
        Set m_Edicion = p_ColEdiciones(m_IDEdicion)
        Avance "Portada: Escribiendo los cambios----Edición " & m_Edicion.Edicion
        Set m_ObjColCambios = m_Edicion.colCambiosConEdicionAnterior
        p_Error = m_Edicion.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ObjColCambios Is Nothing Then
            With wbHoja
                .Range(.Cells(intFila, 3), .Cells(intFila, 6)).MergeCells = True
                .Cells(intFila, 2).Value = m_Edicion.Edicion
                .Cells(intFila, 3).Value = "Sin Cambios"
                .Range(.Cells(intFila, 2), .Cells(intFila, 2)).VerticalAlignment = xlCenter
                .Range(.Cells(intFila, 2), .Cells(intFila, 2)).NumberFormat = "0"
                .Range(.Cells(intFila, 2), .Cells(intFila, 3)).VerticalAlignment = xlCenter
                .Range(.Cells(intFila, 2), .Cells(intFila, 3)).HorizontalAlignment = xlCenter
                .Range(.Cells(intFila, 4), .Cells(intFila, 6)).VerticalAlignment = xlCenter
                .Range(.Cells(intFila, 3), .Cells(intFila, 6)).HorizontalAlignment = xlLeft
                .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Bold = False
                .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Size = 10
                .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Name = "Garamond"
            End With
             intFila = intFila + 1
            GoTo siguiente
        End If
        If Not m_ObjColCambios Is Nothing Then
            For Each m_IDCambio In m_ObjColCambios
                Set m_Cambio = m_ObjColCambios(m_IDCambio)
                m_Descripcion = ""
                If m_Cambio.EdicionInicial = "" Then
                    With wbHoja
                        .Range(.Cells(intFila, 3), .Cells(intFila, 6)).MergeCells = True
                        .Cells(intFila, 2).Value = m_Cambio.EdicionFinal
                        .Cells(intFila, 3).Value = m_Cambio.riesgo
                        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).VerticalAlignment = xlCenter
                        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).NumberFormat = "0"
                        .Range(.Cells(intFila, 2), .Cells(intFila, 3)).VerticalAlignment = xlCenter
                        .Range(.Cells(intFila, 2), .Cells(intFila, 3)).HorizontalAlignment = xlCenter
                        .Range(.Cells(intFila, 4), .Cells(intFila, 6)).VerticalAlignment = xlCenter
                        .Range(.Cells(intFila, 3), .Cells(intFila, 6)).HorizontalAlignment = xlLeft
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Bold = False
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Size = 10
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Name = "Garamond"
                    End With
                    
                Else
                    With wbHoja
                        
                        m_Descripcion = Replace(m_Cambio.Descripcion, vbNewLine, "")
                        .Cells(intFila, 2).Value = m_Cambio.EdicionFinal
                        .Cells(intFila, 3).Value = m_Cambio.riesgo
                        .Cells(intFila, 4).Value = m_Descripcion
                        .Range(.Cells(intFila, 2), .Cells(intFila, 2)).NumberFormat = "0"
                        .Range(.Cells(intFila, 4), .Cells(intFila, 4)).WrapText = True
                        .Range(.Cells(intFila, 4), .Cells(intFila, 6)).MergeCells = True
                        .Range(.Cells(intFila, 2), .Cells(intFila, 3)).VerticalAlignment = xlCenter
                        .Range(.Cells(intFila, 2), .Cells(intFila, 3)).HorizontalAlignment = xlCenter
                        .Range(.Cells(intFila, 4), .Cells(intFila, 6)).VerticalAlignment = xlTop
                        .Range(.Cells(intFila, 4), .Cells(intFila, 6)).HorizontalAlignment = xlLeft
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Bold = False
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Size = 10
                        .Range(.Cells(intFila, 2), .Cells(intFila, 6)).Font.Name = "Garamond"
                    End With
                End If
                
                If m_Descripcion <> "" Then
                    

                    m_Alto = AltoFila(EnumTipoCeldaAlto.PortadaCambios, m_Descripcion)
                    If m_Alto = 0 Then
                        p_Error = "El método AltoFila no ha devuelto un número válido "
                        Err.Raise 1000
                    End If
                    With wbHoja
                        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
                    End With
                End If
                Set m_Cambio = Nothing
                intFila = intFila + 1
            Next
        End If
        
        
siguiente:
        Set m_Edicion = Nothing
        Set m_ObjColCambios = Nothing
    Next
    
    ParteCambios = intFila
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ParteCambios ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
                                    
End Function
Public Function GeneraHojaRiesgo( _
                                    p_ObjRiesgo As riesgo, _
                                    ByRef wbHoja As Object, _
                                    Optional p_FechaCierre As String, _
                                    Optional p_fechaRef As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                        
    Dim intFila As Integer
    Dim intFilaFinal As Integer
    Dim intFilaInicialTabla As Integer
    Dim intFilaFinalTabla As Integer
    Dim m_EstadoRiesgo As Variant
    Dim m_FechaEstado As String
    Dim MiRango As Object
    Dim m_ObjAccionesContingencia As PCAccion
    Dim m_IdAccion As Variant
    Dim m_ObjPMAccion As PMAccion
    Dim m_objEdicion As Edicion
    Dim m_FechaPublicacion As String
    Dim NumeroEstados As Integer
    Dim NumeroEstado As Integer
    Dim m_EstadoParaCelda As String
    Dim m_ObjRiesgoOtrasEdiciones As riesgo
    Dim m_IdRiesgo As Variant
    Dim m_CodigoUnico As String
    Dim m_ObjColRiesgosEstados As Scripting.Dictionary
    Dim m_ConRiesgosDeBiblioteca As EnumSiNo
    On Error GoTo errores
    
    
    intFila = 1
    GeneraHojaRiesgoFicha p_ObjRiesgo, intFila, wbHoja, intFilaFinal, p_fechaRef, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intFila = intFilaFinal + 1
    GeneraHojaRiesgoEstados p_ObjRiesgo, intFila, wbHoja, intFilaFinal, p_FechaCierre, p_fechaRef, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intFila = intFilaFinal + 1
    GeneraHojaRiesgoDatos p_ObjRiesgo, intFila, wbHoja, intFilaFinal, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intFila = intFilaFinal + 1
    GeneraPlanesDeAccion p_ObjRiesgo, intFila, wbHoja, intFilaFinal, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelExcelInforme.GeneraHojaRiesgo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GeneraHojaRiesgoFicha( _
                                        p_ObjRiesgo As riesgo, _
                                        p_FilaInicial As Integer, _
                                        ByRef wbHoja As Object, _
                                        ByRef p_FilaFinal As Integer, _
                                        Optional p_fechaRef As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim m_objEdicion As Edicion
    Dim MiRango As Object
    Dim intFila As Integer
    Dim intFilaInicialTabla As Integer
    Dim intFilaFinalTabla As Integer
    Dim m_FechaPublicacion As String
   
    On Error GoTo errores
    
    
    intFila = p_FilaInicial
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar El Riesgo"
        Err.Raise 1000
    End If
    Set m_objEdicion = p_ObjRiesgo.Edicion
    If m_objEdicion Is Nothing Then
        p_Error = "No se conoce la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Not IsDate(p_fechaRef) Then
        p_fechaRef = Date
    End If
    
    m_FechaPublicacion = CStr(p_fechaRef)

   
    '---------------------------------
    ' ANCHO DE COLUMNAS
    '-------------------------------
        With wbHoja
            .Columns("A:A").ColumnWidth = 2.57
            .Columns("B:B").ColumnWidth = 13.86
            .Columns("C:C").ColumnWidth = 13.14
            .Columns("D:D").ColumnWidth = 12.57
            .Columns("E:E").ColumnWidth = 11.57
            .Columns("F:F").ColumnWidth = 10.71
            .Columns("G:G").ColumnWidth = 9.71
            .Columns("H:H").ColumnWidth = 11.57
            .Columns("I:I").ColumnWidth = 11.43
        End With
        With wbHoja
            .Rows("16:16").RowHeight = 15
            .Rows("18:18").RowHeight = 15
            .Range("A16:I16").MergeCells = False
            .Range("A18:I18").MergeCells = False
        End With
    '---------------------------------
    ' FICHA DE RIESGO
    '-------------------------------
    
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoDatos)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoDatos"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(wbHoja.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
    End With
    intFila = intFila + 1
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraPpal)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraPpal"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Value = "FICHA DE RIESGO"
            .Font.Bold = True
            .Font.Size = 16
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
    End With
    '---------------------------------
    ' Proyecto:
    ' Fecha:
    '-------------------------------
    intFila = intFila + 1 '--->3
    With wbHoja
        .Range(.Cells(intFila, 3), .Cells(intFila, 9)).MergeCells = True
        .Cells(intFila, 2).Value = "Proyecto:"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Value = "Proyecto:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 3))
        With MiRango
            .Value = m_objEdicion.Proyecto.NombreProyecto
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
    End With
    intFila = intFila + 1 '--->4
    With wbHoja
        .Range(.Cells(intFila, 3), .Cells(intFila, 9)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Value = "Fecha:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 3))
        With MiRango
            If IsDate(m_FechaPublicacion) Then
                .Value = Format(m_FechaPublicacion, "mm/dd/yyyy")
            Else
                .Value = "NO PUBLICADO"
                .Font.Color = RGB(255, 0, 0)
            End If
            .Font.Name = "Garamond"
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
        End With
        Set MiRango = .Range(.Cells(intFila - 1, 2), .Cells(intFila, 9))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
    End With
    
    p_FilaFinal = intFila
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.GeneraHojaRiesgoFicha ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function



Private Function GeneraHojaRiesgoEstados( _
                                        p_ObjRiesgo As riesgo, _
                                        p_FilaInicial As Integer, _
                                        ByRef wbHoja As Object, _
                                        ByRef p_FilaFinal As Integer, _
                                        Optional p_FechaCierre As String, _
                                        Optional p_fechaRef As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                        
    Dim intFila As Integer
    Dim intFilaInicialTabla As Integer
    Dim intFilaFinalTabla As Integer
    Dim m_Id As Variant
    Dim m_Resultado As String
    Dim dato As Variant
    Dim m_EstadoRiesgo As String
    Dim m_FechaEstado As String
    Dim MiRango As Object
    Dim m_FechaPublicacion As String
    
    Dim m_ObjColRiesgosEstados As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    
    intFila = p_FilaInicial
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar El Riesgo"
        Err.Raise 1000
    End If
    If Not IsDate(p_fechaRef) Then
        p_fechaRef = Date
    End If
    m_FechaPublicacion = CStr(p_fechaRef)
    
    
    '---------------------------------
    '   Estado del Riesgo
    '-------------------------------
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Value = "Estado del Riesgo"
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764057
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
    End With
    
    
    
    
    '---------------------------------
    ' Estado:
    ' Fecha estado:
    '-------------------------------
    intFila = intFila + 1 '--->5
    intFilaInicialTabla = intFila
    With wbHoja
        .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
        .Cells(intFila, 2).Value = "Estado"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            .HorizontalAlignment = xlCenter
        End With
        .Range(.Cells(intFila, 4), .Cells(intFila, 5)).MergeCells = True
        .Cells(intFila, 4).Value = "Fecha estado"
        Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila, 4))
        With MiRango
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            .HorizontalAlignment = xlCenter
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 5))
        With MiRango.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 16777164
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    With wbHoja
        Set m_ObjColRiesgosEstados = getEstadosDiferentesHastaEdicion(p_ObjRiesgo.Edicion, p_ObjRiesgo.CodigoRiesgo, _
                                m_FechaPublicacion, p_FechaCierre, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        If Not m_ObjColRiesgosEstados Is Nothing Then
            For Each m_Id In m_ObjColRiesgosEstados
                m_Resultado = m_ObjColRiesgosEstados(m_Id)
                dato = Split(m_Resultado, "|")
                m_EstadoRiesgo = dato(0)
                m_FechaEstado = dato(1)
                intFila = intFila + 1
                .Range(.Cells(intFila, 2), .Cells(intFila, 3)).MergeCells = True
                .Cells(intFila, 2).Value = m_EstadoRiesgo
                .Range(.Cells(intFila, 4), .Cells(intFila, 5)).MergeCells = True
                If IsDate(m_FechaEstado) Then
                    .Cells(intFila, 4).Value = Format(m_FechaEstado, "mm/dd/yyyy")
                End If
            Next
        End If
    End With
        
       
    intFilaFinalTabla = intFila
    With wbHoja
        Set MiRango = .Range(.Cells(intFilaInicialTabla, 2), .Cells(intFilaFinalTabla, 5))
        With MiRango
            .Font.Bold = False
            .Font.Size = 10
            .Font.Name = "Garamond"
            .HorizontalAlignment = xlCenter
        End With
        
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
        EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        .Range(.Cells(intFilaInicialTabla, 6), .Cells(intFilaFinalTabla, 9)).MergeCells = True
        Set MiRango = .Range(.Cells(intFilaInicialTabla, 6), .Cells(intFilaFinalTabla, 9))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
        EnumAnchoLinea.Gruesa, p_Error:=p_Error
    End With
    p_FilaFinal = intFila
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.GeneraHojaRiesgoEstados ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function GeneraHojaRiesgoDatos( _
                                        p_ObjRiesgo As riesgo, _
                                        p_FilaInicial As Integer, _
                                        ByRef wbHoja As Object, _
                                        ByRef p_FilaFinal As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                        
    Dim m_ObjProyecto As Proyecto
    Dim m_objEdicion As Edicion
    Dim MiRango As Object
    Dim intFila As Integer
    Dim m_Descripcion As String
    Dim m_CausaRaiz As String
    On Error GoTo errores
    
    
    intFila = p_FilaInicial
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
    If p_ObjRiesgo Is Nothing Then
        p_Error = "Se ha de indicar El Riesgo"
        Err.Raise 1000
    End If
    Set m_objEdicion = p_ObjRiesgo.Edicion
    If m_objEdicion Is Nothing Then
        p_Error = "No se conoce la Edición del Riesgo"
        Err.Raise 1000
    End If
    
    Set m_ObjProyecto = m_objEdicion.Proyecto
    '---------------------------------
    ' Datos del Riesgo
    '-------------------------------
    
    
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Value = "Datos del Riesgo"
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764057
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
    End With
    
'-------------------------------------------------------------------------------------------------------------------
' Tabla  Código riesgo | Detectado por | Impacto Global | Vulnerabilidad | Valoración | Mitigación | Contingencia
'-------------------------------------------------------------------------------------------------------------------
    intFila = intFila + 1
    With wbHoja
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        With MiRango.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 16777164
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        .Cells(intFila, 2).Value = "Código Riesgo"
        .Cells(intFila, 3).Value = "Detectado por"
        .Cells(intFila, 4).Value = "Impacto Global"
        .Cells(intFila, 5).Value = "Vulnerabilidad"
        .Cells(intFila, 6).Value = "Valoración"
        .Cells(intFila, 7).Value = "Mitigación"
        .Cells(intFila, 8).Value = "Contingencia"
        .Cells(intFila, 9).Value = "Materializado"
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
    End With
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 2).Value = p_ObjRiesgo.CodigoRiesgo
        .Cells(intFila, 3).Value = p_ObjRiesgo.DetectadoPor
        .Cells(intFila, 4).Value = p_ObjRiesgo.ImpactoGlobal
        .Cells(intFila, 5).Value = p_ObjRiesgo.Vulnerabilidad
        .Cells(intFila, 6).Value = p_ObjRiesgo.Valoracion
        .Cells(intFila, 7).Value = p_ObjRiesgo.Mitigacion
        .Cells(intFila, 8).Value = p_ObjRiesgo.ContingenciaCalculada
        If IsDate(p_ObjRiesgo.FechaMaterializado) Then
            .Cells(intFila, 9).Value = Format(p_ObjRiesgo.FechaMaterializado, "mm/dd/yyyy")
        End If
        Set MiRango = .Range(.Cells(intFila - 1, 2), .Cells(intFila, 9))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Font.Bold = False
            .Font.Size = 10
            .Font.Name = "Garamond"
            .WrapText = True
        End With
        If IsDate(p_ObjRiesgo.FechaMaterializado) Then
            Set MiRango = .Range(.Cells(intFila - 1, 9), .Cells(intFila, 9))
            With MiRango
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Size = 10
            .Font.Name = "Garamond"
            .WrapText = True
        End With
    End With
    Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
    'origen del riesgo
    intFila = intFila + 1
    With wbHoja
        .Cells(intFila, 2).Value = "Origen"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 3), .Cells(intFila, 9)).MergeCells = True
        Dim m_OrigenCompleto As String
        m_OrigenCompleto = p_ObjRiesgo.Origen & ": " & m_ObjEntorno.ColOrigenRiesgosValores(p_ObjRiesgo.Origen)
        .Cells(intFila, 3).Value = m_OrigenCompleto
        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 3))
        With MiRango
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Size = 10
            .Font.Name = "Garamond"
            
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
    End With
    
    
    '---------------------------------
    ' Descripción del Riesgo
    '-------------------------------
    intFila = intFila + 1
    With wbHoja
         m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
        .Cells(intFila, 2).Value = "Descripción del Riesgo"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
    End With
    intFila = intFila + 1
    With wbHoja
        m_Descripcion = Replace(p_ObjRiesgo.Descripcion, vbNewLine, "")
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoDescripcion, m_Descripcion)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoDescripcion"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        With MiRango
            .MergeCells = True
            .WrapText = True
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlTop
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Value = m_Descripcion
            .Font.Size = 10
            .Font.Name = "Garamond"
            .WrapText = True
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        If m_ObjProyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
                EnumAnchoLinea.Gruesa, p_Error:=p_Error
        Else
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
                EnumAnchoLinea.Gruesa, p_Error:=p_Error
        End If
        
    End With
    If m_ObjProyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
        '---------------------------------
        ' Causa Raíz
        '-------------------------------
        intFila = intFila + 1
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 9)).MergeCells = True
            .Cells(intFila, 2).Value = "Causa Raíz"
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            With MiRango
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
                EnumAnchoLinea.Gruesa, p_Error:=p_Error
            
        End With
        intFila = intFila + 1
        With wbHoja
            m_CausaRaiz = Replace(p_ObjRiesgo.CausaRaiz, vbNewLine, "")
            m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCausaRaiz, m_CausaRaiz)
            If m_Alto = 0 Then
                p_Error = "El método AltoFila no ha devuelto un número válido para CausaRaiz"
                Err.Raise 1000
            End If
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            With MiRango
                .MergeCells = True
                .WrapText = True
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlTop
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .Value = m_CausaRaiz
                .Font.Size = 10
                .Font.Name = "Garamond"
                .WrapText = True
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        End With
    End If
    

    p_FilaFinal = intFila
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.GeneraHojaRiesgoDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GeneraPlanesDeAccion( _
                                        ByRef p_ObjRiesgo As riesgo, _
                                        ByRef p_FilaInicial As Integer, _
                                        ByRef wbHoja As Object, _
                                        ByRef p_FilaFinal As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                        
    Dim intFila As Integer
    Dim intFinalPC As Integer
    Dim intFinalPMAccion As Integer
    Dim intFinalPCAccion As Integer
    Dim intFilaFinal As Integer
    On Error GoTo errores
    
    intFila = p_FilaInicial
    '---------------------------------
    '   Planes de acción
    '-------------------------------
    
    
    PlanesDeAccionCabecera wbHoja, intFila, intFilaFinal, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    intFila = intFilaFinal
    
    If Not p_ObjRiesgo.ColPMs Is Nothing Then
        Planes EnumSiNo.Sí, p_ObjRiesgo, wbHoja, intFila, intFilaFinal, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        intFila = intFilaFinal
    End If
    If Not p_ObjRiesgo.ColPMs Is Nothing Then
        Planes EnumSiNo.No, p_ObjRiesgo, wbHoja, intFila, intFilaFinal, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        intFila = intFilaFinal
    End If
    
    
    
    p_FilaInicial = intFila
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraPlanesDeAccion ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Function PlanesDeAccionCabecera( _
                                        ByRef wbHoja As Object, _
                                        p_FilaInicial As Integer, _
                                        ByRef p_FilaFinal As Integer, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim intFila As Integer
    Dim MiRango As Object
    
    On Error GoTo errores
    intFila = p_FilaInicial
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            .Value = "Planes de acción"
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = 16764057
            .Interior.TintAndShade = 0
            .Interior.PatternTintAndShade = 0
        End With
        
        
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        
    End With
    
    p_FilaFinal = intFila
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.PlanesDeAccionCabecera ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function Planes( _
                        p_EsMitigacion As EnumSiNo, _
                        p_ObjRiesgo As riesgo, _
                        ByRef wbHoja As Object, _
                        p_FilaInicial As Integer, _
                        ByRef p_FilaFinal As Integer, _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim intFila As Integer
    Dim MiRango As Object
    Dim i As Integer
    Dim intFinalPlan As Integer
    Dim intInicialPlanAcciones As Integer
    Dim intFinalPlanAcciones As Integer
    Dim m_Id As Variant
    Dim m_Disparador As String
    Dim m_Plan As Object
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    If p_EsMitigacion = EnumSiNo.Sí Then
        Set m_Col = p_ObjRiesgo.ColPMs
    ElseIf p_EsMitigacion = EnumSiNo.No Then
        Set m_Col = p_ObjRiesgo.ColPCs
    Else
        p_Error = "Se ha de indicar si es o no de mitigación el plan"
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
   ' If p_ObjRiesgo.CodigoRiesgo = "R006" Then Stop
    intFila = p_FilaInicial
    
    With wbHoja
        i = 1
        For Each m_Id In m_Col
            intFila = intFila + 1
            Set m_Plan = m_Col(m_Id)
            If m_Plan Is Nothing Then
                GoTo siguiente
            End If
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            With MiRango
                
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 7))
            MiRango.MergeCells = True
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 30
            If p_EsMitigacion = EnumSiNo.Sí Then
                .Cells(intFila, 2).Value = "Plan de Mitigación (" & i & ")"
            Else
                .Cells(intFila, 2).Value = "plan de contingencia (" & i & ")"
            End If
            .Cells(intFila, 8).Value = "Activación"
            .Cells(intFila, 9).Value = "Desactivación"
           
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
            
            intFila = intFila + 1
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 7))
            MiRango.MergeCells = True
            'On Error Resume Next
            m_Disparador = Replace(m_Plan.DisparadorDelPlan, vbNewLine, "")
            m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoPlanDisparador, m_Disparador)
            If m_Alto = 0 Then
                p_Error = "El método AltoFila no ha devuelto un número válido para cabecera de la ficha de riesgo"
                Err.Raise 1000
            End If
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
            .Cells(intFila, 2).Value = m_Disparador
            If IsDate(m_Plan.FechaActivacionCalculada) Then
                .Cells(intFila, 8).Value = Format(m_Plan.FechaActivacionCalculada, "mm/dd/yyyy")
            End If
            If IsDate(m_Plan.FechaDesActivacionCalculada) Then
                .Cells(intFila, 9).Value = Format(m_Plan.FechaDesActivacionCalculada, "mm/dd/yyyy")
            End If
            
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .VerticalAlignment = xlTop
                .WrapText = True
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            Set MiRango = .Range(.Cells(intFila, 8), .Cells(intFila, 9))
            With MiRango
                .VerticalAlignment = xlTop
                .HorizontalAlignment = xlCenter
                .Font.Bold = False
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
            intFila = intFila + 1
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 4))
            With MiRango
                .MergeCells = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            .Cells(intFila, 2).Value = "Acciones"
            .Cells(intFila, 6).Value = "Responsable acción"
            .Range(.Cells(intFila, 5), .Cells(intFila, 6)).MergeCells = True
            .Cells(intFila, 7).Value = "F. Inicio"
            .Cells(intFila, 8).Value = "F. Fin Prev."
            .Cells(intFila, 9).Value = "F. Fin"
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 9))
            With MiRango
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = "Garamond"
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 16777164
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, _
                EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
            intInicialPlanAcciones = intFila
            PlanAcciones p_EsMitigacion, wbHoja, intFila, intFinalPlanAcciones, m_Plan, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            intFila = intFinalPlanAcciones
            
            Set MiRango = .Range(.Cells(intInicialPlanAcciones, 2), .Cells(intFinalPlanAcciones, 9))
            Recuadrar MiRango, EnumAnchoLinea.fina, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
                EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
            i = i + 1
siguiente:
            Set m_Plan = Nothing
            
        Next
        
    End With
    intFinalPlan = intFila
    p_FilaFinal = intFinalPlan
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Planes ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function AltoFila( _
                        p_TipoCelda As EnumTipoCeldaAlto, _
                        Optional p_Texto As String, _
                        Optional ByRef p_Error As String _
                        ) As Long
    Dim m_NumeroPalabras As Long
    On Error GoTo errores
    
    If p_Texto <> "" Then
        m_NumeroPalabras = NumeroDePalabras(p_Texto)
    End If
    
    If p_TipoCelda = EnumTipoCeldaAlto.RiesgoCabeceraPpal Then
        AltoFila = 21.75
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoCabeceraApartado Then
        AltoFila = 18.75
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoDatos Then
        AltoFila = 38.25
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoDescripcion Then
        If m_NumeroPalabras <= 80 Then
            AltoFila = 57
        ElseIf m_NumeroPalabras > 80 And m_NumeroPalabras <= 112 Then
            AltoFila = 113
        Else
            AltoFila = 153
        End If
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoCausaRaiz Then
        If m_NumeroPalabras <= 80 Then
            AltoFila = 57
        ElseIf m_NumeroPalabras > 80 And m_NumeroPalabras <= 112 Then
            AltoFila = 113
        Else
            AltoFila = 153
        End If
        
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoPAccionesCabecera Then
        AltoFila = 25.5
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.PortadaCambios Then
        If m_NumeroPalabras <= 18 Then
            AltoFila = 15
        ElseIf m_NumeroPalabras > 18 And m_NumeroPalabras <= 50 Then
            AltoFila = 45
        ElseIf m_NumeroPalabras > 51 And m_NumeroPalabras <= 82 Then
            AltoFila = 71
        ElseIf m_NumeroPalabras > 82 And m_NumeroPalabras <= 112 Then
            AltoFila = 113
        Else
            AltoFila = 153
        End If
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoPlanDisparador Then
        If m_NumeroPalabras <= 18 Then
            AltoFila = 15
        ElseIf m_NumeroPalabras > 18 And m_NumeroPalabras <= 50 Then
            AltoFila = 45
        ElseIf m_NumeroPalabras > 51 And m_NumeroPalabras <= 82 Then
            AltoFila = 71
        ElseIf m_NumeroPalabras > 82 And m_NumeroPalabras <= 112 Then
            AltoFila = 113
        Else
            AltoFila = 153
        End If
    ElseIf p_TipoCelda = EnumTipoCeldaAlto.RiesgoPAcciones Then
        If m_NumeroPalabras <= 25 Then
            AltoFila = 40
        ElseIf m_NumeroPalabras > 25 And m_NumeroPalabras <= 60 Then
            AltoFila = 123
        Else
            AltoFila = 160
        End If
    Else
        p_Error = "Tipo de celda no reconocido"
        Err.Raise 1000
    End If
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltoFila ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function PlanAcciones( _
                                p_EsMitigacion As EnumSiNo, _
                                ByRef wbHoja As Object, _
                                p_FilaInicial As Integer, _
                                ByRef p_FilaFinal As Integer, _
                                Optional p_Plan As Object, _
                                Optional ByRef p_Error As String _
                                ) As String
                                            
    
    Dim MiRango As Object
    Dim intFila As Integer
    Dim m_IdAccion As Variant
    Dim m_Accion As Object
    Dim m_ResponsableAccion As String
    Dim dato
    Dim m_AccionTexto As String
    On Error GoTo errores
    
    intFila = p_FilaInicial
    If p_Plan.colAcciones Is Nothing Then
        p_FilaFinal = intFila
         Exit Function
    End If
    With wbHoja
        For Each m_IdAccion In p_Plan.colAcciones
            Set m_Accion = p_Plan.colAcciones(m_IdAccion)
            intFila = intFila + 1
            m_AccionTexto = Replace(m_Accion.Accion, vbNewLine, "")
            
            m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoPAcciones, m_AccionTexto)
            If m_Alto = 0 Then
                p_Error = "El método AltoFila no ha devuelto un número válido para la acción " & m_Accion.IDAccionMitigacion
                Err.Raise 1000
            End If
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
            
            .Range(.Cells(intFila, 2), .Cells(intFila, 4)).MergeCells = True
            .Cells(intFila, 2).Value = m_AccionTexto
            Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
            With MiRango
                .WrapText = True
                .VerticalAlignment = xlTop
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            .Range(.Cells(intFila, 5), .Cells(intFila, 6)).MergeCells = True
            If InStr(1, m_Accion.ResponsableAccion, vbNewLine) <> 0 Then
                dato = Split(m_Accion.ResponsableAccion, vbNewLine)
                m_ResponsableAccion = dato(0)
            Else
                m_ResponsableAccion = m_Accion.ResponsableAccion
            End If
            .Cells(intFila, 5).Value = m_ResponsableAccion
            Set MiRango = .Range(.Cells(intFila, 5), .Cells(intFila, 5))
            With MiRango
                .WrapText = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Font.Size = 10
                .Font.Name = "Garamond"
            End With
            If IsDate(m_Accion.FechaInicio) Then
                .Cells(intFila, 7).Value = Format(m_Accion.FechaInicio, "mm/dd/yyyy")
            End If
            If IsDate(m_Accion.FechaFinPrevista) Then
                .Cells(intFila, 8).Value = Format(m_Accion.FechaFinPrevista, "mm/dd/yyyy")
            End If
            If IsDate(m_Accion.FechaFinReal) Then
                .Cells(intFila, 9).Value = Format(m_Accion.FechaFinReal, "mm/dd/yyyy")
            End If
            Set MiRango = .Range(.Cells(intFila, 7), .Cells(intFila, 9))
            With MiRango
                .Font.Size = 10
                .Font.Name = "Garamond"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                
            End With
            
            Set m_Accion = Nothing
        Next
 
        
    End With
    
    p_FilaFinal = intFila
   
      
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PlanAcciones ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Function GeneraHojaInventarioConCausaRaiz( _
                                                    p_ObjEdicion As Edicion, _
                                                    ByRef wbHoja As Object, _
                                                    Optional p_FechaCierre As String, _
                                                    Optional p_fechaRef As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    Dim MiRango As Object
    Dim intFila As Integer
    Dim intNumeroRiesgos As Long
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_Col As Scripting.Dictionary
    
    Dim m_FechaEstado As String
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EstadoTexto As String
    Dim ColumnaMaxima As Long
    Dim columnaInicial As Long
    Dim m_FechaPublicacion As String
    
    
    On Error GoTo errores
    
    ColumnaMaxima = 14
    columnaInicial = 2
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
       
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se conoce la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Not IsDate(p_fechaRef) Then
        p_fechaRef = Date
    End If
    m_FechaPublicacion = CStr(p_fechaRef)
    
    p_ObjEdicion.OrdenarColeRiesgosAscendentemente = EnumSiNo.Sí
    
    Set p_ObjEdicion.ColRiesgosPorPrioridadTodos = Nothing
    Set m_Col = p_ObjEdicion.ColRiesgosPorPrioridadTodos
    p_Error = p_ObjEdicion.Error
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Set m_Col = p_ObjEdicion.colRiesgos
        If m_Col Is Nothing Then
            p_Error = "No se encuentran los riesgos"
            Err.Raise 1000
        End If
    End If
   
    
    intNumeroRiesgos = m_Col.Count
    '---------------------------------
    ' ANCHO DE COLUMNAS
    '-------------------------------
    With wbHoja
        .Columns("A:A").ColumnWidth = 1.43
        .Columns("B:B").ColumnWidth = 14
        .Columns("C:C").ColumnWidth = 41
        .Columns("D:D").ColumnWidth = 53.14
        .Columns("E:E").ColumnWidth = 17.86
        .Columns("F:F").ColumnWidth = 7.29
        .Columns("G:G").ColumnWidth = 7.29
        .Columns("H:H").ColumnWidth = 9.86
        .Columns("I:I").ColumnWidth = 9.29
        .Columns("J:J").ColumnWidth = 16.71
        .Columns("K:K").ColumnWidth = 14.57
        .Columns("L:L").ColumnWidth = 14.14
        .Columns("M:M").ColumnWidth = 19.71
        .Columns("N:N").ColumnWidth = 13.57
    End With
    With wbHoja
        .Rows("16:16").RowHeight = 15
        .Rows("18:18").RowHeight = 15
        .Range("A16:I16").MergeCells = False
        .Range("A18:I18").MergeCells = False
    End With
    '---------------------------------
    ' FICHA DE RIESGO
    '-------------------------------
    intFila = 1
    With wbHoja
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 35
        .Range(wbHoja.Cells(intFila, columnaInicial), .Cells(intFila, ColumnaMaxima)).MergeCells = True
    End With
    intFila = 2
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraPpal)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para cabecera de la ficha de riesgo"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(wbHoja.Cells(intFila, 2), .Cells(intFila, ColumnaMaxima)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Value = "Inventario de Riesgos Detectados"
            .Font.Bold = True
            .Font.Size = 16
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, ColumnaMaxima))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
    End With
    intFila = intFila + 1 '--->3
     With wbHoja
        .Range(.Cells(intFila, columnaInicial + 1), .Cells(intFila, ColumnaMaxima)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, columnaInicial))
        With MiRango
            .Value = "Proyecto:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, columnaInicial + 1), .Cells(intFila, columnaInicial + 1))
        With MiRango
            .Value = p_ObjEdicion.Proyecto.NombreProyecto
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEAS 3 Y 4 CON PROYECTO Y FECHA
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila + 1, ColumnaMaxima))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    intFila = intFila + 1 '--->4
    With wbHoja
        .Range(.Cells(intFila, columnaInicial + 1), .Cells(intFila, ColumnaMaxima)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, columnaInicial))
        With MiRango
        
            .Value = "Fecha:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, columnaInicial + 1), .Cells(intFila, columnaInicial + 1))
        With MiRango
            If IsDate(m_FechaPublicacion) Then
                .Value = Format(m_FechaPublicacion, "mm/dd/yyyy")
                .HorizontalAlignment = xlLeft
            Else
                .Value = "NO PUBLICADO"
                .Font.Color = RGB(255, 0, 0)
            End If
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        
    End With
    '---------------------------------------------
    '   Identificación Análisis Estado Revisión
    '------------------------------------------------
    intFila = intFila + 1 '--->5
    With wbHoja
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 19.5
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, 5))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 13434828
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 2).Value = "Identificación"
        Set MiRango = .Range(.Cells(intFila, 6), .Cells(intFila, 12))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764057
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 6).Value = "Análisis"
        Set MiRango = .Range(.Cells(intFila, 13), .Cells(intFila, ColumnaMaxima))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10092543
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 13).Value = "Estado"
        
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 2 AL 5  IDENTIFICACIÓN
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, columnaInicial), .Cells(intFila, 5))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 5 AL 12   Análisis
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 6), .Cells(intFila, 12))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 13 AL 14  ESTADO
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 13), .Cells(intFila, ColumnaMaxima))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
    End With
    '---------------------------------------------
    '   Código riesgo ....
    '------------------------------------------------
    intFila = intFila + 1 '--->6
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoDatos)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoDatos"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila + 1) & ":" & CStr(intFila + 1) & "").RowHeight = m_Alto
        
        .Range(.Cells(intFila, 2), .Cells(intFila + 1, 2)).MergeCells = True
        .Cells(intFila, 2).Value = "Código Riesgo"
        .Range(.Cells(intFila, 3), .Cells(intFila + 1, 3)).MergeCells = True
        .Cells(intFila, 3).Value = "Descripción"
        
        .Range(.Cells(intFila, 4), .Cells(intFila + 1, 4)).MergeCells = True
        .Cells(intFila, 4).Value = "Causa Raíz"
        
        .Range(.Cells(intFila, 5), .Cells(intFila + 1, 5)).MergeCells = True
        .Cells(intFila, 5).Value = "Detectado Por"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 5))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 5))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE IDENTIFICACIÓN (codigo Riesgo,Descripción,Detetectado por) hasta el fin de los riesgos
        '   comienza en la fila 8 y acaba en la 8+ intNumeroRiesgos + 1
        '--------------------------------------------------------------------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + intNumeroRiesgos + 1, 5))
        With MiRango
             .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, 4), .Cells(intFila + intNumeroRiesgos + 1, 4))
        With MiRango
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + intNumeroRiesgos + 1, 5))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------------------
    '   Impacto,Plazo,Coste,Calidad,Global
    '------------------------------------------------
    'intFila=6
    With wbHoja
        .Range(.Cells(intFila, 6), .Cells(intFila, 9)).MergeCells = True
        .Cells(intFila, 6).Value = "Impacto"
        .Cells(intFila + 1, 6).Value = "Plazo"
        .Cells(intFila + 1, 7).Value = "Coste"
        .Cells(intFila + 1, 8).Value = "Calidad"
        .Cells(intFila + 1, 9).Value = "Global"
        Set MiRango = .Range(.Cells(intFila, 6), .Cells(intFila + 1, 9))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ANÁLISIS (Impacto,Plazo,Coste,Calidad,Global)
        '   comienza en la fila 6 y acaba en la 7 desde columna 6 hasta la 9
        '--------------------------------------------------------------------------------------------------------------
        
    End With
    '---------------------------------------------
    '   Vulnerabilidad,Valoración,Priorización
    '------------------------------------------------
    'intFila = 6
    With wbHoja
        .Range(.Cells(intFila, 10), .Cells(intFila + 1, 10)).MergeCells = True
        .Cells(intFila, 10).Value = "Vulnerabilidad"
        .Range(.Cells(intFila, 11), .Cells(intFila + 1, 11)).MergeCells = True
        .Cells(intFila, 11).Value = "Valoración"
        .Range(.Cells(intFila, 12), .Cells(intFila + 1, 12)).MergeCells = True
        .Cells(intFila, 12).Value = "Priorización"

        Set MiRango = .Range(.Cells(intFila, 10), .Cells(intFila + 1, 12))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ANÁLISIS (Vulnerabilidad,Valoración,Priorización)
        '   comienza en la fila 6 y acaba en la 7 desde columna 10 hasta la 12
        '--------------------------------------------------------------------------------------------------------------
        
    End With
    '---------------------------------------------
    '   Estado,Fecha,
    '------------------------------------------------
    'intFila=6
    With wbHoja
        .Range(.Cells(intFila, 13), .Cells(intFila + 1, 13)).MergeCells = True
        .Cells(intFila, 13).Value = "Estado"
        .Range(.Cells(intFila, 14), .Cells(intFila + 1, 14)).MergeCells = True
        .Cells(intFila, 14).Value = "Fecha"
        Set MiRango = .Range(.Cells(intFila, 13), .Cells(intFila + 1, 14))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ESTADO (Estado,Fecha)
        '   comienza en la fila 6 y acaba en la 7 desde columna 13 hasta la 14
        '--------------------------------------------------------------------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 13), .Cells(intFila + 1, 14))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    
     '---------------------------------
     ' RECUADRO TABLA PARTE ANÁLISIS
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 6), .Cells(8 + intNumeroRiesgos - 1, 12))
         With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------
     ' RECUADRO TABLA PARTE Estado
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 13), .Cells(8 + intNumeroRiesgos - 1, 14))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------
     ' RECUADRO TABLA PARTE Revisión
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 14), .Cells(8 + intNumeroRiesgos - 1, 14))
        With MiRango
            .Font.Bold = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    intFila = 8
    'primero los que tienen priorización ordenados por priorización y luego los no priorizados ordenados por codigo riesgo
    
    For Each m_IdRiesgo In m_Col
        Set m_ObjRiesgo = m_Col(m_IdRiesgo)
        If m_ObjRiesgo Is Nothing Then GoTo siguiente
        m_Estado = m_ObjRiesgo.EstadoEnum
        
        
        
        With wbHoja
            .Range(.Cells(intFila, 2), .Cells(intFila, 2)).Value = m_ObjRiesgo.CodigoRiesgo
            .Range(.Cells(intFila, 3), .Cells(intFila, 3)).Value = m_ObjRiesgo.Descripcion
            .Range(.Cells(intFila, 4), .Cells(intFila, 4)).Value = m_ObjRiesgo.CausaRaiz
            .Range(.Cells(intFila, 5), .Cells(intFila, 5)).Value = m_ObjRiesgo.DetectadoPor
            .Range(.Cells(intFila, 6), .Cells(intFila, 6)).Value = m_ObjRiesgo.Plazo
            .Range(.Cells(intFila, 7), .Cells(intFila, 7)).Value = m_ObjRiesgo.Coste
            .Range(.Cells(intFila, 8), .Cells(intFila, 8)).Value = m_ObjRiesgo.Calidad
            .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Value = m_ObjRiesgo.ImpactoGlobal
            .Range(.Cells(intFila, 10), .Cells(intFila, 10)).Value = m_ObjRiesgo.Vulnerabilidad
            .Range(.Cells(intFila, 11), .Cells(intFila, 11)).Value = m_ObjRiesgo.Valoracion
            .Range(.Cells(intFila, 12), .Cells(intFila, 12)).Value = m_ObjRiesgo.Priorizacion
            
            If m_Estado = EnumRiesgoEstado.Retirado Then
                m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                m_FechaEstado = m_ObjRiesgo.FechaRetirado
            ElseIf m_Estado = EnumRiesgoEstado.Aceptado Then
                m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                m_FechaEstado = m_ObjRiesgo.FechaMitigacionAceptar
            
            ElseIf m_Estado = EnumRiesgoEstado.Detectado Then
                If p_FechaCierre <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = m_ObjRiesgo.FechaDetectado
                End If
            ElseIf m_Estado = EnumRiesgoEstado.Materializado Then
                If p_FechaCierre <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = "Materializado"
                    m_FechaEstado = m_ObjRiesgo.FechaMaterializado
                End If
            Else
                If p_FechaCierre <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = m_ObjRiesgo.FechaEstado
                End If
            End If
            
            
            .Range(.Cells(intFila, 13), .Cells(intFila, 13)).Value = m_EstadoTexto
            .Range(.Cells(intFila, 14), .Cells(intFila, 14)).Value = Format(m_FechaEstado, "mm/dd/yyyy")
            .Range(.Cells(intFila, 2), .Cells(intFila, 5)).WrapText = True
            .Range(.Cells(intFila, 3), .Cells(intFila, 5)).HorizontalAlignment = xlLeft
            .Range(.Cells(intFila, 6), .Cells(intFila, 10)).HorizontalAlignment = xlCenter
            .Range(.Cells(intFila, 11), .Cells(intFila, 11)).HorizontalAlignment = xlCenter
            .Range(.Cells(intFila, 12), .Cells(intFila, 12)).HorizontalAlignment = xlCenter
            
            
            
            .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").EntireRow.AutoFit
           
        End With
        Set m_ObjRiesgo = Nothing
        
        intFila = intFila + 1
siguiente:
    Next
    With wbHoja
        Set MiRango = .Range(.Cells(6, 2), .Cells(7, 5))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(6, 2), .Cells(7, 5))
       Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ExcelInforme.GeneraHojaInventarioConCausaRaiz ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Function GeneraHojaInventarioSinCausaRaiz( _
                                                    p_ObjEdicion As Edicion, _
                                                    ByRef wbHoja As Object, _
                                                    Optional p_FechaCierre As String, _
                                                    Optional p_fechaRef As String, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String

    Dim MiRango As Object
    Dim intFila As Integer
    Dim intNumeroRiesgos As Integer
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    
    Dim m_FechaPublicacion As String
    Dim m_ObjColRiesgosPorPriorizacion As Scripting.Dictionary
    Dim m_FechaEstado As String
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EstadoTexto As String
    Dim ColumnaMaxima As Long
    Dim columnaInicial As Long
    
    
    On Error GoTo errores
    '--------------------------------------
    ' COMPROBACIÓN DE LOS DATOS
    '--------------------------------------
        
    If p_ObjEdicion Is Nothing Then
        p_Error = "No se conoce la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Not IsDate(p_fechaRef) Then
        p_fechaRef = Date
    End If
    m_FechaPublicacion = CStr(p_fechaRef)
    
    
    p_ObjEdicion.OrdenarColeRiesgosAscendentemente = EnumSiNo.Sí
    
    Set m_ObjColRiesgosPorPriorizacion = p_ObjEdicion.ColRiesgosPorPrioridadTodos
    If m_ObjColRiesgosPorPriorizacion Is Nothing Then
         p_Error = p_ObjEdicion.Error
         If p_Error <> "" Then
            Err.Raise 1000
         End If
    End If
    
    
    intNumeroRiesgos = p_ObjEdicion.colRiesgos.Count
    '---------------------------------
    ' ANCHO DE COLUMNAS
    '-------------------------------
    With wbHoja
        .Columns("A:A").ColumnWidth = 1.43
        .Columns("B:B").ColumnWidth = 14
        .Columns("C:C").ColumnWidth = 93.14
        .Columns("D:D").ColumnWidth = 17.86
        .Columns("E:E").ColumnWidth = 7.29
        .Columns("F:F").ColumnWidth = 7.29
        .Columns("G:G").ColumnWidth = 9.86
        .Columns("H:H").ColumnWidth = 9.29
        .Columns("I:I").ColumnWidth = 16.71
        .Columns("J:J").ColumnWidth = 14.57
        .Columns("K:K").ColumnWidth = 14.14
        .Columns("L:L").ColumnWidth = 19.71
        .Columns("M:M").ColumnWidth = 13.57
    End With
    With wbHoja
        .Rows("16:16").RowHeight = 15
        .Rows("18:18").RowHeight = 15
        .Range("A16:I16").MergeCells = False
        .Range("A18:I18").MergeCells = False
    End With
   
    '---------------------------------
    ' FICHA DE RIESGO
    '-------------------------------
    intFila = 1
    With wbHoja
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 35
        .Range(wbHoja.Cells(intFila, 2), .Cells(intFila, 13)).MergeCells = True
    End With
    intFila = 2
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraPpal)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraPpal"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        .Range(wbHoja.Cells(intFila, 2), .Cells(intFila, 13)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .HorizontalAlignment = xlCenter
            .Value = "Inventario de Riesgos Detectados"
            .Font.Bold = True
            .Font.Size = 16
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 13))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
    End With
    '---------------------------------
    ' Proyecto:
    ' Fecha:
    '-------------------------------
    intFila = intFila + 1 '--->3
     With wbHoja
        .Range(.Cells(intFila, 3), .Cells(intFila, 13)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
            .Value = "Proyecto:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 3))
        With MiRango
            .Value = p_ObjEdicion.Proyecto.NombreProyecto
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEAS 3 Y 4 CON PROYECTO Y FECHA
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 13))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    intFila = intFila + 1 '--->4
    With wbHoja
        .Range(.Cells(intFila, 3), .Cells(intFila, 13)).MergeCells = True
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 2))
        With MiRango
        
            .Value = "Fecha:"
            .Font.Bold = True
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With

        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila, 3))

        With MiRango
            If IsDate(m_FechaPublicacion) Then
                .Value = Format(m_FechaPublicacion, "mm/dd/yyyy")
                .HorizontalAlignment = xlLeft
            Else
                .Value = "NO PUBLICADO"
                .Font.Color = RGB(255, 0, 0)
            End If
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        
    End With
     '---------------------------------------------
    '   Identificación Análisis Estado Revisión
    '------------------------------------------------
    intFila = intFila + 1 '--->5
    With wbHoja
        
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = 19.5
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 4))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 13434828
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 2).Value = "Identificación"
        Set MiRango = .Range(.Cells(intFila, 5), .Cells(intFila, 11))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764057
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 5).Value = "Análisis"
        Set MiRango = .Range(.Cells(intFila, 12), .Cells(intFila, 13))
        With MiRango
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 14
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10092543
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(intFila, 12).Value = "Estado"
        
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 2 AL 4  IDENTIFICACIÓN
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila, 4))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 5 AL 11   Análisis
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 5), .Cells(intFila, 11))
       Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        '---------------------------------------------------
        ' BORDE DE LAS LÍNEA 5 del 12 AL 13  ESTADO
        '---------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 12), .Cells(intFila, 13))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        
    End With
    '---------------------------------------------
    '   Código riesgo ....
    '------------------------------------------------
    intFila = intFila + 1 '--->6
    With wbHoja
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoCabeceraApartado)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoCabeceraApartado"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila) & ":" & CStr(intFila) & "").RowHeight = m_Alto
        m_Alto = AltoFila(EnumTipoCeldaAlto.RiesgoDatos)
        If m_Alto = 0 Then
            p_Error = "El método AltoFila no ha devuelto un número válido para RiesgoDatos"
            Err.Raise 1000
        End If
        .Rows("" & CStr(intFila + 1) & ":" & CStr(intFila + 1) & "").RowHeight = m_Alto
        .Range(.Cells(intFila, 2), .Cells(intFila + 1, 2)).MergeCells = True
        .Cells(intFila, 2).Value = "Código Riesgo"
        .Range(.Cells(intFila, 3), .Cells(intFila + 1, 3)).MergeCells = True
        .Cells(intFila, 3).Value = "Descripción"
        .Range(.Cells(intFila, 4), .Cells(intFila + 1, 4)).MergeCells = True
        .Cells(intFila, 4).Value = "Detectado Por"
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 4))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + 1, 4))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, p_Error:=p_Error
        
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE IDENTIFICACIÓN (codigo Riesgo,Descripción,Detetectado por) hasta el fin de los riesgos
        '   comienza en la fila 8 y acaba en la 8+ intNumeroRiesgos + 1
        '--------------------------------------------------------------------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + intNumeroRiesgos + 1, 4))
        With MiRango
             .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
        End With
        Set MiRango = .Range(.Cells(intFila, 3), .Cells(intFila + intNumeroRiesgos + 1, 3))
        With MiRango
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        Set MiRango = .Range(.Cells(intFila, 2), .Cells(intFila + intNumeroRiesgos + 1, 4))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------------------
    '   Impacto,Plazo,Coste,Calidad,Global
    '------------------------------------------------
    'intFila=6
    With wbHoja
        .Range(.Cells(intFila, 5), .Cells(intFila, 8)).MergeCells = True
        .Cells(intFila, 5).Value = "Impacto"
        .Cells(intFila + 1, 5).Value = "Plazo"
        .Cells(intFila + 1, 6).Value = "Coste"
        .Cells(intFila + 1, 7).Value = "Calidad"
        .Cells(intFila + 1, 8).Value = "Global"
        Set MiRango = .Range(.Cells(intFila, 5), .Cells(intFila + 1, 8))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ANÁLISIS (Impacto,Plazo,Coste,Calidad,Global)
        '   comienza en la fila 6 y acaba en la 7 desde columna 4 hasta la 7
        '--------------------------------------------------------------------------------------------------------------
        
    End With
    '---------------------------------------------
    '   Vulnerabilidad,Valoración,Priorización
    '------------------------------------------------
    'intFila=6
    With wbHoja
        .Range(.Cells(intFila, 9), .Cells(intFila + 1, 9)).MergeCells = True
        .Cells(intFila, 9).Value = "Vulnerabilidad"
        .Range(.Cells(intFila, 10), .Cells(intFila + 1, 10)).MergeCells = True
        .Cells(intFila, 10).Value = "Valoración"
        .Range(.Cells(intFila, 11), .Cells(intFila + 1, 11)).MergeCells = True
        .Cells(intFila, 11).Value = "Priorización"

        Set MiRango = .Range(.Cells(intFila, 9), .Cells(intFila + 1, 11))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ANÁLISIS (Vulnerabilidad,Valoración,Priorización)
        '   comienza en la fila 6 y acaba en la 7 desde columna 8 hasta la 11
        '--------------------------------------------------------------------------------------------------------------
        
    End With
    '---------------------------------------------
    '   Estado,Fecha,
    '------------------------------------------------
    'intFila=6
    With wbHoja
        .Range(.Cells(intFila, 12), .Cells(intFila + 1, 12)).MergeCells = True
        .Cells(intFila, 12).Value = "Estado"
        .Range(.Cells(intFila, 13), .Cells(intFila + 1, 13)).MergeCells = True
        .Cells(intFila, 13).Value = "Fecha"
        Set MiRango = .Range(.Cells(intFila, 12), .Cells(intFila + 1, 13))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        '--------------------------------------------------------------------------------------------------------------
        ' BORDE DE LAS TROZO DE ESTADO (Estado,Fecha)
        '   comienza en la fila 6 y acaba en la 7 desde columna 12 hasta la 13
        '--------------------------------------------------------------------------------------------------------------
        Set MiRango = .Range(.Cells(intFila, 12), .Cells(intFila + 1, 13))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    
     '---------------------------------
     ' RECUADRO TABLA PARTE ANÁLISIS
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 5), .Cells(8 + intNumeroRiesgos - 1, 11))
         With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------
     ' RECUADRO TABLA PARTE Estado
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 12), .Cells(8 + intNumeroRiesgos - 1, 13))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    '---------------------------------
     ' RECUADRO TABLA PARTE Revisión
     '----------------------------------
     'intFila=6
     With wbHoja
        Set MiRango = .Range(.Cells(8, 13), .Cells(8 + intNumeroRiesgos - 1, 13))
        With MiRango
            .Font.Bold = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = "Garamond"
            .Font.Size = 10
            .WrapText = True
        End With
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    intFila = 8
    If p_ObjEdicion.TieneRiesgos = EnumSiNo.Sí Then
        If p_ObjEdicion.TodosLosRiesgosPriorizados = EnumSiNo.No Then
            p_Error = p_ObjEdicion.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "No están todos los riesgos priorizados"
            Err.Raise 1000
        End If
        For Each m_IdRiesgo In m_ObjColRiesgosPorPriorizacion.Keys
            
            Set m_ObjRiesgo = m_ObjColRiesgosPorPriorizacion(m_IdRiesgo)
            m_Estado = m_ObjRiesgo.EstadoEnum
            
            
            With wbHoja
                .Range(.Cells(intFila, 2), .Cells(intFila, 2)).Value = m_ObjRiesgo.CodigoRiesgo
                .Range(.Cells(intFila, 3), .Cells(intFila, 3)).Value = m_ObjRiesgo.Descripcion
                .Range(.Cells(intFila, 4), .Cells(intFila, 4)).Value = m_ObjRiesgo.DetectadoPor
                .Range(.Cells(intFila, 5), .Cells(intFila, 5)).Value = m_ObjRiesgo.Plazo
                .Range(.Cells(intFila, 6), .Cells(intFila, 6)).Value = m_ObjRiesgo.Coste
                .Range(.Cells(intFila, 7), .Cells(intFila, 7)).Value = m_ObjRiesgo.Calidad
                .Range(.Cells(intFila, 8), .Cells(intFila, 8)).Value = m_ObjRiesgo.ImpactoGlobal
                .Range(.Cells(intFila, 9), .Cells(intFila, 9)).Value = m_ObjRiesgo.Vulnerabilidad
                .Range(.Cells(intFila, 10), .Cells(intFila, 10)).Value = m_ObjRiesgo.Valoracion
                .Range(.Cells(intFila, 11), .Cells(intFila, 11)).Value = m_ObjRiesgo.Priorizacion
                If m_Estado = EnumRiesgoEstado.Retirado Then
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = m_ObjRiesgo.FechaRetirado
                ElseIf m_Estado = EnumRiesgoEstado.Aceptado Then
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = m_ObjRiesgo.FechaMitigacionAceptar
                
                ElseIf m_Estado = EnumRiesgoEstado.Detectado Then
                    If p_FechaCierre <> "" Then
                        m_EstadoTexto = "Cerrado"
                        m_FechaEstado = p_FechaCierre
                    Else
                        m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                        m_FechaEstado = m_ObjRiesgo.FechaDetectado
                    End If
                ElseIf m_Estado = EnumRiesgoEstado.Materializado Then
                    If p_FechaCierre <> "" Then
                        m_EstadoTexto = "Cerrado"
                        m_FechaEstado = p_FechaCierre
                    Else
                        m_EstadoTexto = "Materializado"
                        m_FechaEstado = m_ObjRiesgo.FechaMaterializado
                    End If
                Else
                    If p_FechaCierre <> "" Then
                        m_EstadoTexto = "Cerrado"
                        m_FechaEstado = m_FechaPublicacion
                    Else
                        m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                        m_FechaEstado = m_ObjRiesgo.FechaEstado
                    End If
                End If
                .Range(.Cells(intFila, 12), .Cells(intFila, 12)).Value = m_EstadoTexto
                .Range(.Cells(intFila, 13), .Cells(intFila, 13)).Value = Format(m_FechaEstado, "mm/dd/yyyy")
                .Range(.Cells(intFila, 3), .Cells(intFila, 3)).WrapText = True
                .Range(.Cells(intFila, 4), .Cells(intFila, 4)).WrapText = True
                .Range(.Cells(intFila, 3), .Cells(intFila, 5)).HorizontalAlignment = xlLeft
                .Range(.Cells(intFila, 6), .Cells(intFila, 9)).HorizontalAlignment = xlCenter
                .Range(.Cells(intFila, 10), .Cells(intFila, 10)).HorizontalAlignment = xlCenter
                .Range(.Cells(intFila, 11), .Cells(intFila, 11)).HorizontalAlignment = xlCenter
            End With
            Set m_ObjRiesgo = Nothing
            intFila = intFila + 1

        Next
    End If
    
    With wbHoja
        Set MiRango = .Range(.Cells(6, 2), .Cells(7, 4))
        With MiRango
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Font.Name = "Garamond"
            
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        Set MiRango = .Range(.Cells(6, 2), .Cells(7, 4))
        Recuadrar MiRango, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, EnumAnchoLinea.Gruesa, _
            EnumAnchoLinea.Gruesa, EnumAnchoLinea.fina, EnumAnchoLinea.fina, p_Error:=p_Error
        
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraHojaInventarioEdicion ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Public Function Recuadrar( _
                            p_Rango As Object, _
                            Optional p_AnchoLineaArriba As EnumAnchoLinea, _
                            Optional p_AnchoLineaAbajo As EnumAnchoLinea, _
                            Optional p_AnchoLineaIzquierda As EnumAnchoLinea, _
                            Optional p_AnchoLineaDerecha As EnumAnchoLinea, _
                            Optional p_AnchoLineaVerticalInt As EnumAnchoLinea, _
                            Optional p_AnchoLineaHorizontalInt As EnumAnchoLinea, _
                            Optional p_Error As String _
                            ) As String
    
    On Error GoTo errores
    
    
    
    With p_Rango
        If p_AnchoLineaIzquierda = Empty Then
            .Borders(xlEdgeLeft).LineStyle = xlNone
        Else
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaIzquierda
            End With
        End If
        If p_AnchoLineaDerecha = Empty Then
            .Borders(xlEdgeRight).LineStyle = xlNone
        Else
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaDerecha
            End With
        End If
        If p_AnchoLineaArriba = Empty Then
            .Borders(xlEdgeTop).LineStyle = xlNone
        Else
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaArriba
            End With
        End If
        If p_AnchoLineaAbajo = Empty Then
            .Borders(xlEdgeBottom).LineStyle = xlNone
        Else
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaAbajo
            End With
        End If
        If p_AnchoLineaHorizontalInt = Empty Then
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        Else
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaHorizontalInt
            End With
        End If
        If p_AnchoLineaVerticalInt = Empty Then
            .Borders(xlInsideVertical).LineStyle = xlNone
        Else
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = p_AnchoLineaVerticalInt
            End With
        End If
        
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Recuadrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

