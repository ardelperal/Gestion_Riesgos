Attribute VB_Name = "WordInforme"


Option Compare Database
Option Explicit

' ============================================================================
' Módulo: WordInforme
' Descripción: Funciones para generar informes de publicación de riesgos en Word
' Autor: Sistema de Gestión de Riesgos
' Fecha: 2025-11-28
' ============================================================================

' Función principal para generar informe en Word
Public Function GenerarInformeWord( _
                                    p_ObjEdicion As Edicion, _
                                    Optional p_EnPDF As EnumSiNo = EnumSiNo.No, _
                                    Optional p_FechaCierre As String, _
                                    Optional p_FechaPublicacion As String, _
                                    Optional ByRef p_Error As String, _
                                    Optional ByRef p_EnWord As EnumSiNo = EnumSiNo.No _
                                    ) As String

    Dim appWord As Word.Application
    Dim docWord As Word.Document
    Dim m_ObjProyecto As Proyecto
    Dim m_URLWord As String
    Dim m_URLPDF As String
    Dim FSO As Scripting.FileSystemObject
    Dim bExportPDF As Boolean
    
    On Error GoTo errores
    
    ' Validaciones
    If p_ObjEdicion Is Nothing Then
        p_Error = "La edición está sin indicar"
        Err.Raise 1000
    End If
    
    ' Obtener proyecto y entorno
    Set m_ObjProyecto = Constructor.getProyecto(p_ObjEdicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    ' Preparar fechas
    If Not IsDate(p_FechaCierre) Then p_FechaCierre = Date
    If Not IsDate(p_FechaPublicacion) Then p_FechaPublicacion = Date
    
    
    Set FSO = New Scripting.FileSystemObject
    m_URLWord = getURLWord(p_ObjEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    bExportPDF = (p_EnWord = EnumSiNo.No) Or (p_EnPDF = EnumSiNo.Sí)
    
' Verificar que no está abierto (igual que ExcelInforme)
    If Not FSO.FileExists(m_URLWord) Then
        If FicheroAbierto(m_URLWord) Then
            p_Error = "Tiene abierto un informe anterior"
            Err.Raise 1000
        End If
    End If
    
    If bExportPDF Then
        m_URLPDF = FSO.GetParentFolderName(m_URLWord) & "\" & FSO.GetBaseName(m_URLWord) & ".pdf"
        If FSO.FileExists(m_URLPDF) Then
            If FicheroAbierto(m_URLPDF) Then
                p_Error = "Tiene abierto un informe anterior"
                Err.Raise 1000
            End If
        End If
    End If
    
    Avance "Preparando plantilla Word"
    
    ' Copiar plantilla (igual que ExcelInforme)
    Dim plantillaOrigen As String
    plantillaOrigen = m_ObjEntorno.URLPlantillaWord
    FSO.CopyFile plantillaOrigen, m_URLWord, True
    
' Crear aplicación Word
    Set appWord = New Word.Application
    appWord.Visible = True
    appWord.DisplayAlerts = 0  ' wdAlertsNone
    
    ' Abrir documento
    Set docWord = appWord.Documents.Open(m_URLWord)
    
    Avance "Generando página inicial"
    GeneraPaginaInicialWord p_ObjEdicion, m_ObjProyecto, docWord, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando portada"
    GeneraPortadaWord p_ObjEdicion, m_ObjProyecto, docWord, p_FechaPublicacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando pie de página"
    GeneraPiePaginaWord p_ObjEdicion, m_ObjProyecto, docWord, p_FechaPublicacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando cuadro de control"
    GeneraCuadroControlWord p_ObjEdicion, m_ObjProyecto, docWord, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando control de cambios"
    GeneraControlCambiosWord p_ObjEdicion, docWord, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando inventario de riesgos"
    GeneraInventarioRiesgosWord p_ObjEdicion, m_ObjProyecto, docWord, p_FechaPublicacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Avance "Generando fichas de riesgo"
    GeneraFichasRiesgoWord p_ObjEdicion, m_ObjProyecto, docWord, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' Guardar documento
    Avance "Guardando documento Word"
    docWord.Save
    
    If bExportPDF Then
        Avance "Exportando a PDF"
        docWord.Save
        docWord.Close
        On Error Resume Next
        Set docWord = appWord.Documents.Open(m_URLWord)
        If Err.Number <> 0 Then
            Err.Raise 1000
        End If
        docWord.ExportAsFixedFormat m_URLPDF, 17
        m_URLWord = m_URLPDF
    Else
        m_URLWord = m_URLWord
    End If
    
' Cerrar documento y aplicación
    docWord.Close False
    Set docWord = Nothing
    appWord.Quit
    Set appWord = Nothing
    
    GenerarInformeWord = m_URLWord
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarInformeWord ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
    
    ' Limpiar objetos
    If Not docWord Is Nothing Then
        docWord.Close False
        Set docWord = Nothing
    End If
    If Not appWord Is Nothing Then
        appWord.Quit
        Set appWord = Nothing
    End If
End Function

Private Function getURLWord( _
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
        getURLWord = m_ObjEntorno.URLDirectorioLocal & m_Cod & "V" & _
                     Format(p_Edicion.Edicion, "00") & ".docx"
    Else
        getURLWord = m_ObjEntorno.URLDirectorioLocal & m_Cod & "-" & p_Edicion.Edicion & ".docx"
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getURLWord ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



 

' Genera la tabla de inventario de riesgos



' Genera la portada del documento
Private Function GeneraPortadaWord( _
                                    p_ObjEdicion As Edicion, _
                                    p_ObjProyecto As Proyecto, _
                                    docWord As Word.Document, _
                                    Optional p_FechaPublicacion As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    On Error GoTo errores
    
    Dim nom As String
    Dim exp As String
    nom = UCase(p_ObjProyecto.NombreProyecto)
    exp = UCase(p_ObjProyecto.Proyecto)

    ReemplazarMarcador docWord, "Proyecto_Proyecto", exp
    ReemplazarMarcador docWord, "Proyecto_NombreProyecto", nom
    ReemplazarMarcador docWord, "Proyecto_JefeProyecto", p_ObjEdicion.Elaborado
    ReemplazarMarcador docWord, "Proyecto_Cliente", p_ObjProyecto.Cliente
    ReemplazarMarcador docWord, "Portada_Proyecto", exp
    ReemplazarMarcador docWord, "Portada_NombreProyecto", nom

   
    
    GeneraPortadaWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraPortadaWord ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
End Function



' Genera el pie de página



' Genera el cuadro de control de ediciones
Private Function GeneraCuadroControlWord( _
                                          p_ObjEdicion As Edicion, _
                                          p_ObjProyecto As Proyecto, _
                                          docWord As Word.Document, _
                                          Optional ByRef p_Error As String _
                                          ) As String

    Dim tabla As Word.Table
    Dim fila As Word.Row
    Dim colEdiciones As Scripting.Dictionary
    Dim Edicion As Edicion
    Dim i As Integer
    
    On Error GoTo errores
    
    ' Buscar la tabla de cuadro de control
    ' Intentar por marcador primero
    If docWord.Bookmarks.Exists("TablaCuadroControl") Then
        Set tabla = docWord.Bookmarks("TablaCuadroControl").Range.Tables(1)
    ElseIf docWord.Tables.Count >= 2 Then
        ' Si no hay marcador, asumir que es la tabla 2 (la 1 suele ser la portada)
        Set tabla = docWord.Tables(2)
    ElseIf docWord.Tables.Count >= 1 Then
        ' Fallback a tabla 1 si solo hay una
        Set tabla = docWord.Tables(1)
    Else
        p_Error = "No se encontró la tabla de cuadro de control en la plantilla"
        Err.Raise 1000
    End If
    
    Set tabla = tabla ' Asegurar referencia
    
    ' Obtener todas las ediciones del proyecto
    Set colEdiciones = p_ObjProyecto.colEdiciones
    If colEdiciones Is Nothing Then
        p_Error = "No se pudieron cargar las ediciones del proyecto"
        Err.Raise 1000
    End If
    
    Dim filaHeader As Integer
    filaHeader = BuscarFilaPorTexto(tabla, "EDIción")
    If filaHeader = 0 Then filaHeader = BuscarFilaPorTexto(tabla, "EDICION")
    Dim nFilaBase As Integer
    nFilaBase = filaHeader + 1
   
    If tabla.Rows.Count < nFilaBase Then tabla.Rows.Add
    Dim filaBase As Word.Row
    Set filaBase = GetRowByIndexSafe(tabla, nFilaBase)
    LimpiarContenidoFila filaBase
    Dim nFila As Integer
    nFila = nFilaBase

    ' Agregar una fila por cada edición usando Table.Cell(row,col)
    Dim kEd As Variant
    Dim esPrimera As Boolean
    Dim nFilaUltima As Integer
    esPrimera = True
    For Each kEd In colEdiciones
        Set Edicion = colEdiciones(kEd)
        
        If esPrimera Then
            esPrimera = False
        Else
            tabla.Rows.Add
            nFila = nFila + 1
            CopiarFormatoFila filaBase, GetRowByIndexSafe(tabla, nFila)
        End If
        
        tabla.Cell(nFila, 1).Range.Text = Edicion.Edicion
        tabla.Cell(nFila, 2).Range.Text = Format(Edicion.FechaEdicion, "dd/mm/yyyy")
        tabla.Cell(nFila, 3).Range.Text = Edicion.Elaborado
        tabla.Cell(nFila, 4).Range.Text = Edicion.Revisado
        tabla.Cell(nFila, 5).Range.Text = Edicion.Aprobado
        AplicarBordeSuperiorFila GetRowByIndexSafe(tabla, nFila), filaBase
        AplicarGrillaCompletaFila GetRowByIndexSafe(tabla, nFila)
        AplicarBordesVerticalesFilaDesdeBase GetRowByIndexSafe(tabla, nFila), filaBase
        nFilaUltima = nFila
    Next kEd

    If nFilaUltima > 0 Then
        AplicarEstiloUltimaFila GetRowByIndexSafe(tabla, nFilaUltima)
    End If
    
    GeneraCuadroControlWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraCuadroControlWord ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Function GeneraPaginaInicialWord( _
                                        p_ObjEdicion As Edicion, _
                                        p_ObjProyecto As Proyecto, _
                                        docWord As Word.Document, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    On Error GoTo errores

    Dim nom As String
    Dim exp As String
    nom = UCase(p_ObjProyecto.NombreProyecto)
    exp = UCase(p_ObjProyecto.Proyecto)

    ReemplazarMarcador docWord, "Inicio_NombreProyecto", nom
    ReemplazarMarcador docWord, "Inicio_Proyecto", exp

    ReemplazarTextoEnRango docWord.Content, "[NOMBRE_PROYECTO]", nom
    ReemplazarTextoEnRango docWord.Content, "[CODIGO_EXPEDIENTE]", exp

    GeneraPaginaInicialWord = "OK"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "GeneraPaginaInicialWord: " & Err.Number & " - " & Err.Description
    End If
End Function

' Genera la tabla de inventario de riesgos
Private Function GeneraInventarioRiesgosWord( _
                                             p_ObjEdicion As Edicion, _
                                             p_ObjProyecto As Proyecto, _
                                             docWord As Word.Document, _
                                             p_FechaPublicacion As String, _
                                             Optional ByRef p_Error As String _
                                             ) As String

    Dim tabla As Word.Table
    Dim fila As Word.Row
    Dim colRiesgos As Scripting.Dictionary
    Dim riesgo As riesgo
    Dim i As Integer
    
    On Error GoTo errores
    
    ' Reemplazar marcadores de cabecera del inventario
    Dim nom As String
    Dim fechaPub As String
    nom = UCase(p_ObjProyecto.NombreProyecto)
    If Not IsDate(p_FechaPublicacion) Then
        fechaPub = Format(Date, "dd/mm/yyyy")
    Else
        fechaPub = Format(p_FechaPublicacion, "dd/mm/yyyy")
    End If
    ' Compatibilidad con versiones anteriores
    ReemplazarMarcador docWord, "InventarioProyecto", nom
    ' Nuevos marcadores especificados
    ReemplazarMarcador docWord, "inventarioProyecto", nom
    ReemplazarMarcador docWord, "infentarioFecha", fechaPub
    ' Usar p_FechaPublicacion que no está disponible en la firma actual, hay que añadirlo
    ' O usar Date si no está disponible, pero mejor pasar la fecha.
    ' Espera, GeneraInventarioRiesgosWord no recibe p_ObjProyecto ni p_FechaPublicacion en la versión anterior.
    ' Tengo que actualizar la firma de la función primero.

    ' Intentar por marcador primero
    If docWord.Bookmarks.Exists("TablaInventarioRiesgos") Then
        Set tabla = docWord.Bookmarks("TablaInventarioRiesgos").Range.Tables(1)
    ElseIf docWord.Tables.Count >= 4 Then
        ' Si no hay marcador, asumir que es la tabla 4 (Portada + Cuadro + Control + Inventario)
        Set tabla = docWord.Tables(4)
    ElseIf docWord.Tables.Count >= 3 Then
        ' Fallback
        Set tabla = docWord.Tables(3)
    Else
        p_Error = "No se encontró la tabla de inventario de riesgos en la plantilla"
        Err.Raise 1000
    End If
    
    Set tabla = tabla ' Asegurar referencia
    
    ' Obtener riesgos de la edición
    Set colRiesgos = p_ObjEdicion.colRiesgos
    If colRiesgos Is Nothing Then
        p_Error = "No se pudieron cargar los riesgos de la edición"
        Err.Raise 1000
    End If
    
    ' Determinar dinámicamente hasta qué fila mantener como encabezado
    Dim filasEncabezado As Integer
    filasEncabezado = BuscarFilaPorTexto(tabla, "código riesgo")
    If filasEncabezado = 0 Then filasEncabezado = BuscarFilaPorTexto(tabla, "Código riesgo")
    If filasEncabezado = 0 Then filasEncabezado = BuscarFilaPorTexto(tabla, "Codigo riesgo")
    ' Mantener la primera fila de datos de la plantilla justo debajo de los titulares
    Dim idxFilaBase As Integer
    idxFilaBase = filasEncabezado + 1
    Do While tabla.Rows.Count > idxFilaBase
        DeleteRowSafe tabla, idxFilaBase + 1
    Loop
    ' Si no existiera, crearla
    If tabla.Rows.Count < idxFilaBase Then
        tabla.Rows.Add
    End If
    Dim filaBase As Word.Row
    Dim nFila As Integer
    Set filaBase = GetRowByIndexSafe(tabla, idxFilaBase)
    LimpiarContenidoFila filaBase
    nFila = idxFilaBase
    
    Dim k As Variant
    
    ' Agregar filas reutilizando la fila base para el primer riesgo
    Dim esPrimera As Boolean
    esPrimera = True
    For Each k In colRiesgos
        Set riesgo = colRiesgos(k)
        
        If esPrimera Then
            esPrimera = False
        Else
            tabla.Rows.Add
            nFila = nFila + 1
            CopiarFormatoFila filaBase, GetRowByIndexSafe(tabla, nFila)
        End If
        
        ' Mapeo de columnas basado en la imagen
        ' 1: Código riesgo
        tabla.Cell(nFila, 1).Range.Text = riesgo.CodigoRiesgo
        ' 2: Descripción
        tabla.Cell(nFila, 2).Range.Text = riesgo.Descripcion
        ' 3: Causa raíz
        tabla.Cell(nFila, 3).Range.Text = riesgo.CausaRaiz
        ' 4: Detectado Por
        tabla.Cell(nFila, 4).Range.Text = riesgo.DetectadoPor
        ' 5: Impacto Plazo (asumiendo orden)
        tabla.Cell(nFila, 5).Range.Text = riesgo.Plazo
        tabla.Cell(nFila, 6).Range.Text = riesgo.Coste
        tabla.Cell(nFila, 7).Range.Text = riesgo.Calidad
        ' 8: Impacto Global
        tabla.Cell(nFila, 8).Range.Text = riesgo.ImpactoGlobal
        ' 9: Vulnerabilidad
        tabla.Cell(nFila, 9).Range.Text = riesgo.Vulnerabilidad
        ' 10: Valoración
        tabla.Cell(nFila, 10).Range.Text = riesgo.Valoracion
        ' 11: Priorización
        tabla.Cell(nFila, 11).Range.Text = riesgo.Priorizacion
        tabla.Cell(nFila, 12).Range.Text = riesgo.ESTADOCalculadoTexto
        ' 13: Fecha
        tabla.Cell(nFila, 13).Range.Text = Format(riesgo.FechaEstado, "dd/mm/yyyy")
        AplicarBordeSuperiorFila GetRowByIndexSafe(tabla, nFila), filaBase
        AplicarGrillaCompletaFila GetRowByIndexSafe(tabla, nFila)
        AplicarBordesVerticalesFilaDesdeBase GetRowByIndexSafe(tabla, nFila), filaBase
    Next k
    
    GeneraInventarioRiesgosWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraInventarioRiesgosWord ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
End Function

' Genera el pie de página
Private Function GeneraPiePaginaWord( _
                                     p_ObjEdicion As Edicion, _
                                     p_ObjProyecto As Proyecto, _
                                     docWord As Word.Document, _
                                     p_FechaPublicacion As String, _
                                     Optional ByRef p_Error As String _
                                     ) As String

    On Error GoTo errores
    
    ' Reemplazar marcadores en el pie de página
    ' Nota: Los marcadores en encabezados/pies de página son accesibles desde la colección global Bookmarks
    
    ReemplazarMarcador docWord, "PieCodigo", p_ObjProyecto.CodigoDocumento
    ReemplazarMarcador docWord, "PieEdicion", p_ObjEdicion.Edicion
    ReemplazarMarcador docWord, "PieFecha", Format(p_FechaPublicacion, "dd/mm/yyyy")
    
    GeneraPiePaginaWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GeneraPiePaginaWord ha producido el error nº: " & Err.Number & _
                  vbNewLine & "Detalle: " & Err.Description
    End If
End Function
' Función auxiliar para reemplazar marcadores
Private Function ReemplazarMarcador( _
                                     docWord As Word.Document, _
                                     nombreMarcador As String, _
                                     valor As String _
                                     ) As Boolean

    On Error Resume Next
    
    If docWord.Bookmarks.Exists(nombreMarcador) Then
        docWord.Bookmarks(nombreMarcador).Range.Text = valor
        ' Recrear el marcador para poder usarlo de nuevo
        docWord.Bookmarks.Add nombreMarcador, docWord.Bookmarks(nombreMarcador).Range
        ReemplazarMarcador = True
    Else
        ReemplazarMarcador = False
    End If
    
End Function

Private Function GeneraControlCambiosWord( _
                                           p_ObjEdicion As Edicion, _
                                           docWord As Word.Document, _
                                           Optional ByRef p_Error As String _
                                           ) As String

    Dim tabla As Word.Table
    Dim fila As Word.Row
    Dim colRiesgosActual As Scripting.Dictionary
    Dim colRiesgosAnterior As Scripting.Dictionary
    Dim riesgoActual As riesgo
    Dim riesgoAnterior As riesgo
    Dim objEdicionAnterior As Edicion
    Dim hayCambios As Boolean
    
    On Error GoTo errores
    
    ' Buscar la tabla de control de cambios
    If docWord.Bookmarks.Exists("TablaControlCambios") Then
        Set tabla = docWord.Bookmarks("TablaControlCambios").Range.Tables(1)
    ElseIf docWord.Tables.Count >= 2 Then
        Set tabla = docWord.Tables(2)
    Else
        p_Error = "No se encontró la tabla de control de cambios"
        Err.Raise 1000
    End If
    
    Do While tabla.Rows.Count > 1
        tabla.Rows(2).Delete
    Loop
    
    Dim proy As Proyecto
    Dim ed As Edicion
    Dim edByNum As Scripting.Dictionary
    Dim prevDict As Scripting.Dictionary
    Dim colRiesgos As Scripting.Dictionary
    Dim k As Variant, kEd As Variant
    Dim codigo As String
    Dim minEd As Integer, maxEd As Integer, i As Integer
    
    Set proy = Constructor.getProyecto(p_ObjEdicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set edByNum = New Scripting.Dictionary
    edByNum.CompareMode = TextCompare
    minEd = 32767
    maxEd = -32768
    For Each kEd In proy.colEdiciones
        Set ed = proy.colEdiciones(kEd)
        If IsNumeric(ed.Edicion) Then
            If ed.Edicion < minEd Then minEd = ed.Edicion
            If ed.Edicion > maxEd Then maxEd = ed.Edicion
        End If
        edByNum(CStr(ed.Edicion)) = ed
    Next kEd
    Set prevDict = New Scripting.Dictionary
    prevDict.CompareMode = TextCompare
    For i = minEd To maxEd
        If edByNum.Exists(CStr(i)) Then
            Set ed = edByNum(CStr(i))
            Set colRiesgos = ed.colRiesgos
            If Not colRiesgos Is Nothing Then
                For Each k In colRiesgos
                    Set riesgoActual = colRiesgos(k)
                    codigo = riesgoActual.CodigoRiesgo
                    If Not prevDict.Exists(codigo) Then
                        Set fila = tabla.Rows.Add
                        FormatearFilaDatos fila
                        fila.Cells(1).Range.Text = codigo
                        fila.Cells(2).Range.Text = ed.Edicion
                        fila.Cells(3).Range.Text = ConstruirTextoEstadoCambios(riesgoActual, Nothing, True)
                        fila.Cells(4).Range.Text = ConstruirTextoMitigacionCambios(riesgoActual, Nothing, True)
                        fila.Cells(5).Range.Text = ConstruirTextoContingenciaCambios(riesgoActual, Nothing, True)
                        prevDict.Add codigo, riesgoActual
                    Else
                        Set riesgoAnterior = prevDict(codigo)
                        If SonRiesgosDiferentes(riesgoActual, riesgoAnterior) Then
                            Set fila = tabla.Rows.Add
                            FormatearFilaDatos fila
                            fila.Cells(1).Range.Text = codigo
                            fila.Cells(2).Range.Text = ed.Edicion
                            fila.Cells(3).Range.Text = ConstruirTextoEstadoCambios(riesgoActual, riesgoAnterior, False)
                            fila.Cells(4).Range.Text = ConstruirTextoMitigacionCambios(riesgoActual, riesgoAnterior, False)
                            fila.Cells(5).Range.Text = ConstruirTextoContingenciaCambios(riesgoActual, riesgoAnterior, False)
                        End If
                        Set prevDict(codigo) = riesgoActual
                    End If
                Next k
            End If
        End If
    Next i
    
    GeneraControlCambiosWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "GeneraControlCambiosWord: " & Err.Number & " - " & Err.Description
    End If
End Function
Private Function ConstruirTextoEstadoCambios(r As riesgo, rPrev As riesgo, esPrimera As Boolean) As String
    Dim s As String
    If esPrimera Then
        s = "Detectado por: " & r.DetectadoPor & vbCrLf
        s = s & "Origen: " & r.CausaRaiz & vbCrLf
        s = s & "Impacto global: " & r.ImpactoGlobal & vbCrLf
        s = s & "Vulnerabilidad: " & r.Vulnerabilidad & vbCrLf
        s = s & "Valoración: " & r.Valoracion & vbCrLf
        s = s & "Mitigación: " & r.Mitigacion & vbCrLf
        s = s & "Contingencia: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
        s = s & "Materializado: " & IIf(r.Estado = 3, "Sí", "No") & vbCrLf
        s = s & "Estado: " & r.Estado & vbCrLf
        s = s & "Fecha estado: " & Format(r.FechaEstado, "dd/mm/yyyy") & vbCrLf
        s = s & "Priorización: " & r.Priorizacion
        ConstruirTextoEstadoCambios = s
        Exit Function
    End If
    If rPrev Is Nothing Then
        ConstruirTextoEstadoCambios = ConstruirTextoEstado(r)
        Exit Function
    End If
    If r.DetectadoPor <> rPrev.DetectadoPor Then s = s & "Detectado por: " & r.DetectadoPor & vbCrLf
    If r.CausaRaiz <> rPrev.CausaRaiz Then s = s & "Origen: " & r.CausaRaiz & vbCrLf
    If r.ImpactoGlobal <> rPrev.ImpactoGlobal Then s = s & "Impacto global: " & r.ImpactoGlobal & vbCrLf
    If r.Vulnerabilidad <> rPrev.Vulnerabilidad Then s = s & "Vulnerabilidad: " & r.Vulnerabilidad & vbCrLf
    If r.Valoracion <> rPrev.Valoracion Then s = s & "Valoración: " & r.Valoracion & vbCrLf
    If r.Mitigacion <> rPrev.Mitigacion Then s = s & "Mitigación: " & r.Mitigacion & vbCrLf
    If r.RequierePlanContingencia <> rPrev.RequierePlanContingencia Then s = s & "Contingencia: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
    If r.Estado <> rPrev.Estado Then s = s & "Estado: " & r.Estado & vbCrLf
    If r.FechaEstado <> rPrev.FechaEstado Then s = s & "Fecha estado: " & Format(r.FechaEstado, "dd/mm/yyyy") & vbCrLf
    If r.Priorizacion <> rPrev.Priorizacion Then s = s & "Priorización: " & r.Priorizacion
    ConstruirTextoEstadoCambios = s
End Function

Private Function ConstruirTextoMitigacionCambios(r As riesgo, rPrev As riesgo, esPrimera As Boolean) As String
    Dim s As String
    Dim plan As PM, planPrev As PM
    Dim Accion As PMAccion, accPrev As PMAccion
    Dim k As Variant, kAcc As Variant
    Dim idx As Integer
    If esPrimera Then
        ConstruirTextoMitigacionCambios = ConstruirTextoMitigacion(r)
        Exit Function
    End If
    If rPrev Is Nothing Then
        ConstruirTextoMitigacionCambios = ConstruirTextoMitigacion(r)
        Exit Function
    End If
    If r.ColPMs Is Nothing Then Exit Function
    idx = 1
    For Each k In r.ColPMs
        Set plan = r.ColPMs(k)
        If Not rPrev.ColPMs Is Nothing Then
            On Error Resume Next
            Set planPrev = rPrev.ColPMs.item(k)
            On Error GoTo 0
        End If
        If planPrev Is Nothing Then
            s = s & "ID-PM: " & idx & vbCrLf
            s = s & "PM denominador: " & plan.CodMitigacion & vbCrLf
        Else
            If plan.CodMitigacion <> planPrev.CodMitigacion Then
                s = s & "PM denominador: " & plan.CodMitigacion & vbCrLf
            End If
        End If
        If Not plan.colAcciones Is Nothing Then
            For Each kAcc In plan.colAcciones
                Set Accion = plan.colAcciones(kAcc)
                If Not planPrev Is Nothing And Not planPrev.colAcciones Is Nothing Then
                    On Error Resume Next
                    Set accPrev = planPrev.colAcciones.item(kAcc)
                    On Error GoTo 0
                Else
                    Set accPrev = Nothing
                End If
                If accPrev Is Nothing Then
                    s = s & "Descripción PMA ID: " & Accion.Accion & vbCrLf
                    s = s & "Responsable PMA ID: " & Accion.ResponsableAccion & vbCrLf
                    s = s & "Fecha inicio PMA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                    s = s & "Fecha fin prevista PMA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                    s = s & "Fecha fin real PMA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf
                Else
                    If Accion.Accion <> accPrev.Accion Then s = s & "Descripción PMA ID: " & Accion.Accion & vbCrLf
                    If Accion.ResponsableAccion <> accPrev.ResponsableAccion Then s = s & "Responsable PMA ID: " & Accion.ResponsableAccion & vbCrLf
                    If Accion.FechaInicio <> accPrev.FechaInicio Then s = s & "Fecha inicio PMA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                    If Accion.FechaFinPrevista <> accPrev.FechaFinPrevista Then s = s & "Fecha fin prevista PMA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                    If Accion.FechaFinReal <> accPrev.FechaFinReal Then s = s & "Fecha fin real PMA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf
                End If
            Next kAcc
        End If
        idx = idx + 1
    Next k
    ConstruirTextoMitigacionCambios = s
End Function

Private Function ConstruirTextoContingenciaCambios(r As riesgo, rPrev As riesgo, esPrimera As Boolean) As String
    Dim s As String
    Dim plan As PC, planPrev As PC
    Dim Accion As PCAccion, accPrev As PCAccion
    Dim k As Variant, kAcc As Variant
    If esPrimera Then
        ConstruirTextoContingenciaCambios = ConstruirTextoContingencia(r)
        Exit Function
    End If
    If rPrev Is Nothing Then
        ConstruirTextoContingenciaCambios = ConstruirTextoContingencia(r)
        Exit Function
    End If
    If r.ColPCs Is Nothing Then Exit Function
    For Each k In r.ColPCs
        Set plan = r.ColPCs(k)
        If Not rPrev.ColPCs Is Nothing Then
            On Error Resume Next
            Set planPrev = rPrev.ColPCs.item(k)
            On Error GoTo 0
        End If
        If planPrev Is Nothing Then
            s = s & "Requiere PC: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
            s = s & "PC denominador: " & plan.CodContingencia & vbCrLf
        Else
            If r.RequierePlanContingencia <> rPrev.RequierePlanContingencia Then s = s & "Requiere PC: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
            If plan.CodContingencia <> planPrev.CodContingencia Then s = s & "PC denominador: " & plan.CodContingencia & vbCrLf
        End If
        If Not plan.colAcciones Is Nothing Then
            For Each kAcc In plan.colAcciones
                Set Accion = plan.colAcciones(kAcc)
                If Not planPrev Is Nothing And Not planPrev.colAcciones Is Nothing Then
                    On Error Resume Next
                    Set accPrev = planPrev.colAcciones.item(kAcc)
                    On Error GoTo 0
                Else
                    Set accPrev = Nothing
                End If
                If accPrev Is Nothing Then
                    s = s & "Descripción PCA ID: " & Accion.Accion & vbCrLf
                    s = s & "Responsable PCA ID: " & Accion.ResponsableAccion & vbCrLf
                    s = s & "Fecha inicio PCA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                    s = s & "Fecha fin prevista PCA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                    s = s & "Fecha fin real PCA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf
                Else
                    If Accion.Accion <> accPrev.Accion Then s = s & "Descripción PCA ID: " & Accion.Accion & vbCrLf
                    If Accion.ResponsableAccion <> accPrev.ResponsableAccion Then s = s & "Responsable PCA ID: " & Accion.ResponsableAccion & vbCrLf
                    If Accion.FechaInicio <> accPrev.FechaInicio Then s = s & "Fecha inicio PCA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                    If Accion.FechaFinPrevista <> accPrev.FechaFinPrevista Then s = s & "Fecha fin prevista PCA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                    If Accion.FechaFinReal <> accPrev.FechaFinReal Then s = s & "Fecha fin real PCA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf
                End If
            Next kAcc
        End If
    Next k
    ConstruirTextoContingenciaCambios = s
End Function
Private Function ConstruirTextoEstado(r As riesgo) As String
    Dim s As String
    s = "Detectado por: " & r.DetectadoPor & vbCrLf
    s = s & "Origen: " & r.CausaRaiz & vbCrLf
    s = s & "Impacto global: " & r.ImpactoGlobal & vbCrLf
    s = s & "Vulnerabilidad: " & r.Vulnerabilidad & vbCrLf
    s = s & "Valoración: " & r.Valoracion & vbCrLf
    s = s & "Mitigación: " & r.Mitigacion & vbCrLf
    s = s & "Contingencia: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
    s = s & "Materializado: " & IIf(r.Estado = 3, "Sí", "No") & vbCrLf
    s = s & "Estado: " & r.Estado & vbCrLf
    s = s & "Fecha estado: " & Format(r.FechaEstado, "dd/mm/yyyy") & vbCrLf
    s = s & "Priorización: " & r.Priorizacion
    ConstruirTextoEstado = s
End Function

Private Function ConstruirTextoMitigacion(r As riesgo) As String
    Dim s As String
    Dim plan As PM
    Dim Accion As PMAccion
    Dim k As Variant, kAccion As Variant
    Dim counter As Integer
    counter = 1
    
    If r.ColPMs Is Nothing Then Exit Function
    
    For Each k In r.ColPMs
        Set plan = r.ColPMs(k)
        s = s & "ID-PM: " & counter & vbCrLf
        s = s & "PM denominador: " & plan.CodMitigacion & vbCrLf
        
        If Not plan.colAcciones Is Nothing Then
            For Each kAccion In plan.colAcciones
                Set Accion = plan.colAcciones(kAccion)
                s = s & "Descripción PMA ID: " & Accion.Accion & vbCrLf
                s = s & "Responsable PMA ID: " & Accion.ResponsableAccion & vbCrLf
                s = s & "Fecha inicio PMA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                s = s & "Fecha fin prevista PMA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                s = s & "Fecha fin real PMA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf & vbCrLf
            Next kAccion
        End If
        s = s & "----------------" & vbCrLf
        counter = counter + 1
    Next k
    ConstruirTextoMitigacion = s
End Function

Private Function ConstruirTextoContingencia(r As riesgo) As String
    Dim s As String
    Dim plan As PC
    Dim Accion As PCAccion
    Dim k As Variant, kAccion As Variant
    
    If r.ColPCs Is Nothing Then Exit Function
    
    For Each k In r.ColPCs
        Set plan = r.ColPCs(k)
        s = s & "Requiere PC: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
        s = s & "PC denominador: " & plan.CodContingencia & vbCrLf
        
        If Not plan.colAcciones Is Nothing Then
            For Each kAccion In plan.colAcciones
                Set Accion = plan.colAcciones(kAccion)
                s = s & "Descripción PCA ID: " & Accion.Accion & vbCrLf
                s = s & "Responsable PCA ID: " & Accion.ResponsableAccion & vbCrLf
                s = s & "Fecha inicio PCA ID: " & Format(Accion.FechaInicio, "dd/mm/yyyy") & vbCrLf
                s = s & "Fecha fin prevista PCA ID: " & Format(Accion.FechaFinPrevista, "dd/mm/yyyy") & vbCrLf
                s = s & "Fecha fin real PCA ID: " & Format(Accion.FechaFinReal, "dd/mm/yyyy") & vbCrLf & vbCrLf
            Next kAccion
        End If
        s = s & "----------------" & vbCrLf
    Next k
    ConstruirTextoContingencia = s
End Function

Private Function ObtenerEdicionAnterior(p_EdicionActual As Edicion) As Edicion
    Dim proy As Proyecto
    Dim ed As Edicion
    Dim edAnterior As Edicion
    Dim fechaMax As Date
    
    On Error Resume Next
    Set proy = Constructor.getProyecto(p_EdicionActual.IDProyecto, "")
    
    If proy Is Nothing Then Exit Function
    
    fechaMax = #1/1/1900#
    
    For Each ed In proy.colEdiciones
        If ed.IDEdicion <> p_EdicionActual.IDEdicion Then
            ' Buscar la más reciente anterior a la actual
            ' Asumiendo que FechaEdicion o ID determina el orden. Usaremos FechaEdicion.
            If ed.FechaEdicion < p_EdicionActual.FechaEdicion And ed.FechaEdicion > fechaMax Then
                fechaMax = ed.FechaEdicion
                Set edAnterior = ed
            End If
        End If
    Next
    
    Set ObtenerEdicionAnterior = edAnterior
End Function

Private Function SonRiesgosDiferentes(r1 As riesgo, r2 As riesgo) As Boolean
    ' Comparar campos clave
    If r1.Estado <> r2.Estado Then SonRiesgosDiferentes = True: Exit Function
    If r1.ImpactoGlobal <> r2.ImpactoGlobal Then SonRiesgosDiferentes = True: Exit Function
    If r1.Vulnerabilidad <> r2.Vulnerabilidad Then SonRiesgosDiferentes = True: Exit Function
    If r1.Valoracion <> r2.Valoracion Then SonRiesgosDiferentes = True: Exit Function
    If r1.Priorizacion <> r2.Priorizacion Then SonRiesgosDiferentes = True: Exit Function
    
    ' Comparar planes (simplificado: si cambia número de planes o acciones, o fechas clave)
    If r1.ColPMs.Count <> r2.ColPMs.Count Then SonRiesgosDiferentes = True: Exit Function
    If r1.ColPCs.Count <> r2.ColPCs.Count Then SonRiesgosDiferentes = True: Exit Function
    
    SonRiesgosDiferentes = False
End Function



' Genera las fichas de riesgo
Private Function GeneraFichasRiesgoWord( _
                                        p_ObjEdicion As Edicion, _
                                        p_ObjProyecto As Proyecto, _
                                        docWord As Word.Document, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim tablaPlantilla As Word.Table
    Dim tablaActual As Word.Table
    Dim rng As Word.Range
    Dim i As Integer
    Dim colRiesgos As Scripting.Dictionary
    Dim riesgo As riesgo
    
    On Error GoTo errores
    
    ' 1. Localizar la plantilla de la ficha
    If Not docWord.Bookmarks.Exists("PlantillaFichaRiesgo") Then
        p_Error = "No se encontró el marcador 'PlantillaFichaRiesgo' en la plantilla."
        Err.Raise 1000
    End If
    
    Set tablaPlantilla = docWord.Bookmarks("PlantillaFichaRiesgo").Range.Tables(1)
    
    ' Obtener riesgos
    Set colRiesgos = p_ObjEdicion.colRiesgos
    
    ' 2. Analizar estructura de la plantilla (índices de filas para planes)
    Dim idxTituloMit As Integer, idxDatosMit As Integer
    Dim idxTituloCon As Integer, idxDatosCon As Integer
    
    If Not ObtenerIndicesFilasPlan(docWord, "FilaTituloMitigacion", "FilaAccionMitigacion", idxTituloMit, idxDatosMit) Then
        p_Error = "No se encontraron los marcadores de filas para Mitigación."
        Err.Raise 1000
    End If
    
    If Not ObtenerIndicesFilasPlan(docWord, "FilaTituloContingencia", "FilaAccionContingencia", idxTituloCon, idxDatosCon) Then
        p_Error = "No se encontraron los marcadores de filas para Contingencia."
        Err.Raise 1000
    End If
    
    ' 3. Generar Fichas
    If colRiesgos.Count = 0 Then
        LimpiarTablaFicha tablaPlantilla
        GeneraFichasRiesgoWord = "OK"
        Exit Function
    End If
    
    Dim k As Variant
    Dim counter As Integer
    counter = 1
    
    For Each k In colRiesgos
        Set riesgo = colRiesgos(k)
        
        If counter = 1 Then
            Set tablaActual = tablaPlantilla
        Else
            tablaPlantilla.Range.Copy
            Set rng = docWord.Content
            rng.Collapse 0 ' wdCollapseEnd
            rng.InsertBreak 7 ' wdPageBreak
            rng.Paste
            Set tablaActual = docWord.Tables(docWord.Tables.Count)
        End If
        
        RellenarCabeceraFicha tablaActual, riesgo, p_ObjProyecto
        ProcesarPlanesFichaContingencia tablaActual, riesgo.ColPCs, idxTituloCon, idxDatosCon, docWord
        ProcesarPlanesFichaMitigacion tablaActual, riesgo.ColPMs, idxTituloMit, idxDatosMit, docWord
        
        counter = counter + 1
    Next k
    
    GeneraFichasRiesgoWord = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "GeneraFichasRiesgoWord: " & Err.Number & " - " & Err.Description
    End If
End Function

Private Function ObtenerIndicesFilasPlan(doc As Word.Document, sMarcTitulo As String, sMarcDatos As String, ByRef idxTitulo As Integer, ByRef idxDatos As Integer) As Boolean
    On Error Resume Next
    If doc.Bookmarks.Exists(sMarcTitulo) And doc.Bookmarks.Exists(sMarcDatos) Then
        idxTitulo = doc.Bookmarks(sMarcTitulo).Range.Cells(1).RowIndex
        idxDatos = doc.Bookmarks(sMarcDatos).Range.Cells(1).RowIndex
        ObtenerIndicesFilasPlan = True
    Else
        ObtenerIndicesFilasPlan = False
    End If
End Function

Private Sub RellenarCabeceraFicha(tabla As Word.Table, riesgo As riesgo, Proyecto As Proyecto)
    ReemplazarTextoEnRango tabla.Range, "[NOMBRE_PROYECTO]", Proyecto.NombreProyecto
    ReemplazarTextoEnRango tabla.Range, "[FECHA]", Format(Date, "dd/mm/yyyy")
    
    On Error Resume Next
    Dim fDatos As Integer
    fDatos = 10
    
    tabla.Cell(fDatos, 1).Range.Text = riesgo.CodigoRiesgo
    tabla.Cell(fDatos, 2).Range.Text = riesgo.DetectadoPor
    tabla.Cell(fDatos, 3).Range.Text = riesgo.ImpactoGlobal
    tabla.Cell(fDatos, 4).Range.Text = riesgo.Vulnerabilidad
    tabla.Cell(fDatos, 5).Range.Text = riesgo.Valoracion
    tabla.Cell(fDatos, 6).Range.Text = riesgo.Mitigacion
        tabla.Cell(fDatos, 7).Range.Text = IIf(riesgo.RequierePlanContingencia, "Sí", "No")
    tabla.Cell(fDatos, 8).Range.Text = IIf(riesgo.Estado = 3, Format(riesgo.FechaEstado, "dd/mm/yyyy"), "")
    
    tabla.Cell(12, 1).Range.Text = riesgo.Descripcion
    tabla.Cell(14, 1).Range.Text = riesgo.CausaRaiz
    
    If riesgo.Estado = 3 Then
        tabla.Cell(6, 2).Range.Text = "Sí"
        tabla.Cell(6, 3).Range.Text = Format(riesgo.FechaEstado, "dd/mm/yyyy")
    Else
        tabla.Cell(6, 2).Range.Text = "NO"
        tabla.Cell(6, 3).Range.Text = "-"
    End If
    
    If riesgo.Estado = 4 Then
        tabla.Cell(7, 2).Range.Text = "Sí"
        tabla.Cell(7, 3).Range.Text = Format(riesgo.FechaEstado, "dd/mm/yyyy")
    Else
        tabla.Cell(7, 2).Range.Text = "NO"
        tabla.Cell(7, 3).Range.Text = "-"
    End If
End Sub

Private Sub ReemplazarTextoEnRango(rng As Word.Range, sBusca As String, sReemplaza As String)
    With rng.Find
        .Text = sBusca
        .Replacement.Text = sReemplaza
        .Execute Replace:=2
    End With
End Sub

Private Sub ProcesarPlanesFichaMitigacion(tabla As Word.Table, colPlanes As Scripting.Dictionary, idxTituloBase As Integer, idxDatosBase As Integer, docWord As Word.Document)
    Dim i As Integer
    Dim plan As PM
    Dim Accion As PMAccion
    Dim alturaBloque As Integer
    Dim rngBloque As Word.Range
    Dim nPlanes As Integer
    
    alturaBloque = idxDatosBase - idxTituloBase + 1
    
    If colPlanes Is Nothing Or colPlanes.Count = 0 Then
        tabla.Rows(idxDatosBase).Range.Text = ""
        tabla.Cell(idxTituloBase, 1).Range.Text = "Plan de Mitigación (1)"
        Exit Sub
    End If
    
    nPlanes = colPlanes.Count
    
    ' Replicar bloques si hay más de 1 plan
    If nPlanes > 1 Then
        Set rngBloque = docWord.Range(tabla.Rows(idxTituloBase).Range.Start, tabla.Rows(idxDatosBase).Range.End)
        rngBloque.Copy
        
        For i = 2 To nPlanes
            Dim posPegado As Word.Range
            Set posPegado = tabla.Rows(idxDatosBase + (i - 2) * alturaBloque + 1).Range
            posPegado.Collapse 1
            posPegado.Paste
        Next i
    End If
    
    Dim baseActual As Integer
    baseActual = idxTituloBase

    Dim k As Variant
    Dim counter As Integer
    counter = 1

    For Each k In colPlanes
        Set plan = colPlanes(k)

        tabla.Cell(baseActual, 1).Range.Text = "Plan de Mitigación (" & counter & ")"
        
        Dim filaAccion As Integer
        filaAccion = baseActual + alturaBloque - 1
        
        Dim colAcciones As Scripting.Dictionary
        Set colAcciones = plan.colAcciones
        
        If colAcciones Is Nothing Or colAcciones.Count = 0 Then
            tabla.Rows(filaAccion).Range.Text = ""
        Else
            Dim kAccion As Variant
            Dim firstAccion As Boolean
            firstAccion = True
            
            For Each kAccion In colAcciones
                Set Accion = colAcciones(kAccion)
                
                If Not firstAccion Then
                    tabla.Rows(filaAccion).Range.Copy
                    tabla.Rows(filaAccion + 1).Range.Paste
                    filaAccion = filaAccion + 1
                End If
                firstAccion = False
                
                tabla.Cell(filaAccion, 1).Range.Text = Accion.Accion
                tabla.Cell(filaAccion, 2).Range.Text = Accion.ResponsableAccion
                tabla.Cell(filaAccion, 3).Range.Text = Format(Accion.FechaInicio, "dd/mm/yyyy")
                tabla.Cell(filaAccion, 4).Range.Text = Format(Accion.FechaFinPrevista, "dd/mm/yyyy")
                tabla.Cell(filaAccion, 5).Range.Text = Format(Accion.FechaFinReal, "dd/mm/yyyy")
            Next kAccion
        End If
        
        Dim filasExtra As Integer
        If colAcciones Is Nothing Then filasExtra = 0 Else filasExtra = IIf(colAcciones.Count > 1, colAcciones.Count - 1, 0)
        baseActual = baseActual + alturaBloque + filasExtra
        counter = counter + 1
    Next k
End Sub

Private Sub ProcesarPlanesFichaContingencia(tabla As Word.Table, colPlanes As Scripting.Dictionary, idxTituloBase As Integer, idxDatosBase As Integer, docWord As Word.Document)
    Dim i As Integer
    Dim plan As PC
    Dim Accion As PCAccion
    Dim alturaBloque As Integer
    Dim rngBloque As Word.Range
    Dim nPlanes As Integer

    alturaBloque = idxDatosBase - idxTituloBase + 1

    If colPlanes Is Nothing Or colPlanes.Count = 0 Then
        tabla.Rows(idxDatosBase).Range.Text = ""
        tabla.Cell(idxTituloBase, 1).Range.Text = "Plan de Contingencia (1)"
        Exit Sub
    End If

    nPlanes = colPlanes.Count

    If nPlanes > 1 Then
        Set rngBloque = docWord.Range(tabla.Rows(idxTituloBase).Range.Start, tabla.Rows(idxDatosBase).Range.End)
        rngBloque.Copy

        For i = 2 To nPlanes
            Dim posPegado As Word.Range
            Set posPegado = tabla.Rows(idxDatosBase + (i - 2) * alturaBloque + 1).Range
            posPegado.Collapse 1
            posPegado.Paste
        Next i
    End If

    Dim baseActual As Integer
    baseActual = idxTituloBase

    Dim k As Variant
    Dim counter As Integer
    counter = 1

    For Each k In colPlanes
        Set plan = colPlanes(k)

        tabla.Cell(baseActual, 1).Range.Text = "Plan de Contingencia (" & counter & ")"

        Dim filaAccion As Integer
        filaAccion = baseActual + alturaBloque - 1

        Dim colAcciones As Scripting.Dictionary
        Set colAcciones = plan.colAcciones

        If colAcciones Is Nothing Or colAcciones.Count = 0 Then
            tabla.Rows(filaAccion).Range.Text = ""
        Else
            Dim kAccion As Variant
            Dim firstAccion As Boolean
            firstAccion = True

            For Each kAccion In colAcciones
                Set Accion = colAcciones(kAccion)

                If Not firstAccion Then
                    tabla.Rows(filaAccion).Range.Copy
                    tabla.Rows(filaAccion + 1).Range.Paste
                    filaAccion = filaAccion + 1
                End If
                firstAccion = False

                tabla.Cell(filaAccion, 1).Range.Text = Accion.Accion
                tabla.Cell(filaAccion, 2).Range.Text = Accion.ResponsableAccion
                tabla.Cell(filaAccion, 3).Range.Text = Format(Accion.FechaInicio, "dd/mm/yyyy")
                tabla.Cell(filaAccion, 4).Range.Text = Format(Accion.FechaFinPrevista, "dd/mm/yyyy")
                tabla.Cell(filaAccion, 5).Range.Text = Format(Accion.FechaFinReal, "dd/mm/yyyy")
            Next kAccion
        End If

        Dim filasExtra As Integer
        If colAcciones Is Nothing Then filasExtra = 0 Else filasExtra = IIf(colAcciones.Count > 1, colAcciones.Count - 1, 0)
        baseActual = baseActual + alturaBloque + filasExtra
        counter = counter + 1
    Next k
End Sub

Private Sub LimpiarTablaFicha(tabla As Word.Table)
    tabla.Delete
End Sub



' Aplica formato estándar a una fila de datos para asegurar que no herede el estilo del encabezado
Private Sub FormatearFilaDatos(fila As Word.Row)
    On Error Resume Next
    With fila.Range
        .Font.Bold = False
        .Font.Color = 0 ' wdColorBlack
        .Font.Size = 10 ' Tamaño estándar
        .Shading.Texture = 0 ' wdTextureNone
        .Shading.ForegroundPatternColor = -16777216 ' wdColorAutomatic
        .Shading.BackgroundPatternColor = -16777216 ' wdColorAutomatic (Transparente)
    End With
End Sub

Private Sub LimpiarContenidoFila(f As Word.Row)
    Dim c As Integer
    On Error Resume Next
    For c = 1 To f.Cells.Count
        f.Cells(c).Range.Text = ""
    Next c
End Sub

Private Sub CopiarFormatoFila(fBase As Word.Row, fDest As Word.Row)
    Dim i As Integer
    On Error Resume Next
    With fDest.Range.Font
        .Bold = fBase.Range.Font.Bold
        .Color = fBase.Range.Font.Color
        .Size = fBase.Range.Font.Size
    End With
    With fDest.Range.Shading
        .Texture = fBase.Range.Shading.Texture
        .ForegroundPatternColor = fBase.Range.Shading.ForegroundPatternColor
        .BackgroundPatternColor = fBase.Range.Shading.BackgroundPatternColor
    End With
    For i = 1 To 6
        fDest.Borders(i).LineStyle = fBase.Borders(i).LineStyle
        fDest.Borders(i).LineWidth = fBase.Borders(i).LineWidth
        fDest.Borders(i).Color = fBase.Borders(i).Color
    Next i
End Sub

Private Function BuscarFilaPorTexto(tabla As Word.Table, texto As String) As Integer
    Dim rng As Word.Range
    Set rng = tabla.Range
    With rng.Find
        .ClearFormatting
        .Text = texto
        .MatchCase = False
        .MatchWildcards = True
        .Forward = True
        .Wrap = 1
        .Execute
    End With
    If rng.Find.Found Then
        BuscarFilaPorTexto = rng.Cells(1).RowIndex
    Else
        BuscarFilaPorTexto = 0
    End If
End Function

Private Function GetRowByIndexSafe(tbl As Word.Table, idx As Integer) As Word.Row
    Dim c As Integer
    Dim r As Word.Row
    On Error Resume Next
    For c = 1 To tbl.Columns.Count
        Set r = tbl.Cell(idx, c).Range.Rows(1)
        If Not r Is Nothing Then
            Set GetRowByIndexSafe = r
            Exit Function
        End If
    Next c
    Set GetRowByIndexSafe = Nothing
End Function

Private Sub DeleteRowSafe(tbl As Word.Table, idx As Integer)
    Dim r As Word.Row
    Set r = GetRowByIndexSafe(tbl, idx)
    If Not r Is Nothing Then r.Delete
End Sub

Private Sub AplicarBordeSuperiorFila(fDest As Word.Row, fBase As Word.Row)
    On Error Resume Next
    If fDest Is Nothing Or fBase Is Nothing Then Exit Sub
    With fDest.Borders(1)
        .LineStyle = fBase.Borders(1).LineStyle
        .LineWidth = fBase.Borders(1).LineWidth
        .Color = fBase.Borders(1).Color
    End With
End Sub

Private Sub AplicarEstiloUltimaFila(fDest As Word.Row)
    On Error Resume Next
    If fDest Is Nothing Then Exit Sub
    With fDest.Range.Shading
        .Texture = 0
        .ForegroundPatternColor = -16777216
        .BackgroundPatternColor = RGB(197, 217, 241)
    End With
    With fDest.Borders(3)
        .LineStyle = 1
        .LineWidth = 2
    End With
End Sub

Private Sub AplicarBordesVerticalesFilaDesdeBase(fDest As Word.Row, fBase As Word.Row)
    Dim c As Integer
    Dim maxC As Integer
    On Error Resume Next
    If fDest Is Nothing Or fBase Is Nothing Then Exit Sub
    maxC = fDest.Cells.Count
    If fBase.Cells.Count < maxC Then maxC = fBase.Cells.Count
    For c = 1 To maxC
        With fDest.Cells(c).Borders(2)
            .LineStyle = fBase.Cells(c).Borders(2).LineStyle
            .LineWidth = fBase.Cells(c).Borders(2).LineWidth
            .Color = fBase.Cells(c).Borders(2).Color
        End With
        With fDest.Cells(c).Borders(4)
            .LineStyle = fBase.Cells(c).Borders(4).LineStyle
            .LineWidth = fBase.Cells(c).Borders(4).LineWidth
            .Color = fBase.Cells(c).Borders(4).Color
        End With
    Next c
End Sub

Private Sub AplicarGrillaCompletaFila(fRow As Word.Row)
    On Error Resume Next
    If fRow Is Nothing Then Exit Sub
    With fRow.Range.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With fRow.Range.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With fRow.Range.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With fRow.Range.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With fRow.Range.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With fRow.Range.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
End Sub






