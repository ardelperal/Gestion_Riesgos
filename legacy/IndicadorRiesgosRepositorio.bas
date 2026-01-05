Attribute VB_Name = "IndicadorRiesgosRepositorio"
'===============================
' IndicadorRiesgosRepositorioV2.bas
'===============================
Option Compare Database
Option Explicit

' QueryDef temporal (se recrea cada vez)
Private Const QD_NAME As String = "qtmp_IndicadoresRiesgosV2"

' Repositorio V2: devuelve recordset con 5 columnas de indicadores (Mario)
'   - Identificados (TbRiesgos.FechaDetectado)
'   - Retirados     (TbRiesgos.FechaRetirado)
'   - EnOferta      (TbRiesgosAIntegrar.FechaDetectado)
'   - Materializados(TbRiesgos.FechaMaterializado)
'   - Oferta->Gestion (TbRiesgosAIntegrar.Trasladar='Sí')
Public Function IndicadorRiesgosRepositorioV2_GetTabla( _
                                                        ByVal p_dIni As Date, _
                                                        ByVal p_dFin As Date, _
                                                        ByVal p_ListaIDsCsv As String, _
                                                        Optional ByRef p_Error As String _
                                                    ) As DAO.Recordset

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    Dim sIni As String, sFin As String
    Dim whereProy As String

    On Error GoTo errores
    p_Error = ""

    ' Access requiere literales de fecha en #mm/dd/yyyy#
    sIni = Format$(p_dIni, "mm\/dd\/yyyy")
    sFin = Format$(p_dFin, "mm\/dd\/yyyy")

    whereProy = ""
    If Len(Trim$(p_ListaIDsCsv)) > 0 Then
        whereProy = "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ") "
    End If

    ' NOTA IMPORTANTE:
    ' - NO usamos campos inexistentes (EnOferta/DetectadoEnOferta/PasoAFinal) en TbRiesgos
    ' - La parte "oferta" sale de TbRiesgosAIntegrar (ERD: FechaDetectado, Trasladar)
    m_SQL = ""
    m_SQL = m_SQL & "SELECT " & vbCrLf
    m_SQL = m_SQL & "    P.IDProyecto," & vbCrLf
    m_SQL = m_SQL & "    (P.Proyecto & ' ' & P.NombreProyecto) AS ProyectoCompleto," & vbCrLf

    ' 1) Riesgos identificados (ejecución)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaDetectado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosIdentificados," & vbCrLf

    ' 2) Riesgos retirados (ejecución)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaRetirado Is Not Null" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaRetirado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosRetirados," & vbCrLf

    ' 3) Riesgos en oferta (oferta)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND RA.FechaDetectado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosEnOferta," & vbCrLf

   ' 4) Riesgos materializados (usar TbRiesgosMaterializaciones: contar EVENTOS)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbRiesgosMaterializaciones AS M" & vbCrLf
    m_SQL = m_SQL & "        WHERE M.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND Nz(M.EsMaterializacion,'No')='Sí'" & vbCrLf
    m_SQL = m_SQL & "          AND M.Fecha Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosMaterializados," & vbCrLf



   ' 5) Riesgos de oferta que pasan a gestión (oferta -> ejecución): Trasladar='Sí'
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND Trim(Nz(RA.Trasladar,''))='Sí'" & vbCrLf
    m_SQL = m_SQL & "          AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosOfertaPasanGestion" & vbCrLf


    m_SQL = m_SQL & "FROM TbProyectos AS P" & vbCrLf
    m_SQL = m_SQL & whereProy & vbCrLf
    m_SQL = m_SQL & "ORDER BY P.Proyecto, P.NombreProyecto;"

    Set db = getdb()
    Set rs = db.OpenRecordset(m_SQL, dbOpenSnapshot)

    Set IndicadorRiesgosRepositorioV2_GetTabla = rs
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "IndicadorRiesgosRepositorioV2_GetTabla: " & Err.Number & vbCrLf & Err.Description
    End If
    Set IndicadorRiesgosRepositorioV2_GetTabla = Nothing
End Function


'========================================================
' RESUMEN PARA DASHBOARD (1 fila con 5 números)
'  - Identificados: 1 por CodigoUnico (sin duplicar por ediciones)
'  - Retirados:     1 por CodigoUnico (sin duplicar por ediciones)
'  - En oferta:     1 por IDRiesgoExt (TbRiesgosAIntegrar) (sin duplicar)
'  - Materializados: cuenta TODAS las materializaciones (TbRiesgosMaterializaciones)
'                    SOLO EsMaterializacion='Sí' (pueden repetirse por riesgo)
'  - Oferta->Gestión: 1 por IDRiesgoExt con Trasladar='Sí'
'========================================================
Public Function IndicadorRiesgosV2_GetResumen( _
                                                    ByVal p_dIni As Date, _
                                                    ByVal p_dFin As Date, _
                                                    ByVal p_ListaIDsCsv As String, _
                                                    ByRef p_Error As String _
                                                ) As DAO.Recordset

    Dim db As DAO.Database
    Dim sql As String
    Dim sIni As String, sFin As String

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_ListaIDsCsv & "")) = 0 Then
        p_Error = "Lista de proyectos vacía."
        Err.Raise 1000
    End If

    ' Access requiere literales de fecha en #mm/dd/yyyy#
    sIni = Format$(p_dIni, "mm\/dd\/yyyy")
    sFin = Format$(p_dFin, "mm\/dd\/yyyy")

    sql = ""
    sql = sql & "SELECT TOP 1 " & vbCrLf

    ' 1) Identificados (UNIQ por CodigoUnico)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT R.CodigoUnico" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf
    sql = sql & "        AND R.CodigoUnico Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaDetectado Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY R.CodigoUnico" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosIdentificados," & vbCrLf

    ' 2) Retirados (UNIQ por CodigoUnico)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT R.CodigoUnico" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf
    sql = sql & "        AND R.CodigoUnico Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaRetirado Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaRetirado Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY R.CodigoUnico" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosRetirados," & vbCrLf

    ' 3) En oferta (UNIQ por IDRiesgoExt) -> TbRiesgosAIntegrar (join por IDEdicion)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT RA.IDRiesgoExt" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf
    sql = sql & "        AND RA.IDRiesgoExt Is Not Null" & vbCrLf
    sql = sql & "        AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY RA.IDRiesgoExt" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosEnOferta," & vbCrLf

    ' 4) Materializados (cuenta TODAS las materializaciones) -> TbRiesgosMaterializaciones
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM TbRiesgosMaterializaciones AS M" & vbCrLf
    sql = sql & "    WHERE M.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf
    sql = sql & "      AND Nz(M.EsMaterializacion,'')='Sí'" & vbCrLf
    sql = sql & "      AND M.Fecha Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "  ),0) AS RiesgosMaterializados," & vbCrLf

    ' 5) Oferta -> Gestión final (UNIQ por IDRiesgoExt) con Trasladar='Sí'
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT RA.IDRiesgoExt" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf
    sql = sql & "        AND RA.IDRiesgoExt Is Not Null" & vbCrLf
    sql = sql & "        AND Nz(RA.Trasladar,'')='Sí'" & vbCrLf
    sql = sql & "        AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY RA.IDRiesgoExt" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosOfertaPasanGestion" & vbCrLf

    ' Tabla “ancla” para que Access no proteste (1 fila)
    sql = sql & "FROM TbProyectos AS X;" & vbCrLf

    Set db = getdb()
    Set IndicadorRiesgosV2_GetResumen = db.OpenRecordset(sql, dbOpenSnapshot)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "IndicadorRiesgosV2_GetResumen: " & Err.Number & vbCrLf & Err.Description
    End If
    Set IndicadorRiesgosV2_GetResumen = Nothing
End Function




Public Function IndicadorRiesgosV2_SqlDetalle( _
                                                ByVal p_Tipo As IndicadorTile, _
                                                ByVal p_dIni As Date, _
                                                ByVal p_dFin As Date, _
                                                ByVal p_ListaIDsCsv As String _
                                            ) As String

    Dim dIni As String, dFin As String
    dIni = Format$(p_dIni, "mm\/dd\/yyyy")
    dFin = Format$(p_dFin, "mm\/dd\/yyyy")

    Select Case p_Tipo

        ' =========================================================
        ' IDENTIFICADOS (1 fila por RIESGO - evita duplicar por edición)
        ' Cuenta/Detalle por CodigoUnico (si se copia entre ediciones)
        ' =========================================================
        Case itIdentificados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  (P.Proyecto & ' ' & P.NombreProyecto) AS Proyecto," & vbCrLf & _
                "  R.CodigoRiesgo as Código," & vbCrLf & _
                "  Min(R.FechaDetectado) AS Detectado " & vbCrLf & _
                "FROM (TbProyectos AS P" & vbCrLf & _
                "      INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto)" & vbCrLf & _
                "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf & _
                "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf & _
                "  AND R.FechaDetectado Between #" & dIni & "# And #" & dFin & "#" & vbCrLf & _
                "GROUP BY P.IDProyecto, P.Proyecto, P.NombreProyecto, R.CodigoRiesgo, R.CodigoUnico" & vbCrLf & _
                "ORDER BY NombreProyecto, Min(R.FechaDetectado), R.CodigoRiesgo;"

        ' =========================================================
        ' RETIRADOS (1 fila por RIESGO - evita duplicar por edición)
        ' =========================================================
        Case itRetirados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  (P.Proyecto & ' ' & P.NombreProyecto) AS Proyecto," & vbCrLf & _
                "  R.CodigoRiesgo as Código," & vbCrLf & _
                "  Max(R.FechaRetirado) AS Retirado " & vbCrLf & _
                "FROM (TbProyectos AS P" & vbCrLf & _
                "      INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto)" & vbCrLf & _
                "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf & _
                "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ")" & vbCrLf & _
                "  AND R.FechaRetirado Is Not Null" & vbCrLf & _
                "  AND R.FechaRetirado Between #" & dIni & "# And #" & dFin & "#" & vbCrLf & _
                "GROUP BY P.IDProyecto, P.Proyecto, P.NombreProyecto, R.CodigoRiesgo, R.CodigoUnico" & vbCrLf & _
                "ORDER BY NombreProyecto, Max(R.FechaRetirado), R.CodigoRiesgo;"

       ' =========================================================
        ' EN OFERTA (TbRiesgosAIntegrar + TbRiesgos)
        ' Campos EXACTOS: NombreProyecto, CodigoRiesgo, FechaDetectado, Trasladar, FechaAltaRegistro
        ' Sin duplicar por ediciones: 1 fila por RA.IDRiesgo (la primera en el rango)
        ' =========================================================
        Case itEnOferta
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT P.NombreProyecto as Proyecto, TbRiesgos.CodigoRiesgo as Código, Nz(RA.Trasladar,'') AS Trasladar, RA.FechaAltaRegistro as Alta" & vbCrLf & _
                "FROM ((TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto) " & vbCrLf & _
                "  INNER JOIN TbRiesgosAIntegrar AS RA ON E.IDEdicion = RA.IDEdicion) " & vbCrLf & _
                "  LEFT JOIN TbRiesgos ON RA.IDRiesgo = TbRiesgos.IDRiesgo " & vbCrLf & _
                "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ") " & vbCrLf & _
                "  AND RA.FechaAltaRegistro Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "ORDER BY P.NombreProyecto, RA.FechaAltaRegistro;"





        ' =========================================================
        ' OFERTA -> PASA A GESTIÓN FINAL (Trasladar='Sí')
        ' =========================================================
        Case itOfertaTrasladar
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT P.NombreProyecto as Proyecto, TbRiesgos.CodigoRiesgo as Código,  Nz(RA.Trasladar,'') AS Trasladar, RA.FechaAltaRegistro as Alta " & vbCrLf & _
                "FROM ((TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto) " & vbCrLf & _
                "  INNER JOIN TbRiesgosAIntegrar AS RA ON E.IDEdicion = RA.IDEdicion) " & vbCrLf & _
                "  INNER JOIN TbRiesgos ON RA.IDRiesgo = TbRiesgos.IDRiesgo " & vbCrLf & _
                "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ") " & vbCrLf & _
                "  AND RA.FechaAltaRegistro Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "  AND Nz(RA.Trasladar,'')='Sí' " & vbCrLf & _
                "ORDER BY P.NombreProyecto, RA.FechaAltaRegistro;"


        ' =========================================================
        ' MATERIALIZADOS (EVENTOS): 1 fila por materialización (EsMaterializacion='Sí')
        ' Aquí NO agrupamos: Calidad quiere contar repeticiones.
        ' Incluimos Detectado/Retirado del riesgo EN ESA EDICIÓN (para contexto).
        ' =========================================================
        Case itMaterializados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  P.NombreProyecto as Proyecto, " & vbCrLf & _
                "  M.CodigoRiesgo as Código, " & vbCrLf & _
                "  M.Fecha AS Materialización " & vbCrLf & _
                "FROM (TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbRiesgosMaterializaciones AS M ON P.IDProyecto = M.IDProyecto) " & vbCrLf & _
                "  LEFT JOIN TbRiesgos AS R ON (R.IDEdicion = M.IDEdicion AND R.CodigoRiesgo = M.CodigoRiesgo) " & vbCrLf & _
                "WHERE P.IDProyecto In (" & p_ListaIDsCsv & ") " & vbCrLf & _
                "  AND Nz(M.EsMaterializacion,'No')='Sí' " & vbCrLf & _
                "  AND M.Fecha Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "ORDER BY P.NombreProyecto, M.Fecha, M.CodigoRiesgo;"


        Case Else
            IndicadorRiesgosV2_SqlDetalle = ""

    End Select
End Function


Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    On Error GoTo salir
    Set tdf = CurrentDb.TableDefs(tableName)
    For Each fld In tdf.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next
salir:
End Function

Public Function Oferta_KeyField() As String
    ' Devuelve el nombre de campo “clave riesgo” en TbRiesgosAIntegrar
    If FieldExists("TbRiesgosAIntegrar", "CodigoUnico") Then
        Oferta_KeyField = "CodigoUnico"
    ElseIf FieldExists("TbRiesgosAIntegrar", "CodigoRiesgo") Then
        Oferta_KeyField = "CodigoRiesgo"
    ElseIf FieldExists("TbRiesgosAIntegrar", "Riesgo") Then
        Oferta_KeyField = "Riesgo"
    Else
        ' último recurso: no ideal, pero evita romper
        Oferta_KeyField = "IDEdicion"
    End If
End Function

Public Function Oferta_DateField() As String
    ' Devuelve el campo fecha “de alta/detectado” en TbRiesgosAIntegrar
    If FieldExists("TbRiesgosAIntegrar", "FechaAltaRegistro") Then
        Oferta_DateField = "FechaAltaRegistro"
    ElseIf FieldExists("TbRiesgosAIntegrar", "FechaDetectado") Then
        Oferta_DateField = "FechaDetectado"
    Else
        Oferta_DateField = "FechaAltaRegistro" ' fallback
    End If
End Function

Public Function Riesgo_KeyField() As String
    ' Clave estable en TbRiesgos
    If FieldExists("TbRiesgos", "CodigoUnico") Then
        Riesgo_KeyField = "CodigoUnico"
    Else
        Riesgo_KeyField = "CodigoRiesgo"
    End If
End Function





