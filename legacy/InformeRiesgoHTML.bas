

Option Compare Database
Option Explicit

' =========================================================================
' Módulo: InformeRiesgoHTML
' Proyecto: GESTIÓN DE RIESGOS
' Descripción: Generador de informes 100% autosuficiente con LOGO SVG PROPIO.
'              Diseño corporativo premium basado en Indicadores.html.
' =========================================================================

Public Function GenerarInformeRiesgoHTML( _
                                        Optional ByVal p_IDRiesgo As String, _
                                        Optional ByVal p_hWnd As Long, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_Riesgo As riesgo
    Dim m_HTML As String
    Dim m_URL As String

    On Error GoTo errores
    p_Error = ""

    ' 1. Obtención del objeto Riesgo
    If p_IDRiesgo <> "" Then
        Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    ElseIf Not m_ObjRiesgoActivo Is Nothing Then
        Set m_Riesgo = m_ObjRiesgoActivo
    End If

    If m_Riesgo Is Nothing Then
        p_Error = "No se ha indicado un riesgo para generar el informe"
        Err.Raise 1000
    End If

    ' 2. Construcción del HTML
    m_HTML = ConstruirInformeRiesgoHTML(m_Riesgo, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    ' 3. Guardar en formato UTF-8 Real (Garantiza tildes y eñes)
    m_URL = GuardarInformeHTML_UTF8(m_Riesgo, m_HTML, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    ' 4. Abrir en el navegador predeterminado
    If p_hWnd = 0 Then
        On Error Resume Next
        p_hWnd = Application.hWndAccessApp
        On Error GoTo errores
    End If
    
    Ejecutar p_hWnd, "open", m_URL, "", "", 1

    GenerarInformeRiesgoHTML = m_URL
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en GenerarInformeRiesgoHTML: " & Err.Description
    End If
End Function

Public Function GenerarInformeEdicionHTML( _
                                        ByVal p_Edicion As Edicion, _
                                        Optional ByVal p_hWnd As Long, _
                                        Optional ByVal p_FechaCierre As String, _
                                        Optional ByVal p_FechaPublicacion As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_HTML As String
    Dim m_URL As String

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "Se ha de indicar la edición"
        Err.Raise 1000
    End If

    m_HTML = ConstruirInformeEdicionHTML(p_Edicion, p_FechaCierre, p_FechaPublicacion, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    m_URL = GuardarInformeEdicionHTML_UTF8(p_Edicion, m_HTML, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    If p_hWnd = 0 Then
        On Error Resume Next
        p_hWnd = Application.hWndAccessApp
        On Error GoTo errores
    End If
    Ejecutar p_hWnd, "open", m_URL, "", "", 1

    GenerarInformeEdicionHTML = m_URL
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en GenerarInformeEdicionHTML: " & Err.Description
    End If
End Function

Private Function ConstruirInformeRiesgoHTML(ByVal p_Riesgo As riesgo, ByRef p_Error As String) As String
    Dim html As String
    Dim sNemotecnico As String
    On Error GoTo errores

    On Error Resume Next
    sNemotecnico = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Nemotecnico, "")
    On Error GoTo errores
    
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html lang='es'>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "    <meta charset='UTF-8'>" & vbCrLf
    html = html & "    <title>Gestión de Riesgos - Ficha " & p_Riesgo.CodigoRiesgo & "</title>" & vbCrLf
    html = html & GetEstilosCSS_Corporativos()
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf
    
    ' HEADER CON LOGO SVG
    html = html & "    <header>" & vbCrLf
    html = html & "        <div class='header-container container'>" & vbCrLf
    html = html & "            <div class='logo-container'>" & vbCrLf
    html = html & "                <div class='logo-wrapper'>" & GetLogoSVG() & "</div>" & vbCrLf
    html = html & "                <div class='header-text'>" & vbCrLf
    If Trim$(sNemotecnico) <> "" Then
        html = html & "                    <h1>GESTIÓN DE RIESGOS - " & HTMLSafe(sNemotecnico) & "</h1>" & vbCrLf
    Else
        html = html & "                    <h1>GESTIÓN DE RIESGOS</h1>" & vbCrLf
    End If
    html = html & "                    <p>Detalle Técnico del Escenario</p>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "            <div class='header-info'>" & vbCrLf
    html = html & "                <p><strong>CÓDIGO: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</strong></p>" & vbCrLf
    html = html & "                <p>Fecha: " & Format(Date, "dd/mm/yyyy") & "</p>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </header>" & vbCrLf

    ' Contenedor Principal
    html = html & "    <div class='container main-container'>" & vbCrLf
    
    ' Sistema de Pestañas
    html = html & "        <div class='tab-container'>" & vbCrLf
    html = html & "            <button class='tab-btn active' onclick=""openTab(event, 'general')"">Datos Generales</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'datosProyecto')"">Datos de Proyecto</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'mitigacion')"">Mitigación</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'contingencia')"">Contingencia</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'materializaciones')"">Materializaciones</button>" & vbCrLf
    html = html & "            <button class='tab-btn' onclick=""openTab(event, 'publicabilidad')"">Publicabilidad</button>" & vbCrLf
    html = html & "        </div>" & vbCrLf

    ' Sección General
    html = html & "        <div id='general' class='tab-content active'>" & vbCrLf & ConstruirSeccionGeneral(p_Riesgo) & "</div>" & vbCrLf

    html = html & "        <div id='datosProyecto' class='tab-content'>" & vbCrLf & ConstruirSeccionDatosProyecto(p_Riesgo) & "</div>" & vbCrLf
    
    ' Sección Mitigación
    html = html & "        <div id='mitigacion' class='tab-content'>" & vbCrLf & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Mitigacion) & "</div>" & vbCrLf
    
    ' Sección Contingencia
    html = html & "        <div id='contingencia' class='tab-content'>" & vbCrLf & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Contingencia) & "</div>" & vbCrLf

    html = html & "        <div id='materializaciones' class='tab-content'>" & vbCrLf & ConstruirSeccionMaterializaciones(p_Riesgo) & "</div>" & vbCrLf
    html = html & "        <div id='publicabilidad' class='tab-content'>" & vbCrLf & ConstruirSeccionPublicabilidad(p_Riesgo) & "</div>" & vbCrLf

    html = html & "    </div>" & vbCrLf
    
    html = html & "    <footer><p>© " & Year(Date) & " Telefónica - Aplicación GESTIÓN DE RIESGOS</p></footer>" & vbCrLf
    html = html & GetScriptsJS()
    html = html & "</body>" & vbCrLf
    html = html & "</html>"
    
    ConstruirInformeRiesgoHTML = html
    Exit Function
errores:
    p_Error = "Error en ConstruirInformeRiesgoHTML: " & Err.Description
End Function

Private Function ConstruirInformeEdicionHTML( _
                                            ByVal p_Edicion As Edicion, _
                                            ByVal p_FechaCierre As String, _
                                            ByVal p_FechaPublicacion As String, _
                                            ByRef p_Error As String _
                                            ) As String

    Dim html As String
    Dim m_Proyecto As Proyecto
    Dim sNombreProyecto As String, sExpediente As String, sCliente As String, sCodigoDocumento As String
    Dim sJefeProyecto As String, sEdicion As String, sFechaPub As String, sFechaCierre As String
    Dim sNemotecnico As String

    On Error GoTo errores
    p_Error = ""

    Set m_Proyecto = p_Edicion.Proyecto
    If m_Proyecto Is Nothing Then
        Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    End If

    sNombreProyecto = Nz(m_Proyecto.NombreProyecto, Nz(m_Proyecto.Proyecto, ""))
    sExpediente = Nz(m_Proyecto.Proyecto, "")
    sCliente = Nz(m_Proyecto.Cliente, "")
    sCodigoDocumento = Nz(m_Proyecto.CodigoDocumento, "")
    sJefeProyecto = Nz(p_Edicion.Elaborado, "")
    sEdicion = Nz(p_Edicion.Edicion, "")
    sNemotecnico = ""
    On Error Resume Next
    sNemotecnico = Nz(m_Proyecto.Expediente.Nemotecnico, "")
    On Error GoTo errores

    If Not IsDate(p_FechaPublicacion) Then
        sFechaPub = Format$(Date, "dd/mm/yyyy")
    Else
        sFechaPub = Format$(CDate(p_FechaPublicacion), "dd/mm/yyyy")
    End If

    If Not IsDate(p_FechaCierre) Then
        sFechaCierre = ""
    Else
        sFechaCierre = Format$(CDate(p_FechaCierre), "dd/mm/yyyy")
    End If

    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html lang='es'>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "    <meta charset='UTF-8'>" & vbCrLf
    html = html & "    <title>Gestión de Riesgos - Informe Edición " & HTMLSafe(sEdicion) & "</title>" & vbCrLf
    html = html & GetEstilosCSS_Corporativos()
    html = html & GetEstilosCSS_InformeEdicion_Print()
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf

    html = html & "    <header>" & vbCrLf
    html = html & "        <div class='header-container container'>" & vbCrLf
    html = html & "            <div class='logo-container'>" & vbCrLf
    html = html & "                <div class='logo-wrapper'>" & GetLogoSVG() & "</div>" & vbCrLf
    html = html & "                <div class='header-text'>" & vbCrLf
    If Trim$(sNemotecnico) <> "" Then
        html = html & "                    <h1>GESTIÓN DE RIESGOS - " & HTMLSafe(sNemotecnico) & "</h1>" & vbCrLf
    Else
        html = html & "                    <h1>GESTIÓN DE RIESGOS</h1>" & vbCrLf
    End If
    html = html & "                    <p>Informe de edición</p>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "            <div class='header-info'>" & vbCrLf
    html = html & "                <p><strong>EDICIÓN: " & HTMLSafe(sEdicion) & "</strong></p>" & vbCrLf
    html = html & "                <p>Fecha: " & HTMLSafe(sFechaPub) & "</p>" & vbCrLf
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </header>" & vbCrLf

    html = html & "    <div class='container main-container'>" & vbCrLf
    Avance "Portada ..."
    html = html & ConstruirSeccionPortadaEdicionHTML(sNombreProyecto, sExpediente, sJefeProyecto, sCliente, sCodigoDocumento, sEdicion, sFechaPub)
    Avance "Cuadro de control ..."
    html = html & ConstruirSeccionCuadroControlHTML(p_Edicion, m_Proyecto)
    Avance "Control Cambios ..."
    html = html & ConstruirSeccionControlCambiosHTML(p_Edicion, m_Proyecto, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Avance "Inventario ..."
    html = html & ConstruirSeccionInventarioRiesgosHTML(p_Edicion, sFechaCierre, sFechaPub, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Avance "Riesgos ..."
    html = html & ConstruirSeccionFichasRiesgoHTML(p_Edicion, m_Proyecto, sFechaCierre, sFechaPub, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    html = html & "    </div>" & vbCrLf
    Avance "footer ..."
    html = html & ConstruirPiePaginaImpresionHTML(sCodigoDocumento, sFechaPub, sEdicion)
    html = html & "    <footer><p>© " & Year(Date) & " Telefónica - Aplicación GESTIÓN DE RIESGOS</p></footer>" & vbCrLf
    html = html & "</body>" & vbCrLf
    html = html & "</html>"

    ConstruirInformeEdicionHTML = html
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ConstruirInformeEdicionHTML: " & Err.Description
    End If
End Function

Private Function ConstruirSeccionPortadaEdicionHTML( _
                                                    ByVal p_NombreProyecto As String, _
                                                    ByVal p_Expediente As String, _
                                                    ByVal p_JefeProyecto As String, _
                                                    ByVal p_Cliente As String, _
                                                    ByVal p_CodigoDocumento As String, _
                                                    ByVal p_Edicion As String, _
                                                    ByVal p_FechaPub As String _
                                                    ) As String

    Dim s As String
    s = "<section class='report-section print-page'>" & vbCrLf
    s = s & "  <div class='report-cover'>" & vbCrLf
    s = s & "    <div class='report-cover-title'>INFORME DE GESTIÓN DE RIESGOS</div>" & vbCrLf
    s = s & "    <div class='report-cover-kv'>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Proyecto</div><div class='kv-v'>" & HTMLSafe(UCase$(p_Expediente)) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Nombre proyecto</div><div class='kv-v'>" & HTMLSafe(UCase$(p_NombreProyecto)) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Jefe del proyecto</div><div class='kv-v'>" & HTMLSafe(p_JefeProyecto) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Cliente</div><div class='kv-v'>" & HTMLSafe(p_Cliente) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Código documento</div><div class='kv-v'>" & HTMLSafe(p_CodigoDocumento) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Edición</div><div class='kv-v'>" & HTMLSafe(p_Edicion) & "</div></div>" & vbCrLf
    s = s & "      <div class='kv'><div class='kv-k'>Fecha publicación</div><div class='kv-v'>" & HTMLSafe(p_FechaPub) & "</div></div>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf
    ConstruirSeccionPortadaEdicionHTML = s
End Function

Private Function ConstruirSeccionCuadroControlHTML( _
                                                    ByVal p_EdicionActual As Edicion, _
                                                    ByVal p_Proyecto As Proyecto _
                                                    ) As String
    Dim s As String
    Dim edByNum As Scripting.Dictionary
    Dim minEd As Long, maxEd As Long, i As Long
    Dim kEd As Variant
    Dim ed As Edicion
    Dim edNum As Long
    Dim sFecha As String
    Dim sElab As String, sRev As String, sApr As String
    
    edNum = 0
    If Not p_EdicionActual Is Nothing Then
        If IsNumeric(p_EdicionActual.Edicion) Then edNum = CLng(p_EdicionActual.Edicion)
    End If
    If edNum <= 0 Then edNum = 0
    
    Set edByNum = New Scripting.Dictionary
    edByNum.CompareMode = TextCompare
    minEd = 2147483647
    maxEd = -2147483647
    
    If Not p_Proyecto Is Nothing Then
        For Each kEd In p_Proyecto.colEdiciones
            Set ed = p_Proyecto.colEdiciones(kEd)
            If Not ed Is Nothing Then
                If IsNumeric(ed.Edicion) Then
                    If CLng(ed.Edicion) <= edNum Or edNum = 0 Then
                        If CLng(ed.Edicion) < minEd Then minEd = CLng(ed.Edicion)
                        If CLng(ed.Edicion) > maxEd Then maxEd = CLng(ed.Edicion)
                        edByNum.Add CStr(ed.Edicion), ed
                    End If
                End If
            End If
        Next kEd
    End If
    
    s = "<section class='report-section print-page'>" & vbCrLf
    s = s & "  <h2>Cuadro de control</h2>" & vbCrLf
    s = s & "  <div class='card'>" & vbCrLf
    s = s & "    <table class='report-table report-table-small'>" & vbCrLf
    s = s & "      <thead><tr><th>Edición</th><th>Fecha</th><th>Elaborado</th><th>Revisado</th><th>Aprobado</th></tr></thead>" & vbCrLf
    s = s & "      <tbody>" & vbCrLf
    
    If minEd = 2147483647 Or maxEd = -2147483647 Then
        s = s & "        <tr><td colspan='5'>No hay ediciones.</td></tr>" & vbCrLf
    Else
        For i = minEd To maxEd
            If edByNum.Exists(CStr(i)) Then
                Set ed = edByNum(CStr(i))
                
                sFecha = ""
                If IsDate(ed.FechaPublicacion) Then
                    sFecha = Format$(CDate(ed.FechaPublicacion), "dd/mm/yyyy")
                ElseIf IsDate(ed.FechaEdicion) Then
                    sFecha = Format$(CDate(ed.FechaEdicion), "dd/mm/yyyy")
                End If
                
                sElab = Nz(ed.Elaborado, "")
                sRev = Nz(ed.Revisado, "")
                sApr = Nz(ed.Aprobado, "")
                
                s = s & "        <tr>" & _
                        "<td>" & HTMLSafe(Nz(ed.Edicion, "")) & "</td>" & _
                        "<td>" & HTMLSafe(sFecha) & "</td>" & _
                        "<td>" & HTMLSafe(sElab) & "</td>" & _
                        "<td>" & HTMLSafe(sRev) & "</td>" & _
                        "<td>" & HTMLSafe(sApr) & "</td>" & _
                        "</tr>" & vbCrLf
            End If
        Next i
    End If
    
    s = s & "      </tbody>" & vbCrLf
    s = s & "    </table>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf
    
    ConstruirSeccionCuadroControlHTML = s
End Function

Private Function ConstruirPiePaginaImpresionHTML( _
                                                ByVal p_CodigoDocumento As String, _
                                                ByVal p_FechaPublicacion As String, _
                                                ByVal p_Edicion As String _
                                                ) As String
    Dim s As String
    s = "<div class='print-footer'>" & vbCrLf
    s = s & "  <div class='print-footer-left'>" & vbCrLf
    s = s & "    <div><strong>Código:</strong> " & HTMLSafe(p_CodigoDocumento) & "</div>" & vbCrLf
    s = s & "    <div><strong>Fecha:</strong> " & HTMLSafe(p_FechaPublicacion) & "</div>" & vbCrLf
    s = s & "    <div><strong>Edición:</strong> " & HTMLSafe(p_Edicion) & "</div>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "  <div class='print-footer-right'>Página <span class='page-number'></span></div>" & vbCrLf
    s = s & "</div>" & vbCrLf
    ConstruirPiePaginaImpresionHTML = s
End Function

Private Function ConstruirSeccionControlCambiosHTML( _
                                                    ByVal p_EdicionActual As Edicion, _
                                                    ByVal p_Proyecto As Proyecto, _
                                                    ByRef p_Error As String _
                                                    ) As String
    On Error GoTo errores
    p_Error = ""

    ConstruirSeccionControlCambiosHTML = ControlCambios_ConstruirSeccionControlCambiosHTML(p_EdicionActual, p_Proyecto, p_Error)
    Exit Function

errores:
    If Err.Number <> 1000 Then p_Error = "Error en ConstruirSeccionControlCambiosHTML: " & Err.Description
End Function

Private Function SonRiesgosDiferentes_HTML(ByVal r1 As riesgo, ByVal r2 As riesgo) As Boolean
    On Error Resume Next
    If r2 Is Nothing Then
        SonRiesgosDiferentes_HTML = True
        Exit Function
    End If
    If r1.Estado <> r2.Estado Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.ImpactoGlobal <> r2.ImpactoGlobal Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Vulnerabilidad <> r2.Vulnerabilidad Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Valoracion <> r2.Valoracion Then SonRiesgosDiferentes_HTML = True: Exit Function
    If r1.Priorizacion <> r2.Priorizacion Then SonRiesgosDiferentes_HTML = True: Exit Function
    If (r1.ColPMs Is Nothing) Xor (r2.ColPMs Is Nothing) Then SonRiesgosDiferentes_HTML = True: Exit Function
    If Not r1.ColPMs Is Nothing And Not r2.ColPMs Is Nothing Then
        If r1.ColPMs.Count <> r2.ColPMs.Count Then SonRiesgosDiferentes_HTML = True: Exit Function
    End If
    If (r1.ColPCs Is Nothing) Xor (r2.ColPCs Is Nothing) Then SonRiesgosDiferentes_HTML = True: Exit Function
    If Not r1.ColPCs Is Nothing And Not r2.ColPCs Is Nothing Then
        If r1.ColPCs.Count <> r2.ColPCs.Count Then SonRiesgosDiferentes_HTML = True: Exit Function
    End If
    SonRiesgosDiferentes_HTML = False
End Function

Private Function ConstruirTextoEstadoCambiosHTML(ByVal r As riesgo, ByVal rPrev As riesgo, ByVal esPrimera As Boolean) As String
    Dim s As String
    If esPrimera Or rPrev Is Nothing Then
        s = ""
        s = s & "Detectado por: " & Nz(r.DetectadoPor, "") & vbCrLf
        s = s & "Origen: " & Nz(r.CausaRaiz, "") & vbCrLf
        s = s & "Impacto global: " & Nz(r.ImpactoGlobal, "") & vbCrLf
        s = s & "Vulnerabilidad: " & Nz(r.Vulnerabilidad, "") & vbCrLf
        s = s & "Valoración: " & Nz(r.Valoracion, "") & vbCrLf
        s = s & "Mitigación: " & Nz(r.Mitigacion, "") & vbCrLf
        s = s & "Contingencia: " & r.RequierePlanContingencia & vbCrLf
        s = s & "Materializado: " & IIf(IsDate(r.FechaMaterializado), "Sí", "No") & vbCrLf
        s = s & "Estado: " & Nz(r.Estado, "") & vbCrLf
        s = s & "Fecha estado: " & FormatoFecha(r.FechaEstado) & vbCrLf
        s = s & "Priorización: " & Nz(r.Priorizacion, "")
        ConstruirTextoEstadoCambiosHTML = HTMLSafeLargo(s)
        Exit Function
    End If

    s = ""
    If Nz(r.DetectadoPor, "") <> Nz(rPrev.DetectadoPor, "") Then s = s & "Detectado por: " & Nz(r.DetectadoPor, "") & vbCrLf
    If Nz(r.CausaRaiz, "") <> Nz(rPrev.CausaRaiz, "") Then s = s & "Origen: " & Nz(r.CausaRaiz, "") & vbCrLf
    If Nz(r.ImpactoGlobal, "") <> Nz(rPrev.ImpactoGlobal, "") Then s = s & "Impacto global: " & Nz(r.ImpactoGlobal, "") & vbCrLf
    If Nz(r.Vulnerabilidad, "") <> Nz(rPrev.Vulnerabilidad, "") Then s = s & "Vulnerabilidad: " & Nz(r.Vulnerabilidad, "") & vbCrLf
    If Nz(r.Valoracion, "") <> Nz(rPrev.Valoracion, "") Then s = s & "Valoración: " & Nz(r.Valoracion, "") & vbCrLf
    If Nz(r.Mitigacion, "") <> Nz(rPrev.Mitigacion, "") Then s = s & "Mitigación: " & Nz(r.Mitigacion, "") & vbCrLf
    If r.RequierePlanContingencia <> rPrev.RequierePlanContingencia Then s = s & "Contingencia: " & IIf(r.RequierePlanContingencia, "Sí", "No") & vbCrLf
    If Nz(r.Estado, "") <> Nz(rPrev.Estado, "") Then s = s & "Estado: " & Nz(r.Estado, "") & vbCrLf
    If Nz(r.FechaEstado, "") <> Nz(rPrev.FechaEstado, "") Then s = s & "Fecha estado: " & FormatoFecha(r.FechaEstado) & vbCrLf
    If Nz(r.Priorizacion, "") <> Nz(rPrev.Priorizacion, "") Then s = s & "Priorización: " & Nz(r.Priorizacion, "")
    ConstruirTextoEstadoCambiosHTML = HTMLSafeLargo(s)
End Function

Private Function ConstruirTextoPlanesCambiosHTML(ByVal colActual As Scripting.Dictionary, ByVal colPrev As Scripting.Dictionary) As String
    Dim sAct As String, sPrev As String
    sAct = SerializarPlanesHTML(colActual)
    sPrev = SerializarPlanesHTML(colPrev)
    If sAct = sPrev Then
        ConstruirTextoPlanesCambiosHTML = ""
    Else
        ConstruirTextoPlanesCambiosHTML = sAct
    End If
End Function

Private Function SerializarPlanesHTML(ByVal colPlanes As Scripting.Dictionary) As String
    Dim s As String
    Dim k As Variant, kAcc As Variant
    Dim plan As Object, Accion As Object

    On Error Resume Next
    If colPlanes Is Nothing Then
        SerializarPlanesHTML = ""
        Exit Function
    End If
    If colPlanes.Count = 0 Then
        SerializarPlanesHTML = ""
        Exit Function
    End If

    For Each k In colPlanes
        Set plan = colPlanes(k)
        If plan Is Nothing Then GoTo siguientePlan

        s = s & Nz(plan.ESTADOCalculadoTexto, "") & ": " & Nz(plan.DisparadorDelPlan, "") & vbCrLf

        If Not plan.colAcciones Is Nothing Then
            For Each kAcc In plan.colAcciones
                Set Accion = plan.colAcciones(kAcc)
                If Not Accion Is Nothing Then
                    s = s & "- " & Nz(Accion.Accion, "") & " (" & Nz(Accion.ResponsableAccion, "") & ") " & _
                        FormatoFecha(Accion.FechaInicio) & " / " & FormatoFecha(Accion.FechaFinPrevista) & " / " & FormatoFecha(Accion.FechaFinReal) & vbCrLf
                End If
            Next kAcc
        End If

siguientePlan:
        Set plan = Nothing
    Next k

    SerializarPlanesHTML = HTMLSafeLargo(Trim$(s))
End Function

Private Function ConstruirSeccionInventarioRiesgosHTML( _
                                                        ByVal p_Edicion As Edicion, _
                                                        ByVal p_FechaCierre As String, _
                                                        ByVal p_FechaPublicacion As String, _
                                                        ByRef p_Error As String _
                                                        ) As String

    Dim s As String
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As riesgo
    Dim p As Proyecto
    Dim incluirCausaRaiz As Boolean
    Dim colSpanTotal As Long
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EstadoTexto As String
    Dim m_FechaEstado As Variant
    Dim iRow As Long

    On Error GoTo errores
    p_Error = ""

    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos
    
    incluirCausaRaiz = False
    On Error Resume Next
    Set p = p_Edicion.Proyecto
    If Not p Is Nothing Then
        incluirCausaRaiz = (p.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí)
    End If
    On Error GoTo errores
    
    If incluirCausaRaiz Then
        colSpanTotal = 13
    Else
        colSpanTotal = 12
    End If

    s = "<section class='report-section print-page'>" & vbCrLf
    s = s & "  <h2>Inventario de Riesgos Detectados</h2>" & vbCrLf
    s = s & "  <div class='card'>" & vbCrLf
    s = s & "    <div class='meta-row'><div><strong>Fecha cierre:</strong> " & HTMLSafe(p_FechaCierre) & "</div><div><strong>Fecha publicación:</strong> " & HTMLSafe(p_FechaPublicacion) & "</div></div>" & vbCrLf
    s = s & "    <div class='table-scroll'>" & vbCrLf
    s = s & "      <table class='report-table report-table-small'>" & vbCrLf
    s = s & "        <thead>" & vbCrLf
    s = s & "          <tr style='background-color: #D6EAF8; color: #333;'>" & vbCrLf
    If incluirCausaRaiz Then
        s = s & _
                "<th colspan='4' style='text-align:center; border-right: 1px solid #ccc;'>Identificación</th>" & _
                "<th colspan='7' style='text-align:center; border-right: 1px solid #ccc;'>Análisis</th>" & _
                "<th colspan='2' style='text-align:center;'>Estado</th>" & _
                "</tr>" & vbCrLf
        s = s & "          <tr style='background-color: #ECF0F1; color: #333;'>" & vbCrLf
        s = s & _
                "<th>Código Riesgo</th>" & _
                "<th>Descripción</th>" & _
                "<th>Causa Raíz</th>" & _
                "<th>Detectado Por</th>" & _
                "<th>Plazo</th><th>Coste</th><th>Calidad</th><th>Global</th>" & _
                "<th>Vulnerabilidad</th>" & _
                "<th>Valoración</th>" & _
                "<th>Priorización</th>" & _
                "<th>Estado</th>" & _
                "<th>Fecha</th>" & _
                "</tr>" & vbCrLf
    Else
        s = s & _
                "<th colspan='3' style='text-align:center; border-right: 1px solid #ccc;'>Identificación</th>" & _
                "<th colspan='7' style='text-align:center; border-right: 1px solid #ccc;'>Análisis</th>" & _
                "<th colspan='2' style='text-align:center;'>Estado</th>" & _
                "</tr>" & vbCrLf
        s = s & "          <tr style='background-color: #ECF0F1; color: #333;'>" & vbCrLf
        s = s & _
                "<th>Código Riesgo</th>" & _
                "<th>Descripción</th>" & _
                "<th>Detectado Por</th>" & _
                "<th>Plazo</th><th>Coste</th><th>Calidad</th><th>Global</th>" & _
                "<th>Vulnerabilidad</th>" & _
                "<th>Valoración</th>" & _
                "<th>Priorización</th>" & _
                "<th>Estado</th>" & _
                "<th>Fecha</th>" & _
                "</tr>" & vbCrLf
    End If
    s = s & "        </thead><tbody>" & vbCrLf

    iRow = 0
    If col Is Nothing Or col.Count = 0 Then
        s = s & "          <tr><td colspan='" & CStr(colSpanTotal) & "'>La edición no tiene riesgos.</td></tr>" & vbCrLf
    Else
        For Each k In col
            Set r = col(k)
            If r Is Nothing Then GoTo siguienteInv
            
            iRow = iRow + 1
            Dim sRowStyle As String
            If iRow Mod 2 = 0 Then
                sRowStyle = " style='background-color: #F4F6F7;'"
            Else
                sRowStyle = " style='background-color: #FFFFFF;'"
            End If
            
            m_Estado = r.EstadoEnum
            m_EstadoTexto = ""
            m_FechaEstado = ""
            
            If m_Estado = EnumRiesgoEstado.Retirado Then
                m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                m_FechaEstado = r.FechaRetirado
            ElseIf m_Estado = EnumRiesgoEstado.Aceptado Then
                m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                m_FechaEstado = r.FechaMitigacionAceptar
            ElseIf m_Estado = EnumRiesgoEstado.Detectado Then
                If Nz(p_FechaCierre, "") <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = r.FechaDetectado
                End If
            ElseIf m_Estado = EnumRiesgoEstado.Materializado Then
                If Nz(p_FechaCierre, "") <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = "Materializado"
                    m_FechaEstado = r.FechaMaterializado
                End If
            Else
                If Nz(p_FechaCierre, "") <> "" Then
                    m_EstadoTexto = "Cerrado"
                    m_FechaEstado = p_FechaCierre
                Else
                    m_EstadoTexto = getEstadoRiesgoTexto(m_Estado)
                    m_FechaEstado = r.FechaEstado
                End If
            End If
            
            If Not IsDate(m_FechaEstado) Then m_FechaEstado = ""

            If incluirCausaRaiz Then
            s = s & "          <tr" & sRowStyle & ">" & _
                    "<td>" & HTMLSafe(Nz(r.CodigoRiesgo, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Descripcion, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.CausaRaiz, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.DetectadoPor, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Plazo, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Coste, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Calidad, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.ImpactoGlobal, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Vulnerabilidad, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Valoracion, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(r.Priorizacion, "")) & "</td>" & _
                    "<td>" & HTMLSafe(Nz(m_EstadoTexto, "")) & "</td>" & _
                    "<td>" & HTMLSafe(FormatoFecha(m_FechaEstado)) & "</td>" & _
                    "</tr>" & vbCrLf
            Else
                s = s & "          <tr" & sRowStyle & ">" & _
                        "<td>" & HTMLSafe(Nz(r.CodigoRiesgo, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Descripcion, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.DetectadoPor, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Plazo, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Coste, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Calidad, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.ImpactoGlobal, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Vulnerabilidad, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Valoracion, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(r.Priorizacion, "")) & "</td>" & _
                        "<td>" & HTMLSafe(Nz(m_EstadoTexto, "")) & "</td>" & _
                        "<td>" & HTMLSafe(FormatoFecha(m_FechaEstado)) & "</td>" & _
                        "</tr>" & vbCrLf
            End If
siguienteInv:
            Set r = Nothing
        Next k
    End If

    s = s & "        </tbody></table>" & vbCrLf
    s = s & "    </div>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf

    ConstruirSeccionInventarioRiesgosHTML = s
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ConstruirSeccionInventarioRiesgosHTML: " & Err.Description
    End If
End Function

Private Function ConstruirSeccionFichasRiesgoHTML( _
                                                ByVal p_Edicion As Edicion, _
                                                ByVal p_Proyecto As Proyecto, _
                                                ByVal p_FechaCierre As String, _
                                                ByVal p_FechaPublicacion As String, _
                                                ByRef p_Error As String _
                                                ) As String

    Dim s As String
    Dim col As Scripting.Dictionary
    Dim k As Variant
    Dim r As riesgo

    On Error GoTo errores
    p_Error = ""

    Set col = p_Edicion.ColRiesgosPorPrioridadTodos
    If col Is Nothing Then Set col = p_Edicion.colRiesgos

    If col Is Nothing Or col.Count = 0 Then
        ConstruirSeccionFichasRiesgoHTML = ""
        Exit Function
    End If

    For Each k In col
        Set r = col(k)
        If r Is Nothing Then GoTo siguienteFicha
        Avance "Riesgos " & r.CodigoRiesgo & " ..."
        s = s & ConstruirFichaRiesgoHTML(r, p_Proyecto, p_FechaCierre, p_FechaPublicacion)
siguienteFicha:
        Set r = Nothing
    Next k

    ConstruirSeccionFichasRiesgoHTML = s
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ConstruirSeccionFichasRiesgoHTML: " & Err.Description
    End If
End Function

Private Function ConstruirFichaRiesgoHTML( _
                                        ByVal p_Riesgo As riesgo, _
                                        ByVal p_Proyecto As Proyecto, _
                                        ByVal p_FechaCierre As String, _
                                        ByVal p_FechaPublicacion As String _
                                        ) As String
    Dim s As String
    Dim sError As String

    s = "<section class='report-section print-page'>" & vbCrLf
    s = s & "  <h2>Ficha de riesgo: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</h2>" & vbCrLf
    s = s & "  <div class='card'>" & vbCrLf
    s = s & "    <div class='meta-row'><div><strong>Proyecto:</strong> " & HTMLSafe(Nz(p_Proyecto.NombreProyecto, Nz(p_Proyecto.Proyecto, ""))) & "</div><div><strong>Fecha:</strong> " & HTMLSafe(Format$(Date, "dd/mm/yyyy")) & "</div></div>" & vbCrLf
    
    ' 1. Histórico de Estados (Moved up)
    s = s & ConstruirTablaEstadosHistoricosHTML(p_Riesgo, p_FechaCierre, p_FechaPublicacion, sError)

    ' 2. Datos del Riesgo (Tabla Horizontal Centrada)
    s = s & "    <div class='card' style='margin-top:18px; border: 1px solid var(--grey-2); text-align: center;'>" & vbCrLf
    s = s & "      <h3 style='margin-top:0; text-align: center; background-color: #D6EAF8; padding: 5px; border-radius: 4px;'>Datos del Riesgo</h3>" & vbCrLf
    s = s & "      <table class='report-table' style='width:100%; text-align:center; margin: 0 auto;'>" & vbCrLf
    s = s & "        <thead>" & vbCrLf
    s = s & "          <tr style='background-color: #ECF0F1;'>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Código</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Detectado por</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Origen</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Impacto</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Vuln.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Valor.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Mitig.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Conting.</th>" & vbCrLf
    s = s & "            <th style='text-align:center;'>Mat.</th>" & vbCrLf
    s = s & "          </tr>" & vbCrLf
    s = s & "        </thead>" & vbCrLf
    s = s & "        <tbody>" & vbCrLf
    s = s & "          <tr>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.CodigoRiesgo, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.DetectadoPor, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Origen, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.ImpactoGlobal, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Vulnerabilidad, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Valoracion, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.Mitigacion, "")) & "</td>" & vbCrLf
    s = s & "            <td>" & HTMLSafe(Nz(p_Riesgo.ContingenciaCalculada, "")) & "</td>" & vbCrLf
    
    Dim sFechaMat As String, sEstiloMat As String
    sFechaMat = ""
    sEstiloMat = ""
    If IsDate(p_Riesgo.FechaMaterializado) Then
        sFechaMat = Format(p_Riesgo.FechaMaterializado, "dd/mm/yyyy")
        sEstiloMat = " style='color:red; font-weight:bold;'"
    End If
    s = s & "            <td" & sEstiloMat & ">" & HTMLSafe(sFechaMat) & "</td>" & vbCrLf
    s = s & "          </tr>" & vbCrLf
    s = s & "        </tbody>" & vbCrLf
    s = s & "      </table>" & vbCrLf
    
    s = s & "    </div>" & vbCrLf

    s = s & "    <div class='card' style='margin-top:18px; border: 1px solid var(--grey-2);'>" & vbCrLf
    s = s & "      <h3 style='margin-top:0;'>Descripción</h3>" & vbCrLf
    s = s & "      <div>" & HTMLSafeLargo(Nz(p_Riesgo.Descripcion, "")) & "</div>" & vbCrLf
    
    If Nz(p_Riesgo.CausaRaiz, "") <> "" Then
        s = s & "      <h3 style='margin-top:16px;'>Causa raíz</h3>" & vbCrLf
        s = s & "      <div>" & HTMLSafeLargo(Nz(p_Riesgo.CausaRaiz, "")) & "</div>" & vbCrLf
    End If
    
    s = s & "    </div>" & vbCrLf

    s = s & "    <div class='card' style='margin-top:18px; border: 1px solid var(--grey-2);'>" & vbCrLf
    s = s & "      " & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Mitigacion) & vbCrLf
    s = s & "    </div>" & vbCrLf

    s = s & "    <div class='card' style='margin-top:18px; border: 1px solid var(--grey-2);'>" & vbCrLf
    s = s & "      " & ConstruirSeccionPlanes(p_Riesgo, EnumTipoPlan.Contingencia) & vbCrLf
    s = s & "    </div>" & vbCrLf

    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf

    ConstruirFichaRiesgoHTML = s
End Function

Public Function GetEstilosCSS_InformeEdicion_Print() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & ".report-section{ margin-top: 20px; }" & vbCrLf
    css = css & ".report-cover{ background: white; border: 1px solid var(--grey-2); border-radius: 18px; padding: 46px; }" & vbCrLf
    css = css & ".report-cover-title{ font-size: 34px; font-weight: 900; color: var(--grey-9); letter-spacing: 0.5px; }" & vbCrLf
    css = css & ".report-cover-subtitle{ margin-top: 6px; font-size: 18px; color: var(--grey-6); }" & vbCrLf
    css = css & ".report-cover-kv{ margin-top: 28px; display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; }" & vbCrLf
    css = css & ".kv{ border: 1px solid var(--grey-2); border-radius: 12px; padding: 12px 14px; }" & vbCrLf
    css = css & ".kv-k{ font-size: 12px; color: var(--grey-6); font-weight: 800; text-transform: uppercase; letter-spacing: 0.6px; }" & vbCrLf
    css = css & ".kv-v{ margin-top: 6px; font-size: 14px; color: var(--grey-9); font-weight: 700; }" & vbCrLf
    css = css & ".report-table{ width: 100%; border-collapse: collapse; }" & vbCrLf
    css = css & ".report-table th{ text-align: left; padding: 10px 10px; border-bottom: 2px solid var(--grey-2); background: #fcfcfc; color: var(--grey-9); white-space: nowrap; }" & vbCrLf
    css = css & ".report-table td{ padding: 10px 10px; border-bottom: 1px solid #f6f6f6; vertical-align: top; }" & vbCrLf
    css = css & ".report-table-small th, .report-table-small td{ font-size: 12px; }" & vbCrLf
    css = css & ".cc-table{ table-layout: fixed; }" & vbCrLf
    css = css & ".cc-table th{ white-space: normal; }" & vbCrLf
    css = css & ".cc-table td{ word-break: break-word; }" & vbCrLf
    css = css & ".cc-col-codigo{ width: 60px; }" & vbCrLf
    css = css & ".cc-col-edicion{ width: 40px; }" & vbCrLf
    css = css & ".cc-col-estado{ width: 20%; }" & vbCrLf
    css = css & ".cc-col-pm{ width: 35%; }" & vbCrLf
    css = css & ".cc-col-pc{ width: 35%; }" & vbCrLf
    css = css & ".cc-added{ color: #0b7a28; font-weight: 800; }" & vbCrLf
    css = css & ".cc-deleted{ color: #b42318; font-weight: 800; }" & vbCrLf
    css = css & ".meta-row{ display:flex; justify-content:space-between; gap: 16px; margin-bottom: 14px; color: var(--grey-6); font-size: 13px; }" & vbCrLf
    css = css & ".table-scroll{ overflow-x: auto; }" & vbCrLf
    css = css & ".print-footer{ display:none; }" & vbCrLf
    css = css & "@media print {" & vbCrLf
    css = css & "  @page{ margin: 16mm 14mm 22mm 14mm; }" & vbCrLf
    css = css & "  body{ background: white; counter-reset: page; }" & vbCrLf
    css = css & "  .main-container{ margin-top: 0; padding-bottom: 70px; }" & vbCrLf
    css = css & "  header{ border-bottom: none; }" & vbCrLf
    css = css & "  .print-page{ break-after: page; page-break-after: always; }" & vbCrLf
    css = css & "  .card{ box-shadow: none; }" & vbCrLf
    css = css & "  footer{ display:none; }" & vbCrLf
    css = css & "  .print-footer{ display:flex; align-items:flex-end; justify-content:space-between; gap: 16px; position: fixed; left: 0; right: 0; bottom: 0; padding: 10px 16px; border-top: 1px solid #e6e6e6; background: white; color: #000; font-size: 11px; }" & vbCrLf
    css = css & "  .print-footer-left{ display:flex; flex-direction:column; gap: 2px; }" & vbCrLf
    css = css & "  .print-footer-right{ white-space: nowrap; }" & vbCrLf
    css = css & "  .page-number:after{ content: counter(page); }" & vbCrLf
    css = css & "}" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_InformeEdicion_Print = css
End Function

Public Function GetEstilosCSS_Correos() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & "table { width: 100%; border-collapse: separate; border-spacing: 0; background: white; border: 1px solid var(--grey-2); border-radius: 18px; overflow: hidden; margin-bottom: 20px; }" & vbCrLf
    css = css & "table td, table th { padding: 10px 12px; border-bottom: 1px solid #f6f6f6; vertical-align: top; font-size: 13px; color: var(--grey-6); }" & vbCrLf
    css = css & "table tr:last-child td { border-bottom: none; }" & vbCrLf
    css = css & "td.Cabecera, td.ColespanArriba { background: #fcfcfc; color: var(--grey-9); font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; font-size: 12px; }" & vbCrLf
    css = css & "td.ColespanArriba { text-align: left; }" & vbCrLf
    css = css & "td.centrado { text-align: center; }" & vbCrLf
    css = css & "a { color: var(--tele-blue); text-decoration: none; }" & vbCrLf
    css = css & "a:hover { text-decoration: underline; }" & vbCrLf
    css = css & "p { margin: 0 0 12px 0; }" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_Correos = css
End Function

Private Function GuardarInformeEdicionHTML_UTF8(ByVal p_Edicion As Edicion, ByVal p_HTML As String, ByRef p_Error As String) As String
    Dim m_Ruta As String
    Dim m_Stream As Object
    Dim m_FSO As Object

    On Error GoTo errores
    p_Error = ""

    m_Ruta = GetURLInformeEdicionHTML(p_Edicion, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    If m_FSO.FileExists(m_Ruta) Then
        If FicheroAbierto(m_Ruta) Then
            p_Error = "Tiene abierto un informe anterior"
            Err.Raise 1000
        End If
    End If

    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2
    m_Stream.Charset = "utf-8"
    m_Stream.Open
    m_Stream.WriteText p_HTML
    m_Stream.SaveToFile m_Ruta, 2
    m_Stream.Close

    GuardarInformeEdicionHTML_UTF8 = m_Ruta
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error UTF-8: " & Err.Description
    End If
End Function

Private Function GetURLInformeEdicionHTML( _
                                        ByVal p_Edicion As Edicion, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_Cod As String
    Dim m_Proyecto As Proyecto

    On Error GoTo errores
    p_Error = ""

    If p_Edicion Is Nothing Then
        p_Error = "La edición hay que introducirla"
        Err.Raise 1000
    End If

    Set m_Proyecto = p_Edicion.Proyecto
    If m_Proyecto Is Nothing Then
        Set m_Proyecto = Constructor.getProyecto(p_Edicion.IDProyecto, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    End If

    m_Cod = Nz(m_Proyecto.CodigoDocumento, "")
    If m_Cod = "" Then
        p_Error = "No se sabe el código del documento del proyecto al que pertenece el informe"
        Err.Raise 1000
    End If

    m_Cod = Replace(m_Cod, "/", "_")
    m_Cod = Replace(m_Cod, "\", "_")
    m_Cod = Replace(m_Cod, vbNewLine, "")

    If m_Proyecto.Juridica = "TdE" Then
        GetURLInformeEdicionHTML = m_ObjEntorno.URLDirectorioLocal & m_Cod & "V" & Format(p_Edicion.Edicion, "00") & ".html"
    Else
        GetURLInformeEdicionHTML = m_ObjEntorno.URLDirectorioLocal & m_Cod & "-" & p_Edicion.Edicion & ".html"
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GetURLInformeEdicionHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

' -------------------------------------------------------------------------
' COPIA TU CÓDIGO SVG AQUÍ
' -------------------------------------------------------------------------
Public Function GetLogoSVG() As String
    ' SVG Logo Telefónica (T de Puntos - Ancho) - Fondo Azul / Logo Blanco
    Dim svg As String
    svg = "<svg id='Layer_1' data-name='Layer 1' xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1920 802'>" & _
          "<defs><style>.cls-1{fill:#0066FF;}.cls-2{fill:#FFFFFF;}</style></defs>" & _
          "<rect class='cls-1' x='-13' y='-15' width='1946' height='832'/>" & _
          "<g><g>" & _
          "<circle class='cls-2' cx='275.37' cy='277.85' r='52.41'/>" & _
          "<circle class='cls-2' cx='398.52' cy='277.85' r='52.41'/>" & _
          "<circle class='cls-2' cx='521.67' cy='277.85' r='52.41'/>" & _
          "<circle class='cls-2' cx='398.52' cy='401' r='52.41'/>" & _
          "<circle class='cls-2' cx='398.52' cy='524.15' r='52.41'/>" & _
          "</g><g>" & _
          "<path class='cls-2' d='m685.79,360.16h-48.84v-29.07h127.9v29.07h-48.84v133.72h-30.23v-133.72Z'/>" & _
          "<path class='cls-2' d='m866.94,458.99c-2.56,9.3-17.21,37.21-54.65,37.21-34.88,0-60.47-25.58-60.47-61.63s25.58-61.63,60.47-61.63c32.56,0,58.14,25.58,58.14,59.3,0,3.49-.47,6.28-.7,8.37l-.47,3.26h-88.37c2.56,16.51,14.88,27.91,31.4,27.91,13.72,0,22.09-7.9,24.42-12.79h30.23Zm-25.58-34.88c-2.56-15.12-12.56-26.75-29.07-26.75-17.67,0-27.91,11.63-31.4,26.75h60.47Z'/>" & _
          "<path class='cls-2' d='m887.85,331.09h29.07v162.79h-29.07v-162.79Z'/>" & _
          "<path class='cls-2' d='m1049.46,458.99c-2.56,9.3-17.21,37.21-54.65,37.21-34.88,0-60.46-25.58-60.46-61.63s25.58-61.63,60.46-61.63c32.56,0,58.14,25.58,58.14,59.3,0,3.49-.47,6.28-.7,8.37l-.47,3.26h-88.37c2.56,16.51,14.88,27.91,31.4,27.91,13.72,0,22.09-7.9,24.42-12.79h30.23Zm-25.58-34.88c-2.56-15.12-12.56-26.75-29.07-26.75-17.67,0-27.91,11.63-31.4,26.75h60.47Z'/>" & _
          "<path class='cls-2' d='m1079.68,403.18h-19.77v-27.91h19.77v-17.44c0-17.67,11.4-29.07,29.07-29.07h25.58v25.58h-17.44c-4.65,0-8.14,3.49-8.14,8.14v12.79h25.58v27.91h-25.58v90.7h-29.07v-90.7Z'/>" & _
          "<path class='cls-2' d='m1259.61,434.58c0,36.05-25.58,61.63-60.47,61.63s-60.47-25.58-60.47-61.63,25.58-61.63,60.47-61.63,60.47,25.58,60.47,61.63Zm-29.07,0c0-20.93-13.95-34.88-31.4-34.88s-31.4,13.95-31.4,34.88,13.95,34.88,31.4,34.88,31.4-13.95,31.4-34.88Z'/>" & _
          "<path class='cls-2' d='m1277.03,375.28h26.75l2.32,11.63h1.16c2.09-2.56,4.89-4.88,7.91-6.98,5.35-3.49,13.49-6.98,24.65-6.98,26.74,0,46.51,19.77,46.51,50v70.93h-29.07v-68.6c0-15.12-10.46-25.58-25.58-25.58s-25.58,10.46-25.58,25.58v68.6h-29.07v-118.6Z'/>" & _
          "<path class='cls-2' d='m1571.36,449.69c-3.72,18.6-18.37,46.51-55.81,46.51-34.88,0-60.46-25.58-60.46-61.63s25.58-61.63,60.46-61.63c37.44,0,52.09,27.91,55.81,45.35h-29.07c-2.56-6.75-9.3-18.6-26.75-18.6s-31.4,13.95-31.4,34.88,13.95,34.88,31.4,34.88,24.19-11.63,26.75-19.77h29.07Z'/>" & _
          "<path class='cls-2' d='m1659.26,482.25h-1.16c-2.09,2.56-4.88,4.88-8.14,6.98-5.58,3.49-13.72,6.98-25.58,6.98-26.97,0-43.02-16.51-43.02-36.04,0-23.26,16.28-39.54,48.83-39.54h26.75v-2.32c0-13.02-7.91-22.09-20.93-22.09s-19.77,8.37-20.93,15.12h-29.07c2.56-19.54,18.37-38.37,50-38.37s50,20,50,45.35v75.58h-24.42l-2.32-11.63Zm-2.32-39.53h-24.42c-15.12,0-22.09,5.81-22.09,15.11s6.75,15.12,18.6,15.12c17.67,0,27.91-10.23,27.91-26.75v-3.49Z'/>" & _
          "<path class='cls-2' d='m1408.59,375.28h29.07v118.6h-29.07v-118.6Z'/>" & _
          "<circle class='cls-2' cx='1423.13' cy='341.02' r='17.2'/>" & _
          "<polygon class='cls-2' points='1209.61 325.28 1239.85 325.28 1211.93 359 1187.52 359 1209.61 325.28'/>" & _
          "</g></g>" & _
          "</svg>"
    
    GetLogoSVG = svg
End Function

Private Function ConstruirTablaEstadosHistoricosHTML( _
                                                    ByVal p_Riesgo As riesgo, _
                                                    ByVal p_FechaCierre As String, _
                                                    ByVal p_FechaPublicacion As String, _
                                                    ByRef p_Error As String _
                                                    ) As String
    Dim ColEstados As Scripting.Dictionary
    Dim k As Variant
    Dim s As String
    Dim sEstado As String, sFecha As String
    Dim partes As Variant
    Dim m_FechaPub As String

    On Error GoTo errores
    
    If Not IsDate(p_FechaPublicacion) Then
        m_FechaPub = Format(Date, "dd/mm/yyyy")
    Else
        m_FechaPub = p_FechaPublicacion
    End If

    Set ColEstados = getEstadosDiferentesHastaEdicion(p_Riesgo.Edicion, p_Riesgo.CodigoRiesgo, m_FechaPub, p_FechaCierre, p_Error)
    If p_Error <> "" Then Exit Function
    If ColEstados Is Nothing Then Exit Function
    If ColEstados.Count = 0 Then Exit Function

    s = "<div class='card' style='margin-top:18px; border: 1px solid var(--grey-2);'>" & vbCrLf
    s = s & "  <h3 style='margin-top:0;'>Histórico de Estados</h3>" & vbCrLf
    s = s & "  <table class='report-table report-table-small'>" & vbCrLf
    s = s & "    <thead><tr><th>Estado</th><th>Fecha</th></tr></thead>" & vbCrLf
    s = s & "    <tbody>" & vbCrLf

    For Each k In ColEstados
        partes = Split(ColEstados(k), "|")
        sEstado = partes(0)
        sFecha = ""
        If UBound(partes) >= 1 Then sFecha = partes(1)
        
        If IsDate(sFecha) Then sFecha = Format(sFecha, "dd/mm/yyyy")
        
        s = s & "      <tr>" & vbCrLf
        s = s & "        <td>" & HTMLSafe(sEstado) & "</td>" & vbCrLf
        s = s & "        <td>" & HTMLSafe(sFecha) & "</td>" & vbCrLf
        s = s & "      </tr>" & vbCrLf
    Next k

    s = s & "    </tbody>" & vbCrLf
    s = s & "  </table>" & vbCrLf
    s = s & "</div>" & vbCrLf

    ConstruirTablaEstadosHistoricosHTML = s
    Exit Function

errores:
    p_Error = "Error en ConstruirTablaEstadosHistoricosHTML: " & Err.Description
End Function

Private Function ConstruirSeccionGeneral(p_Riesgo As riesgo) As String
    Dim s As String
    Dim sTecnico As String
    Dim sEdicion As String, sEsActivo As String
    Dim sTituloCausaRaiz As String
    Dim sCausaRaiz As String
    
    ' Obtención segura del técnico de calidad
    On Error Resume Next
    sTecnico = p_Riesgo.Edicion.Proyecto.UsuarioCalidad.Nombre
    If Err.Number <> 0 Then sTecnico = "-"
    On Error GoTo 0
    
    ' Obtención segura de Edición
    On Error Resume Next
    sEdicion = p_Riesgo.Edicion.Edicion
    If p_Riesgo.Edicion.EsActivo = 1 Then sEsActivo = "Sí" Else sEsActivo = "No"
    On Error GoTo 0
    
    ' Fila 1: KPIs Principales (Compacto - Una sola fila)
    s = "<div class='grid-compact' style='margin-bottom:20px;'>" & vbCrLf
    
    ' 1. Estado
    s = s & "    <div class='card-mini'><h3>Estado</h3><div class='kpi-value-mini' style='color:var(--tele-blue);'>" & HTMLSafe(p_Riesgo.ESTADOCalculadoTexto) & "</div></div>" & vbCrLf
    ' 2. Priorización
    s = s & "    <div class='card-mini'><h3>Priorización</h3><div class='kpi-value-mini'>" & p_Riesgo.Priorizacion & "</div></div>" & vbCrLf
    ' 3. Impacto Global
    s = s & "    <div class='card-mini'><h3>Impacto Global</h3><div class='kpi-value-mini'>" & HTMLSafe(p_Riesgo.ImpactoGlobalCalculado) & "</div></div>" & vbCrLf
    ' 4. Origen
    s = s & "    <div class='card-mini'><h3>Origen</h3><div class='kpi-value-mini' style='font-weight:normal;'>" & HTMLSafe(p_Riesgo.Origen) & "</div></div>" & vbCrLf
    ' 5. Edición
    s = s & "    <div class='card-mini'><h3>Edición</h3><div class='kpi-value-mini'>" & HTMLSafe(sEdicion) & "</div></div>" & vbCrLf
    ' 6. ¿Activa?
    s = s & "    <div class='card-mini'><h3>¿Activa?</h3><div class='kpi-value-mini' style='color:" & IIf(sEsActivo = "Sí", "#00C853", "#D50000") & ";'>" & sEsActivo & "</div></div>" & vbCrLf
    
    s = s & "</div>" & vbCrLf
    
    ' Fila 2: Fechas y Personas
    s = s & "<div class='grid' style='margin-top:20px;'>" & vbCrLf
    s = s & "    <div class='card'><h4>Detalles de Detección</h4>" & _
            "<p><strong>Fecha Detectado:</strong> " & FormatoFecha(p_Riesgo.FechaDetectado) & "</p>" & _
            "<p><strong>Detectado Por:</strong> " & HTMLSafe(p_Riesgo.DetectadoPor) & "</p></div>" & vbCrLf
            
    s = s & "    <div class='card'><h4>Gestión y Calidad</h4>" & _
            "<p><strong>Técnico Calidad:</strong> " & HTMLSafe(sTecnico) & "</p>" & _
            "<p><strong>Fecha Aceptado:</strong> " & FormatoFecha(p_Riesgo.FechaMitigacionAceptar) & "</p>" & _
            "<p><strong>Fecha Retirado:</strong> " & FormatoFecha(p_Riesgo.FechaRetirado) & "</p></div>" & vbCrLf
    s = s & "</div>" & vbCrLf

    ' Fila 3: Textos Largos
    s = s & "<div class='card' style='margin-top:20px;'>" & vbCrLf
    s = s & "    <h2>Descripción de Riesgo</h2><p>" & HTMLSafeLargo(p_Riesgo.Descripcion) & "</p>" & vbCrLf
    sCausaRaiz = Nz(p_Riesgo.CausaRaiz, "")
    sTituloCausaRaiz = "Análisis Causa Raíz"
    If Trim$(sCausaRaiz) = "" Then
        sTituloCausaRaiz = sTituloCausaRaiz & " (No aplica)"
        sCausaRaiz = "-"
    End If
    s = s & "    <h2 style='margin-top:20px;'>" & sTituloCausaRaiz & "</h2><p>" & HTMLSafeLargo(sCausaRaiz) & "</p>" & vbCrLf
    s = s & "</div>" & vbCrLf
    
    ConstruirSeccionGeneral = s
End Function

Private Function ConstruirSeccionPlanes(p_Riesgo As riesgo, Tipo As EnumTipoPlan) As String
    Dim s As String, colPlanes As Scripting.Dictionary, m_IDP As Variant, m_IDPA As Variant
    Dim m_Plan As Object, m_Accion As Object, m_NombrePropiedad As String
    
    ' Título de la sección y Selección de Colección
    If Tipo = Mitigacion Then
        Set colPlanes = p_Riesgo.ColPMs
        s = "<h2>Planes de Mitigación</h2>"
    Else
        Set colPlanes = p_Riesgo.ColPCs
        s = "<h2>Planes de Contingencia</h2>"
    End If
    
    If colPlanes Is Nothing Then
        s = s & "<div class='card'><p>No hay planes registrados para este tipo.</p></div>"
    ElseIf colPlanes.Count = 0 Then
        s = s & "<div class='card'><p>No hay planes registrados para este tipo.</p></div>"
    Else
        For Each m_IDP In colPlanes
            Set m_Plan = colPlanes(m_IDP)
            m_NombrePropiedad = "Cod" & IIf(Tipo = Mitigacion, "Mitigacion", "Contingencia")
            
            s = s & "<div class='card' style='border-left: 6px solid var(--tele-blue); margin-bottom: 25px;'>" & vbCrLf
            
            ' Cabecera del Plan con Estado
            s = s & "    <h3 style='color: var(--tele-blue); display:flex; justify-content:space-between;'>" & _
                    "<span>" & HTMLSafe(m_Plan.getPropiedad(m_NombrePropiedad)) & ": " & HTMLSafe(m_Plan.DisparadorDelPlan) & "</span>" & _
                    "<span style='font-size:0.8em; background:#f0f0f0; padding:2px 8px; border-radius:4px; color:#333;'>" & HTMLSafe(m_Plan.ESTADOCalculadoTexto) & "</span>" & _
                    "</h3>" & vbCrLf
            
            If Not m_Plan.colAcciones Is Nothing Then
                If m_Plan.colAcciones.Count > 0 Then
                    s = s & "    <table class='action-table'>" & vbCrLf
                    s = s & "        <thead><tr>" & _
                            "<th>Acción</th>" & _
                            "<th>Responsable</th>" & _
                            "<th>Fecha Inicio</th>" & _
                            "<th>F. Fin Prevista</th>" & _
                            "<th>F. Fin Real</th>" & _
                            "</tr></thead><tbody>" & vbCrLf
                            
                    For Each m_IDPA In m_Plan.colAcciones
                        Set m_Accion = m_Plan.colAcciones(m_IDPA)
                        s = s & "        <tr>" & _
                                "<td><strong>" & HTMLSafe(m_Accion.Accion) & "</strong></td>" & _
                                "<td>" & HTMLSafe(m_Accion.ResponsableAccion) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaInicio) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaFinPrevista) & "</td>" & _
                                "<td>" & FormatoFecha(m_Accion.FechaFinReal) & "</td>" & _
                                "</tr>" & vbCrLf
                    Next
                    s = s & "    </tbody></table>" & vbCrLf
                End If
            End If
            s = s & "</div>" & vbCrLf
        Next
    End If
    ConstruirSeccionPlanes = s
End Function

Private Function ConstruirSeccionMaterializaciones(p_Riesgo As riesgo) As String
    Dim s As String
    Dim col As Scripting.Dictionary
    Dim clavesOrdenadas As Variant
    Dim i As Long
    Dim m_Mat As RiesgoMaterializacion
    Dim FechaInicio As String
    Dim fechaFin As String
    
    On Error Resume Next
    Set col = p_Riesgo.ColMaterializaciones
    On Error GoTo 0
    
    s = "<h2>Materializaciones</h2>"
    
    If col Is Nothing Then
        s = s & "<div class='card'><p>No hay materializaciones registradas para este riesgo.</p></div>"
        ConstruirSeccionMaterializaciones = s
        Exit Function
    End If
    
    If col.Count = 0 Then
        s = s & "<div class='card'><p>No hay materializaciones registradas para este riesgo.</p></div>"
        ConstruirSeccionMaterializaciones = s
        Exit Function
    End If
    
    clavesOrdenadas = OrdenarClavesMaterializacionesPorFecha(col)
    
    s = s & "<div class='card'>"
    s = s & "<table class='action-table'>"
    s = s & "<thead><tr><th>Fecha materialización</th><th>Fecha desmaterialización</th><th>Es materialización</th></tr></thead>"
    s = s & "<tbody>"
    
    FechaInicio = ""
    fechaFin = ""
    
    For i = LBound(clavesOrdenadas) To UBound(clavesOrdenadas)
        Set m_Mat = col(clavesOrdenadas(i))
        If EsTextoSi(m_Mat.EsMaterializacion) Then
            If FechaInicio <> "" Then
                s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>-</td><td>Sí</td></tr>"
            End If
            FechaInicio = m_Mat.Fecha
            fechaFin = ""
        Else
            If FechaInicio <> "" Then
                fechaFin = m_Mat.Fecha
                s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>" & FormatoFecha(fechaFin) & "</td><td>Sí</td></tr>"
                FechaInicio = ""
                fechaFin = ""
            End If
        End If
        Set m_Mat = Nothing
    Next
    
    If FechaInicio <> "" Then
        s = s & "<tr><td>" & FormatoFecha(FechaInicio) & "</td><td>-</td><td>Sí</td></tr>"
    End If
    
    s = s & "</tbody></table>"
    s = s & "</div>"
    
    ConstruirSeccionMaterializaciones = s
End Function

Private Function OrdenarClavesMaterializacionesPorFecha(ByVal p_Col As Scripting.Dictionary) As Variant
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant
    
    keys = p_Col.keys
    
    If UBound(keys) <= LBound(keys) Then
        OrdenarClavesMaterializacionesPorFecha = keys
        Exit Function
    End If
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If FechaSerial(p_Col(keys(i)).Fecha) > FechaSerial(p_Col(keys(j)).Fecha) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    
    OrdenarClavesMaterializacionesPorFecha = keys
End Function

Private Function FechaSerial(ByVal p_Valor As Variant) As Double
    If IsDate(p_Valor) Then
        FechaSerial = CDbl(CDate(p_Valor))
    Else
        FechaSerial = 0
    End If
End Function

Private Function EsTextoSi(ByVal p_Valor As Variant) As Boolean
    Dim t As String
    t = Nz(p_Valor, "")
    If Len(t) = 0 Then
        EsTextoSi = False
        Exit Function
    End If
    EsTextoSi = (UCase$(Left$(t, 1)) = "S")
End Function

Private Function ConstruirSeccionPublicabilidad(p_Riesgo As riesgo) As String
    Dim s As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Publicable As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto
    Dim m_Key As Variant
    Dim m_Check As Scripting.Dictionary
    Dim m_Estado As EnumPublicabilidadCheckEstado
    Dim m_Label As String
    Dim m_Clase As String
    Dim m_Detalle As String
    Dim m_VeredictoTexto As String
    Dim m_VeredictoClase As String

    On Error GoTo errores

    If ConstruirDatosPublicabilidadRiesgo(p_Riesgo, m_Datos, , m_Error) = EnumSiNo.No Then
        s = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(m_Error) & "</p></div>"
        ConstruirSeccionPublicabilidad = s
        Exit Function
    End If

    m_Publicable = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        s = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(m_Error) & "</p></div>"
        ConstruirSeccionPublicabilidad = s
        Exit Function
    End If

    Select Case m_Veredicto
        Case EnumPublicabilidadVeredicto.Publicable
            m_VeredictoTexto = "Publicable"
            m_VeredictoClase = "verdict-publicable"
        Case EnumPublicabilidadVeredicto.NoPublicable
            m_VeredictoTexto = "No publicable"
            m_VeredictoClase = "verdict-no-publicable"
        Case Else
            m_VeredictoTexto = "No aplica"
            m_VeredictoClase = "verdict-no-aplica"
    End Select

    s = "<div class='card verdict-card'>"
    s = s & "<div><h2 style='margin:0 0 6px 0;'>Publicabilidad</h2>" & _
            "<div style='color:var(--grey-6); font-size:13px;'>Resumen de comprobaciones</div></div>"
    s = s & "<div class='verdict-badge " & m_VeredictoClase & "'>" & m_VeredictoTexto & "</div>"
    s = s & "</div>"

    s = s & "<div class='checklist'>"
    If Not m_Checks Is Nothing Then
        For Each m_Key In m_Checks
            Set m_Check = m_Checks(m_Key)
            m_Estado = m_Check("estado")
            m_Label = EstadoPublicabilidadLabel(m_Estado)
            m_Clase = EstadoPublicabilidadClass(m_Estado)

            s = s & "<div class='check-item'>"
            s = s & "<div class='check-left'><div class='check-text'>" & HTMLSafe(CStr(m_Check("texto"))) & "</div>"
            If m_Check.Exists("detalle") Then
                m_Detalle = CStr(m_Check("detalle"))
                If m_Detalle <> "" Then
                    s = s & "<div class='check-detail'>" & HTMLSafe(m_Detalle) & "</div>"
                End If
            End If
            s = s & "</div>"
            s = s & "<div class='check-state " & m_Clase & "'>" & m_Label & "</div>"
            s = s & "</div>"
        Next
    End If
    s = s & "</div>"

    ConstruirSeccionPublicabilidad = s
    Exit Function

errores:
    ConstruirSeccionPublicabilidad = "<div class='card'><p>Error al evaluar publicabilidad: " & HTMLSafe(Err.Description) & "</p></div>"
End Function

Private Function EstadoPublicabilidadLabel(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String
    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            EstadoPublicabilidadLabel = "Cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            EstadoPublicabilidadLabel = "No cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            EstadoPublicabilidadLabel = "No aplica"
        Case Else
            EstadoPublicabilidadLabel = "Desconocido"
    End Select
End Function

Private Function EstadoPublicabilidadClass(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String
    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            EstadoPublicabilidadClass = "state-cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            EstadoPublicabilidadClass = "state-no-cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            EstadoPublicabilidadClass = "state-no-aplica"
        Case Else
            EstadoPublicabilidadClass = "state-no-aplica"
    End Select
End Function
Public Function GetEstilosCSS_Corporativos() As String
    Dim css As String
    css = "<style>" & vbCrLf
    css = css & ":root { --tele-blue: #0066FF; --tele-white: #F2F4FF; --pure-white: #FFFFFF; --grey-9: #031A34; --grey-6: #58617A; --grey-2: #D1D5E4; }" & vbCrLf
    css = css & "body { font-family: 'Segoe UI', Arial, sans-serif; background-color: var(--tele-white); color: var(--grey-6); margin: 0; padding: 0; }" & vbCrLf
    css = css & "header { background-color: var(--tele-blue); color: white; padding: 25px 0; border-bottom: 4px solid #0055D4; }" & vbCrLf
    css = css & ".container { max-width: 1200px; margin: 0 auto; padding: 0 24px; }" & vbCrLf
    css = css & ".header-container { display: flex; justify-content: space-between; align-items: center; }" & vbCrLf
    css = css & ".logo-container { display: flex; align-items: center; gap: 20px; }" & vbCrLf
    css = css & ".logo-wrapper { width: 140px; height: 60px; display: flex; align-items: center; justify-content: flex-start; }" & vbCrLf
    css = css & ".logo-wrapper svg { width: 100%; height: 100%; object-fit: contain; }" & vbCrLf
    css = css & ".header-text h1 { font-size: 26px; margin: 0; font-weight: 800; }" & vbCrLf
    css = css & ".header-text p { margin: 0; opacity: 0.85; font-size: 14px; }" & vbCrLf
    css = css & ".main-container { margin-top: 35px; padding-bottom: 60px; }" & vbCrLf
    css = css & ".tab-container { display: flex; gap: 10px; }" & vbCrLf
    css = css & ".tab-btn { padding: 14px 28px; border: 1px solid var(--grey-2); background: #E8EBF7; cursor: pointer; border-radius: 12px 12px 0 0; font-weight: bold; color: var(--grey-6); transition: 0.3s; }" & vbCrLf
    css = css & ".tab-btn.active { background: var(--pure-white); border-bottom: 2px solid var(--pure-white); color: var(--tele-blue); }" & vbCrLf
    css = css & ".tab-content { display: none; background: var(--pure-white); padding: 35px; border: 1px solid var(--grey-2); border-radius: 0 24px 24px 24px; box-shadow: 0 4px 15px rgba(0,0,0,0.06); }" & vbCrLf
    css = css & ".tab-content.active { display: block; animation: fadeIn 0.4s; }" & vbCrLf
    css = css & ".subtab-container { display:flex; gap:10px; margin: 0 0 18px 0; }" & vbCrLf
    css = css & ".subtab-btn { padding: 10px 18px; border: 1px solid var(--grey-2); background: #F3F5FC; cursor: pointer; border-radius: 12px; font-weight: 700; color: var(--grey-6); transition: 0.3s; }" & vbCrLf
    css = css & ".subtab-btn.active { background: var(--pure-white); color: var(--tele-blue); border-color: var(--tele-blue); }" & vbCrLf
    css = css & ".subtab-content { display:none; }" & vbCrLf
    css = css & ".subtab-content.active { display:block; }" & vbCrLf
    css = css & ".card { background: white; padding: 25px; border-radius: 18px; border: 1px solid var(--grey-2); }" & vbCrLf
    css = css & ".card-mini { background: white; padding: 15px; border-radius: 12px; border: 1px solid var(--grey-2); display: flex; flex-direction: column; justify-content: center; min-width: 0; }" & vbCrLf
    css = css & ".verdict-card { display: flex; align-items: center; justify-content: space-between; gap: 15px; margin-bottom: 20px; }" & vbCrLf
    css = css & ".verdict-badge { padding: 8px 14px; border-radius: 999px; font-weight: bold; font-size: 12px; letter-spacing: 0.5px; text-transform: uppercase; }" & vbCrLf
    css = css & ".verdict-publicable { background: #E8F5E9; color: #1B5E20; }" & vbCrLf
    css = css & ".verdict-no-publicable { background: #FFEBEE; color: #B71C1C; }" & vbCrLf
    css = css & ".verdict-no-aplica { background: #ECEFF1; color: #455A64; }" & vbCrLf
    css = css & ".checklist { display: grid; gap: 12px; }" & vbCrLf
    css = css & ".check-item { border: 1px solid var(--grey-2); border-radius: 12px; padding: 12px 14px; display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; }" & vbCrLf
    css = css & ".check-left { flex: 1; }" & vbCrLf
    css = css & ".check-text { font-weight: 600; color: var(--grey-9); }" & vbCrLf
    css = css & ".check-detail { margin-top: 4px; font-size: 12px; color: var(--grey-6); }" & vbCrLf
    css = css & ".check-state { font-size: 11px; font-weight: bold; padding: 4px 8px; border-radius: 999px; text-transform: uppercase; }" & vbCrLf
    css = css & ".state-cumple { background: #E8F5E9; color: #1B5E20; }" & vbCrLf
    css = css & ".state-no-cumple { background: #FFEBEE; color: #B71C1C; }" & vbCrLf
    css = css & ".state-no-aplica { background: #ECEFF1; color: #455A64; }" & vbCrLf
    css = css & ".grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 25px; }" & vbCrLf
    css = css & ".grid-compact { display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; }" & vbCrLf
    css = css & "@media (max-width: 1100px) { .grid-compact { grid-template-columns: repeat(3, 1fr); } }" & vbCrLf
    css = css & "@media (max-width: 600px) { .grid-compact { grid-template-columns: repeat(2, 1fr); } }" & vbCrLf
    css = css & "h2 { color: var(--grey-9); border-bottom: 2px solid var(--tele-blue); padding-bottom: 8px; margin-bottom: 20px; }" & vbCrLf
    css = css & "h3 { font-size: 14px; margin: 0 0 5px 0; color: var(--grey-6); text-transform: uppercase; letter-spacing: 0.5px; }" & vbCrLf
    css = css & ".kpi-value { font-size: 30px; font-weight: bold; color: var(--grey-9); }" & vbCrLf
    css = css & ".kpi-value-mini { font-size: 20px; font-weight: bold; color: var(--grey-9); line-height: 1.2; }" & vbCrLf
    css = css & ".action-table { width: 100%; border-collapse: collapse; margin-top: 20px; }" & vbCrLf
    css = css & ".action-table th { text-align: left; padding: 12px; border-bottom: 2px solid var(--grey-2); background: #fcfcfc; }" & vbCrLf
    css = css & ".action-table td { padding: 12px; border-bottom: 1px solid #f6f6f6; font-size: 14px; }" & vbCrLf
    css = css & "footer { text-align: center; padding: 40px; color: var(--grey-6); font-size: 12px; }" & vbCrLf
    css = css & "@keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }" & vbCrLf
    css = css & "</style>" & vbCrLf
    GetEstilosCSS_Corporativos = css
End Function

Private Function GetScriptsJS() As String
    GetScriptsJS = "<script>" & vbCrLf & _
                   "function openTab(evt, tabName) {" & vbCrLf & _
                   "  var i, tabcontent, tablinks;" & vbCrLf & _
                   "  tabcontent = document.getElementsByClassName('tab-content');" & vbCrLf & _
                   "  for (i = 0; i < tabcontent.length; i++) { tabcontent[i].classList.remove('active'); }" & vbCrLf & _
                   "  tablinks = document.getElementsByClassName('tab-btn');" & vbCrLf & _
                   "  for (i = 0; i < tablinks.length; i++) { tablinks[i].classList.remove('active'); }" & vbCrLf & _
                   "  document.getElementById(tabName).classList.add('active');" & vbCrLf & _
                   "  evt.currentTarget.classList.add('active');" & vbCrLf & _
                   "}" & vbCrLf & _
                   "function openSubTab(evt, tabName, parentId) {" & vbCrLf & _
                   "  var i, parent, tabcontent, tablinks;" & vbCrLf & _
                   "  parent = document.getElementById(parentId);" & vbCrLf & _
                   "  if (!parent) { return; }" & vbCrLf & _
                   "  tabcontent = parent.getElementsByClassName('subtab-content');" & vbCrLf & _
                   "  for (i = 0; i < tabcontent.length; i++) { tabcontent[i].classList.remove('active'); }" & vbCrLf & _
                   "  tablinks = parent.getElementsByClassName('subtab-btn');" & vbCrLf & _
                   "  for (i = 0; i < tablinks.length; i++) { tablinks[i].classList.remove('active'); }" & vbCrLf & _
                   "  document.getElementById(tabName).classList.add('active');" & vbCrLf & _
                   "  evt.currentTarget.classList.add('active');" & vbCrLf & _
                   "}" & vbCrLf & _
                   "</script>"
End Function

Private Function ConstruirSeccionDatosProyecto(p_Riesgo As riesgo) As String
    Dim s As String

    s = "<h2>Datos de Proyecto</h2>"
    s = s & "<div class='subtab-container'>" & vbCrLf
    s = s & "  <button class='subtab-btn active' onclick=""openSubTab(event, 'dp_generales', 'datosProyecto')"">Datos generales</button>" & vbCrLf
    s = s & "  <button class='subtab-btn' onclick=""openSubTab(event, 'dp_responsables', 'datosProyecto')"">Responsables</button>" & vbCrLf
    s = s & "</div>" & vbCrLf

    s = s & "<div id='dp_generales' class='subtab-content active'>" & vbCrLf
    s = s & ConstruirSeccionDatosProyectoGenerales(p_Riesgo)
    s = s & "</div>" & vbCrLf

    s = s & "<div id='dp_responsables' class='subtab-content'>" & vbCrLf
    s = s & ConstruirSeccionDatosProyectoResponsables(p_Riesgo)
    s = s & "</div>" & vbCrLf

    ConstruirSeccionDatosProyecto = s
End Function

Private Function ConstruirSeccionDatosProyectoGenerales(p_Riesgo As riesgo) As String
    Dim s As String
    Dim sNemotecnico As String
    Dim sTitulo As String
    Dim sCliente As String
    Dim sNombreProyecto As String
    Dim sJuridica As String
    Dim sEnUTE As String
    Dim sFechaFirmaContrato As String
    Dim sFechaPrevistaCierre As String
    Dim sFechaCierre As String
    Dim sFechaProximaPublicacion As String
    Dim sTextoFechaProximaPublicacion As String
    Dim sCodigoDocumento As String

    sNemotecnico = "-": sTitulo = "-": sCliente = "-": sNombreProyecto = "-": sJuridica = "-"
    sEnUTE = "-": sFechaFirmaContrato = "": sFechaPrevistaCierre = "": sFechaCierre = "": sFechaProximaPublicacion = "": sTextoFechaProximaPublicacion = "-": sCodigoDocumento = "-"

    On Error Resume Next
    sNemotecnico = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Nemotecnico, "-")
    sTitulo = Nz(p_Riesgo.Edicion.Proyecto.Expediente.Titulo, "-")
    sCliente = Nz(p_Riesgo.Edicion.Proyecto.Cliente, "-")
    sNombreProyecto = Nz(p_Riesgo.Edicion.Proyecto.NombreProyecto, "-")
    sJuridica = Nz(p_Riesgo.Edicion.Proyecto.Juridica, "-")
    sEnUTE = Nz(p_Riesgo.Edicion.Proyecto.EnUTECalculado, "-")
    sFechaFirmaContrato = Nz(p_Riesgo.Edicion.Proyecto.Expediente.FechaFirmaContrato, "")
    If sFechaFirmaContrato = "" Then sFechaFirmaContrato = Nz(p_Riesgo.Edicion.Proyecto.FechaFirmaContrato, "")
    sFechaPrevistaCierre = Nz(p_Riesgo.Edicion.Proyecto.FechaPrevistaCierre, "")
    sFechaCierre = Nz(p_Riesgo.Edicion.Proyecto.FechaCierre, "")
    sFechaProximaPublicacion = Nz(p_Riesgo.Edicion.Proyecto.FechaMaxProximaPublicacion, "")
    sCodigoDocumento = Nz(p_Riesgo.Edicion.Proyecto.CodigoDocumento, "-")
    On Error GoTo 0

    If IsDate(sFechaCierre) Then
        sTextoFechaProximaPublicacion = "No aplica"
    Else
        sTextoFechaProximaPublicacion = FormatoFecha(sFechaProximaPublicacion)
    End If

    s = "<div class='grid'>" & vbCrLf
    s = s & "  <div class='card'><h4>Identificación</h4>" & _
            "<p><strong>Nemotécnico:</strong> " & HTMLSafe(sNemotecnico) & "</p>" & _
            "<p><strong>Título:</strong> " & HTMLSafe(sTitulo) & "</p>" & _
            "<p><strong>Código Documento:</strong> " & HTMLSafe(sCodigoDocumento) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Datos generales</h4>" & _
            "<p><strong>Cliente:</strong> " & HTMLSafe(sCliente) & "</p>" & _
            "<p><strong>Nombre Proyecto:</strong> " & HTMLSafe(sNombreProyecto) & "</p>" & _
            "<p><strong>Jurídica:</strong> " & HTMLSafe(sJuridica) & "</p>" & _
            "<p><strong>En UTE:</strong> " & HTMLSafe(sEnUTE) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Fechas</h4>" & _
            "<p><strong>Firma contrato:</strong> " & FormatoFecha(sFechaFirmaContrato) & "</p>" & _
            "<p><strong>Prevista cierre:</strong> " & FormatoFecha(sFechaPrevistaCierre) & "</p>" & _
            "<p><strong>Cierre:</strong> " & FormatoFecha(sFechaCierre) & "</p>" & _
            "<p><strong>Próxima publicación:</strong> " & HTMLSafe(sTextoFechaProximaPublicacion) & "</p></div>" & vbCrLf

    s = s & "</div>" & vbCrLf
    ConstruirSeccionDatosProyectoGenerales = s
End Function

Private Function ConstruirSeccionDatosProyectoResponsables(p_Riesgo As riesgo) As String
    Dim s As String
    Dim sCorreoResponsableCalidad As String
    Dim sTecnicoCalidad As String
    Dim sRACs As String
    Dim sCorreosRACs As String
    Dim sAutorizados As String

    sCorreoResponsableCalidad = "-": sTecnicoCalidad = "-"
    sRACs = "-": sCorreosRACs = "-": sAutorizados = "-"

    On Error Resume Next
    sTecnicoCalidad = Nz(p_Riesgo.Edicion.Proyecto.UsuarioCalidad.Nombre, "-")
    sCorreoResponsableCalidad = Nz(p_Riesgo.Edicion.Proyecto.CorreoResponsableCalidad, "-")
    sRACs = Nz(p_Riesgo.Edicion.Proyecto.Expediente.CadenaRACs, "-")
    sCorreosRACs = Nz(p_Riesgo.Edicion.Proyecto.Expediente.CadenaCorreoRACs, "-")
    sAutorizados = Nz(p_Riesgo.Edicion.Proyecto.CadenaNombreAutorizados, "-")
    On Error GoTo 0

    If sAutorizados <> "-" Then
        sAutorizados = Replace(sAutorizados, "|", vbCrLf)
    End If

    s = "<div class='grid'>" & vbCrLf
    s = s & "  <div class='card'><h4>Calidad</h4>" & _
            "<p><strong>Técnico Calidad:</strong> " & HTMLSafe(sTecnicoCalidad) & "</p>" & _
            "<p><strong>Correo Responsable:</strong> " & HTMLSafe(sCorreoResponsableCalidad) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>RAC</h4>" & _
            "<p><strong>RACs:</strong><br>" & HTMLSafeLargo(Replace(sRACs, "|", vbCrLf)) & "</p>" & _
            "<p><strong>Correos:</strong><br>" & HTMLSafeLargo(Replace(sCorreosRACs, ";", vbCrLf)) & "</p></div>" & vbCrLf

    s = s & "  <div class='card'><h4>Autorizados</h4><p>" & HTMLSafeLargo(sAutorizados) & "</p></div>" & vbCrLf
    s = s & "</div>" & vbCrLf

    ConstruirSeccionDatosProyectoResponsables = s
End Function

Private Function GuardarInformeHTML_UTF8(ByVal p_Riesgo As riesgo, ByVal p_HTML As String, ByRef p_Error As String) As String
    Dim m_Ruta As String, m_Stream As Object
    On Error GoTo errores
    m_Ruta = Environ("TEMP") & "\HPS_Report_" & SanitizarNombreArchivo(p_Riesgo.CodigoRiesgo) & ".html"
    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2: m_Stream.Charset = "utf-8": m_Stream.Open
    m_Stream.WriteText p_HTML: m_Stream.SaveToFile m_Ruta, 2: m_Stream.Close
    GuardarInformeHTML_UTF8 = m_Ruta
    Exit Function
errores:
    p_Error = "Error UTF-8: " & Err.Description
End Function

' --- Utilidades de Seguridad y Formato ---
Private Function HTMLSafe(ByVal p_Texto As String) As String
    Dim m_T As String: m_T = Nz(p_Texto, "")
    m_T = Replace(m_T, "&", "&amp;"): m_T = Replace(m_T, "<", "&lt;"): m_T = Replace(m_T, ">", "&gt;"): m_T = Replace(m_T, """", "&quot;")
    HTMLSafe = m_T
End Function

Private Function HTMLSafeLargo(ByVal p_Texto As String) As String
    Dim m_T As String: m_T = HTMLSafe(p_Texto)
    m_T = Replace(m_T, vbCrLf, "<br>"): m_T = Replace(m_T, vbLf, "<br>")
    HTMLSafeLargo = m_T
End Function

Private Function FormatoFecha(ByVal p_Valor As Variant) As String
    If IsDate(p_Valor) Then FormatoFecha = Format$(CDate(p_Valor), "dd/mm/yyyy") Else FormatoFecha = "-"
End Function

Private Function SanitizarNombreArchivo(ByVal p_Texto As String) As String
    Dim m_T As String: m_T = p_Texto
    m_T = Replace(m_T, "/", "_"): m_T = Replace(m_T, "\", "_"): m_T = Replace(m_T, ":", "_")
    SanitizarNombreArchivo = m_T
End Function








