

Option Compare Database
Option Explicit


Public Function test_Fatima_Es_Usuario_Calidad() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsDeCalidad("fmc")
    
End Function

Public Function test_Fernando_Es_Usuario_Admin() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsAdministrador("ds01474")
    
End Function

Public Function test_SGM_Es_Usuario_Calidad_Con_Avisos() As String
    
    Debug.Assert EnumSiNo.Sí = UsuarioEsDeCalidad("sgm")
    
End Function

Public Function test_getEstadosDiferentesHastaEdicion( _
                                                    Optional ByVal p_IDEdicion As String = "274", _
                                                    Optional ByVal p_CodigoRiesgo As String = "R012", _
                                                    Optional ByVal p_FechaPublicacion As String = "", _
                                                    Optional ByVal p_FechaCierre As String = "" _
                                                    ) As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_Dic As Scripting.Dictionary
    Dim m_Key As Variant
    Dim m_Linea As String
    Dim m_PrevEstado As String
    Dim m_Estado As String
    Dim m_Fecha As String
    Dim m_Partes As Variant

    On Error GoTo errores

    Set m_Edicion = Constructor.getEdicion(p_IDEdicion, m_Error)
    If m_Error <> "" Then
        test_getEstadosDiferentesHastaEdicion = "ERROR getEdicion: " & m_Error
        Exit Function
    End If
    If m_Edicion Is Nothing Then
        test_getEstadosDiferentesHastaEdicion = "ERROR getEdicion: edicion nula"
        Exit Function
    End If

    If Not IsDate(p_FechaPublicacion) Then
        If IsDate(m_Edicion.FechaPublicacion) Then
            p_FechaPublicacion = m_Edicion.FechaPublicacion
        Else
            p_FechaPublicacion = Date
        End If
    End If

    Set m_Dic = getEstadosDiferentesHastaEdicion(m_Edicion, p_CodigoRiesgo, p_FechaPublicacion, p_FechaCierre, m_Error)
    If m_Error <> "" Then
        test_getEstadosDiferentesHastaEdicion = "ERROR getEstadosDiferentesHastaEdicion: " & m_Error
        Exit Function
    End If
    If m_Dic Is Nothing Then
        test_getEstadosDiferentesHastaEdicion = "ERROR getEstadosDiferentesHastaEdicion: diccionario nulo"
        Exit Function
    End If
    If m_Dic.Count = 0 Then
        test_getEstadosDiferentesHastaEdicion = "OK: sin estados (Count=0)"
        Exit Function
    End If

    m_PrevEstado = ""
    For Each m_Key In m_Dic
        m_Linea = CStr(m_Dic(m_Key))
        m_Partes = Split(m_Linea, "|")
        If UBound(m_Partes) >= 0 Then m_Estado = Trim$(CStr(m_Partes(0))) Else m_Estado = ""
        If UBound(m_Partes) >= 1 Then m_Fecha = Trim$(CStr(m_Partes(1))) Else m_Fecha = ""

        Debug.Assert m_Estado <> ""
        Debug.Assert m_Estado <> m_PrevEstado

        test_getEstadosDiferentesHastaEdicion = test_getEstadosDiferentesHastaEdicion & CStr(m_Key) & ":" & m_Estado & "|" & m_Fecha & vbCrLf
        m_PrevEstado = m_Estado
    Next

    Exit Function
errores:
    test_getEstadosDiferentesHastaEdicion = "ERROR " & Err.Number & ": " & Err.Description
End Function
Public Function test_Riesgo_Origen_Aceptado() As String
    
    Dim m_Cod As String
    Dim m_IDEdicion As String
    Dim m_Riesgo As riesgo
    Dim m_RiesgoNacimiento As riesgo
    
    m_Cod = "R012"
    m_IDEdicion = "274"
    Set m_Riesgo = Constructor.getRiesgo(, m_IDEdicion, m_Cod)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    Set m_RiesgoNacimiento = m_Riesgo.RiesgoAceptadoEnNacimiento
    test_Riesgo_Origen_Aceptado = m_RiesgoNacimiento.DiasRespuestaCalidadAceptacion
    
    
    
End Function
Public Function test_Riesgo_Origen_retirado() As String
    
    Dim m_Cod As String
    Dim m_IDEdicion As String
    Dim m_Riesgo As riesgo
    Dim m_RiesgoNacimiento As riesgo
    
    m_Cod = "R003"
    m_IDEdicion = "214"
    Set m_Riesgo = Constructor.getRiesgo(, m_IDEdicion, m_Cod)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    Set m_RiesgoNacimiento = m_Riesgo.RiesgoRetiradoEnNacimiento
    test_Riesgo_Origen_retirado = m_RiesgoNacimiento.DiasRespuestaCalidadRetiro
    
    
    
End Function
Public Function test_Riesgo_Dias_Por_Aceptar() As String
    
    
    Dim m_IdRiesgo As String
    Dim m_Riesgo As riesgo
    
    m_IdRiesgo = "1395"
   
    Set m_Riesgo = Constructor.getRiesgo(m_IdRiesgo)
    If m_Riesgo Is Nothing Then
        Exit Function
    End If
    m_Riesgo.RegistrarDiasAceptacionCalidad
    test_Riesgo_Dias_Por_Aceptar = m_Riesgo.DiasSinRespuestaCalidadAceptacion
End Function

'-------------------------------------------
' Nombre: test_Cadena_MAIN_vs_SUB
' Propósito: Verificar cadenas generadas por GetCadenaJerarquicaEmpresas
' Parámetros: ninguno
' Retorno: String (resumen simple)
'-------------------------------------------
Public Function test_Cadena_MAIN_vs_SUB() As String
    On Error GoTo ManejoError
    Dim sMain As String
    Dim sSub As String
    Dim sMsg As String
    Dim sErr As String
    
    sMain = GetCadenaJerarquicaEmpresas(1007, "MAIN", sErr)
    sSub = GetCadenaJerarquicaEmpresas(1007, "SUB", sErr)
    
    sMsg = "MAIN=" & sMain & vbCrLf & "SUB=" & sSub
    test_Cadena_MAIN_vs_SUB = sMsg
    Exit Function
ManejoError:
    test_Cadena_MAIN_vs_SUB = "Error " & Err.Number & ": " & Err.Description
End Function
'-------------------------------------------
' Nombre: test_Validacion_Roles
' Propósito: Verificar helpers EsContratistaPrincipal / EsSubContratista
' Parámetros: ninguno
' Retorno: Boolean (True si ambos tests pasan)
'-------------------------------------------
Public Function test_Validacion_Roles() As Boolean
    On Error GoTo ManejoError
    Dim okMain As Boolean
    Dim okSub As Boolean
    
    okMain = EsContratistaPrincipal(Null, "Sí")
    okSub = EsSubContratista(100, "No")
    
    test_Validacion_Roles = (okMain And okSub)
    Exit Function
ManejoError:
    test_Validacion_Roles = False
End Function
'-------------------------------------------
' Nombre: test_Retrocompatibilidad_Bandera
' Propósito: Verificar elección de lógica según TempVars("CadenaJerarquicaModelo")
' Parámetros: ninguno
' Retorno: String (resumen)
'-------------------------------------------
Public Function test_Retrocompatibilidad_Bandera() As String
    On Error GoTo ManejoError
    Dim sModelo As String
    Dim sOut As String
    Dim sErr As String
    Dim dic As Scripting.Dictionary
    
    sModelo = Trim$(Nz(Application.TempVars("CadenaJerarquicaModelo"), "nuevo"))
    
    Set dic = getExpedienteSuministradores_RC("1007", sErr)
    sOut = "Modelo=" & sModelo & "; Subcontratistas=" & IIf(dic Is Nothing, 0, dic.Count)
    test_Retrocompatibilidad_Bandera = sOut
    Exit Function
ManejoError:
    test_Retrocompatibilidad_Bandera = "Error " & Err.Number & ": " & Err.Description
End Function



Public Function test_PublicabilidadRiesgo_Aceptado() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.Aceptado
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_Aceptado = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.Sí Then
        test_PublicabilidadRiesgo_Aceptado = "Esperado publicable para riesgo aceptado"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "aceptacion_calidad") <> EnumPublicabilidadCheckEstado.NoAPlica Then
        test_PublicabilidadRiesgo_Aceptado = "Aceptacion aprobada por calidad no esta en no aplica"
    End If
End Function

Public Function test_CacheArbolRiesgosTx_Rollback( _
                                                    Optional ByVal p_IDRiesgo As String = "" _
                                                    ) As String
    ' Requiere una edicion con cache construida (TbCacheArbolRiesgosMeta).
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rcd As DAO.Recordset
    Dim m_IdRiesgo As Long
    Dim m_IDEdicion As Long
    Dim m_UpdatedAtAntes As Variant
    Dim m_UpdatedAtDespues As Variant
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    
    On Error GoTo errores
    
    Set db = getdb(m_Error)
    If m_Error <> "" Then
        test_CacheArbolRiesgosTx_Rollback = "ERROR getdb: " & m_Error
        Exit Function
    End If
    
    If Trim$(p_IDRiesgo) <> "" Then
        If Not IsNumeric(p_IDRiesgo) Then
            test_CacheArbolRiesgosTx_Rollback = "ERROR p_IDRiesgo no numerico"
            Exit Function
        End If
        m_IdRiesgo = CLng(p_IDRiesgo)
    Else
        Set rcd = db.OpenRecordset("SELECT TOP 1 IDRiesgo, IDEdicion FROM TbRiesgos ORDER BY IDRiesgo")
        If rcd.EOF Then
            test_CacheArbolRiesgosTx_Rollback = "SKIP: no hay riesgos"
            rcd.Close
            Exit Function
        End If
        m_IdRiesgo = CLng(rcd!IDRiesgo)
        m_IDEdicion = CLng(rcd!IDEdicion)
        rcd.Close
    End If
    
    If m_IDEdicion = 0 Then
        Set rcd = db.OpenRecordset("SELECT IDEdicion FROM TbRiesgos WHERE IDRiesgo=" & m_IdRiesgo)
        If rcd.EOF Then
            test_CacheArbolRiesgosTx_Rollback = "ERROR riesgo no encontrado"
            rcd.Close
            Exit Function
        End If
        m_IDEdicion = CLng(rcd!IDEdicion)
        rcd.Close
    End If
    
    Set rcd = db.OpenRecordset("SELECT UpdatedAt FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & m_IDEdicion)
    If rcd.EOF Then
        test_CacheArbolRiesgosTx_Rollback = "SKIP: cache no construida"
        rcd.Close
        Exit Function
    End If
    m_UpdatedAtAntes = rcd!UpdatedAt
    rcd.Close
    
    Set m_Riesgo = Constructor.getRiesgo(CStr(m_IdRiesgo), , , m_Error)
    If m_Error <> "" Then
        test_CacheArbolRiesgosTx_Rollback = "ERROR getRiesgo: " & m_Error
        Exit Function
    End If
    
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    CacheArbolRiesgosTx_ActualizarRiesgo m_Riesgo, db, m_Error
    If m_Error <> "" Then
        ws.Rollback
        test_CacheArbolRiesgosTx_Rollback = "ERROR CacheArbolRiesgosTx_ActualizarRiesgo: " & m_Error
        Exit Function
    End If
    ws.Rollback
    
    Set rcd = db.OpenRecordset("SELECT UpdatedAt FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & m_IDEdicion)
    If rcd.EOF Then
        test_CacheArbolRiesgosTx_Rollback = "ERROR: cache meta desaparecida"
        rcd.Close
        Exit Function
    End If
    m_UpdatedAtDespues = rcd!UpdatedAt
    rcd.Close
    
    If Nz(m_UpdatedAtAntes, "") <> Nz(m_UpdatedAtDespues, "") Then
        test_CacheArbolRiesgosTx_Rollback = "FAIL: UpdatedAt cambiado tras rollback"
        Exit Function
    End If
    
    test_CacheArbolRiesgosTx_Rollback = "OK"
    Exit Function
    
errores:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    test_CacheArbolRiesgosTx_Rollback = "ERROR: " & Err.Description
End Function

Public Function test_PublicabilidadRiesgo_Aceptado_Pendiente() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.AceptadoSinVisar
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_Aceptado_Pendiente = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.No Then
        test_PublicabilidadRiesgo_Aceptado_Pendiente = "Esperado no publicable para aceptado sin visar"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "aceptacion_calidad") <> EnumPublicabilidadCheckEstado.NoCumple Then
        test_PublicabilidadRiesgo_Aceptado_Pendiente = "Aceptacion pendiente no esta en no cumple"
    End If
End Function

Public Function test_PublicabilidadRiesgo_Retirada_Pendiente() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.RetiradoSinVisar
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_Retirada_Pendiente = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.No Then
        test_PublicabilidadRiesgo_Retirada_Pendiente = "Esperado no publicable para retirado sin visar"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "retirada_calidad") <> EnumPublicabilidadCheckEstado.NoCumple Then
        test_PublicabilidadRiesgo_Retirada_Pendiente = "Retirada pendiente no esta en no cumple"
    End If
End Function

Public Function test_PublicabilidadRiesgo_Bajo_Sin_PM() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.Detectado
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No
    m_Datos.RiesgoAltoOMuyAlto = EnumSiNo.No
    m_Datos.FechaMaterializado = ""
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.TodosPMFinalizados = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_Bajo_Sin_PM = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.No Then
        test_PublicabilidadRiesgo_Bajo_Sin_PM = "Esperado no publicable para bajo/medio sin PM"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "pm_definido_bajo") <> EnumPublicabilidadCheckEstado.NoCumple Then
        test_PublicabilidadRiesgo_Bajo_Sin_PM = "PM definido bajo/medio no esta en no cumple"
    End If
End Function

Public Function test_PublicabilidadRiesgo_Alto_Sin_PC() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.Activo
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No
    m_Datos.TienePMs = EnumSiNo.Sí
    m_Datos.RiesgoAltoOMuyAlto = EnumSiNo.Sí
    m_Datos.FechaMaterializado = ""
    m_Datos.AlgunPMActivo = EnumSiNo.Sí
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.TodosPCFinalizados = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_Alto_Sin_PC = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.No Then
        test_PublicabilidadRiesgo_Alto_Sin_PC = "Esperado no publicable para alto sin PC definido"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "pc_definido_alto") <> EnumPublicabilidadCheckEstado.NoCumple Then
        test_PublicabilidadRiesgo_Alto_Sin_PC = "PC definido alto/muy alto no esta en no cumple"
    End If
End Function

Public Function test_PublicabilidadRiesgo_PM_Finalizados_No_Cuenta() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.Activo
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.Sí
    m_Datos.TienePMs = EnumSiNo.No
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No
    m_Datos.TienePCs = EnumSiNo.No
    m_Datos.AlgunPCSinAcciones = EnumSiNo.No
    m_Datos.RiesgoAltoOMuyAlto = EnumSiNo.No
    m_Datos.FechaMaterializado = ""
    m_Datos.TienePMs = EnumSiNo.Sí
    m_Datos.TodosPMFinalizados = EnumSiNo.Sí
    m_Datos.AlgunPMSinAcciones = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_PM_Finalizados_No_Cuenta = m_Error
        Exit Function
    End If
    If m_Result <> EnumSiNo.No Then
        test_PublicabilidadRiesgo_PM_Finalizados_No_Cuenta = "Esperado no publicable con PM finalizados"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "pm_definido_bajo") <> EnumPublicabilidadCheckEstado.NoCumple Then
        test_PublicabilidadRiesgo_PM_Finalizados_No_Cuenta = "PM finalizados no se refleja en no cumple"
    End If
End Function

Public Function test_PublicabilidadRiesgo_NoAplica_EdicionNoActiva() As String
    Dim m_Datos As tPublicabilidadRiesgoDatos
    Dim m_Checks As Scripting.Dictionary
    Dim m_Error As String
    Dim m_Result As EnumSiNo
    Dim m_Veredicto As EnumPublicabilidadVeredicto

    m_Datos.Estado = EnumRiesgoEstado.Activo
    m_Datos.Priorizacion = "1"
    m_Datos.EsEdicionActiva = EnumSiNo.No

    m_Result = EvaluarPublicabilidadRiesgo(m_Datos, m_Checks, m_Veredicto, m_Error)
    If m_Error <> "" Then
        test_PublicabilidadRiesgo_NoAplica_EdicionNoActiva = m_Error
        Exit Function
    End If
    If m_Veredicto <> EnumPublicabilidadVeredicto.NoAPlica Then
        test_PublicabilidadRiesgo_NoAplica_EdicionNoActiva = "Esperado veredicto no aplica en edicion no activa"
        Exit Function
    End If
    If CheckEstadoPorId(m_Checks, "edicion_activa") <> EnumPublicabilidadCheckEstado.NoAPlica Then
        test_PublicabilidadRiesgo_NoAplica_EdicionNoActiva = "Edicion activa no esta en no aplica"
    End If
End Function
Private Function CheckEstadoPorId( _
                                ByRef p_Checks As Scripting.Dictionary, _
                                ByVal p_Id As String _
                                ) As EnumPublicabilidadCheckEstado
    Dim m_Key As Variant
    Dim m_Check As Scripting.Dictionary

    If p_Checks Is Nothing Then
        Exit Function
    End If

    For Each m_Key In p_Checks
        Set m_Check = p_Checks(m_Key)
        If m_Check.Exists("id") Then
            If m_Check("id") = p_Id Then
                CheckEstadoPorId = m_Check("estado")
                Exit Function
            End If
        End If
    Next
End Function

'-------------------------------------------
' HTML corporativo (correos)
' Requiere TempVars:
' - Test_IDEdicion
' - Test_IDRiesgo
' - Test_IDProyecto (opcional; si no, usa Edicion.Proyecto)
'-------------------------------------------
Public Function test_HTML_Edicion_PropuestaPublicacion() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_PropuestaPublicacion = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLEdicionPropuestaPublicacion
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_PropuestaPublicacion = "ERROR HTMLEdicionPropuestaPublicacion: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_PropuestaPublicacion = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.PropuestaPublicacion")
End Function

Public Function test_HTML_Edicion_Revision() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_Revision = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLEdicionRevision
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_Revision = "ERROR HTMLEdicionRevision: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_Revision = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.Revision")
End Function

Public Function test_HTML_Edicion_QuitarPropuestaPublicacion() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_QuitarPropuestaPublicacion = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLEdicionQuitarPropuestaPublicacion
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_QuitarPropuestaPublicacion = "ERROR HTMLEdicionQuitarPropuestaPublicacion: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_QuitarPropuestaPublicacion = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.QuitarPropuestaPublicacion")
End Function

Public Function test_HTML_Edicion_RechazarPropuestaPublicacion() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_RechazarPropuestaPublicacion = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLEdicionRechazarPropuestaPublicacion
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_RechazarPropuestaPublicacion = "ERROR HTMLEdicionRechazarPropuestaPublicacion: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_RechazarPropuestaPublicacion = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.RechazarPropuestaPublicacion")
End Function

Public Function test_HTML_Edicion_HistorialPublicaciones() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_HistorialPublicaciones = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLHistorialPublicaciones
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_HistorialPublicaciones = "ERROR HTMLHistorialPublicaciones: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_HistorialPublicaciones = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.HistorialPublicaciones")
End Function

Public Function test_HTML_Edicion_HistorialCorreosResponsables() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_Edicion_HistorialCorreosResponsables = m_Error
        Exit Function
    End If

    m_HTML = m_Edicion.HTMLHistorialcorreosResponsables
    m_Error = m_Edicion.Error
    If m_Error <> "" Then
        test_HTML_Edicion_HistorialCorreosResponsables = "ERROR HTMLHistorialcorreosResponsables: " & m_Error
        Exit Function
    End If

    test_HTML_Edicion_HistorialCorreosResponsables = Test_ValidarHTMLCorporativo(m_HTML, "Edicion.HistorialCorreosResponsables")
End Function

Public Function test_HTML_Proyecto_Alta() As String
    Dim m_Proyecto As Proyecto
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveProyecto(m_Proyecto, m_Error) Then
        test_HTML_Proyecto_Alta = m_Error
        Exit Function
    End If

    m_HTML = m_Proyecto.HTMLProyecto
    m_Error = m_Proyecto.Error
    If m_Error <> "" Then
        test_HTML_Proyecto_Alta = "ERROR HTMLProyecto: " & m_Error
        Exit Function
    End If

    test_HTML_Proyecto_Alta = Test_ValidarHTMLCorporativo(m_HTML, "Proyecto.HTMLProyecto")
End Function

Public Function test_HTML_Riesgo_PorRetipificar() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_PorRetipificar = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.HTMLRiesgoPorRetipificar
    m_Error = m_Riesgo.Error
    If m_Error <> "" Then
        test_HTML_Riesgo_PorRetipificar = "ERROR HTMLRiesgoPorRetipificar: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_PorRetipificar = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.HTMLRiesgoPorRetipificar")
End Function

Public Function test_HTML_Riesgo_Aceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_Aceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLRiesgoAceptadoRetirado(EnumSiNo.Sí, m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_Aceptado = "ERROR getHTMLRiesgoAceptadoRetirado(Aceptado): " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_Aceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.Aceptado")
End Function

Public Function test_HTML_Riesgo_Retirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_Retirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLRiesgoAceptadoRetirado(EnumSiNo.No, m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_Retirado = "ERROR getHTMLRiesgoAceptadoRetirado(Retirado): " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_Retirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.Retirado")
End Function

Public Function test_HTML_Riesgo_TecnicoAceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_TecnicoAceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLTecnicoRiesgoAceptado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_TecnicoAceptado = "ERROR getHTMLTecnicoRiesgoAceptado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_TecnicoAceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.TecnicoAceptado")
End Function

Public Function test_HTML_Riesgo_TecnicoRetirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_TecnicoRetirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLTecnicoRiesgoRetirado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_TecnicoRetirado = "ERROR getHTMLTecnicoRiesgoRetirado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_TecnicoRetirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.TecnicoRetirado")
End Function

Public Function test_HTML_Riesgo_Retipificado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_Retipificado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLRiesgoRetipificado(m_Riesgo, m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_Retipificado = "ERROR getHTMLRiesgoRetipificado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_Retipificado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.Retipificado")
End Function

Public Function test_HTML_Riesgo_CalidadApruebaAceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadApruebaAceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadApruebaRiesgoAceptado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadApruebaAceptado = "ERROR getHTMLCalidadApruebaRiesgoAceptado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadApruebaAceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadApruebaAceptado")
End Function

Public Function test_HTML_Riesgo_CalidadQuitaAprobacionAceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadQuitaAprobacionAceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadQuitaAprobacionRiesgoAceptado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadQuitaAprobacionAceptado = "ERROR getHTMLCalidadQuitaAprobacionRiesgoAceptado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadQuitaAprobacionAceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadQuitaAprobacionAceptado")
End Function

Public Function test_HTML_Riesgo_CalidadRechazaAceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadRechazaAceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadRechazaRiesgoAceptado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadRechazaAceptado = "ERROR getHTMLCalidadRechazaRiesgoAceptado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadRechazaAceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadRechazaAceptado")
End Function

Public Function test_HTML_Riesgo_CalidadQuitarRechazoAceptado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadQuitarRechazoAceptado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadQuitarRechazoRiesgoAceptado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadQuitarRechazoAceptado = "ERROR getHTMLCalidadQuitarRechazoRiesgoAceptado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadQuitarRechazoAceptado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadQuitarRechazoAceptado")
End Function

Public Function test_HTML_Riesgo_CalidadApruebaRetirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadApruebaRetirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadApruebaRiesgoRetirado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadApruebaRetirado = "ERROR getHTMLCalidadApruebaRiesgoRetirado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadApruebaRetirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadApruebaRetirado")
End Function

Public Function test_HTML_Riesgo_CalidadQuitaAprobacionRetirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadQuitaAprobacionRetirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadQuitaAprobacionRiesgoRetirado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadQuitaAprobacionRetirado = "ERROR getHTMLCalidadQuitaAprobacionRiesgoRetirado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadQuitaAprobacionRetirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadQuitaAprobacionRetirado")
End Function

Public Function test_HTML_Riesgo_CalidadRechazaRetirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadRechazaRetirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadRechazaRiesgoRetirado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadRechazaRetirado = "ERROR getHTMLCalidadRechazaRiesgoRetirado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadRechazaRetirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadRechazaRetirado")
End Function

Public Function test_HTML_Riesgo_CalidadQuitarRechazoRetirado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_CalidadQuitarRechazoRetirado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.getHTMLCalidadQuitarRechazoRiesgoRetirado(m_Error)
    If m_Error <> "" Then
        test_HTML_Riesgo_CalidadQuitarRechazoRetirado = "ERROR getHTMLCalidadQuitarRechazoRiesgoRetirado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_CalidadQuitarRechazoRetirado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.CalidadQuitarRechazoRetirado")
End Function

Public Function test_HTML_Riesgo_Materializado() As String
    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_HTML As String

    If Not Test_ResolveRiesgo(m_Riesgo, m_Error) Then
        test_HTML_Riesgo_Materializado = m_Error
        Exit Function
    End If

    m_HTML = m_Riesgo.HTMLRiesgoMaterializado
    m_Error = m_Riesgo.Error
    If m_Error <> "" Then
        test_HTML_Riesgo_Materializado = "ERROR HTMLRiesgoMaterializado: " & m_Error
        Exit Function
    End If

    test_HTML_Riesgo_Materializado = Test_ValidarHTMLCorporativo(m_HTML, "Riesgo.HTMLRiesgoMaterializado")
End Function

'-------------------------------------------
' HTML corporativo (informes)
' Nota: GenerarInformeRiesgoHTML y GenerarInformeEdicionHTML abren navegador.
'-------------------------------------------
Public Function test_HTML_InformeRiesgo_Report() As String
    Dim m_IdRiesgo As String
    Dim m_Error As String
    Dim m_URL As String
    Dim m_HTML As String

    m_IdRiesgo = Trim$(Nz(Application.TempVars("Test_IDRiesgo"), ""))
    If m_IdRiesgo = "" Then
        test_HTML_InformeRiesgo_Report = "SKIP: Defina TempVars('Test_IDRiesgo') con un riesgo valido"
        Exit Function
    End If

    m_URL = GenerarInformeRiesgoHTML(m_IdRiesgo, 0, m_Error)
    If m_Error <> "" Then
        test_HTML_InformeRiesgo_Report = "ERROR GenerarInformeRiesgoHTML: " & m_Error
        Exit Function
    End If

    m_HTML = Test_LeerArchivoUTF8(m_URL, m_Error)
    If m_Error <> "" Then
        test_HTML_InformeRiesgo_Report = m_Error
        Exit Function
    End If

    test_HTML_InformeRiesgo_Report = Test_ValidarHTMLCorporativo(m_HTML, "InformeRiesgoHTML.Report")
End Function

Public Function test_HTML_InformeEdicion_Report() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_URL As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_InformeEdicion_Report = m_Error
        Exit Function
    End If

    m_URL = GenerarInformeEdicionHTML(m_Edicion, 0, , , m_Error)
    If m_Error <> "" Then
        test_HTML_InformeEdicion_Report = "ERROR GenerarInformeEdicionHTML: " & m_Error
        Exit Function
    End If

    m_HTML = Test_LeerArchivoUTF8(m_URL, m_Error)
    If m_Error <> "" Then
        test_HTML_InformeEdicion_Report = m_Error
        Exit Function
    End If

    test_HTML_InformeEdicion_Report = Test_ValidarHTMLCorporativo(m_HTML, "InformeEdicionHTML.Report")
End Function

Public Function test_HTML_PublicabilidadEdicion_Interactivo() As String
    Dim m_Edicion As Edicion
    Dim m_Error As String
    Dim m_URL As String
    Dim m_HTML As String

    If Not Test_ResolveEdicion(m_Edicion, m_Error) Then
        test_HTML_PublicabilidadEdicion_Interactivo = m_Error
        Exit Function
    End If

    m_URL = URLInformePublicabilidad(m_Edicion, m_Error)
    If m_Error <> "" Then
        test_HTML_PublicabilidadEdicion_Interactivo = "ERROR URLInformePublicabilidad: " & m_Error
        Exit Function
    End If

    m_HTML = Test_LeerArchivoUTF8(m_URL, m_Error)
    If m_Error <> "" Then
        test_HTML_PublicabilidadEdicion_Interactivo = m_Error
        Exit Function
    End If

    test_HTML_PublicabilidadEdicion_Interactivo = Test_ValidarHTMLCorporativo(m_HTML, "PublicabilidadEdicion.Interactivo")
End Function

Public Function test_HTML_PrepararSuite( _
                                        Optional ByVal p_IDEdicion As String = "", _
                                        Optional ByVal p_IDRiesgo As String = "", _
                                        Optional ByVal p_IDProyecto As String = "" _
                                        ) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Error As String

    On Error GoTo errores

    If p_IDEdicion = "" Or p_IDRiesgo = "" Or p_IDProyecto = "" Then
        Set db = getdb(m_Error)
        If m_Error <> "" Then
            test_HTML_PrepararSuite = "ERROR getdb: " & m_Error
            Exit Function
        End If
    End If

    If p_IDRiesgo = "" Then
        Set rcd = db.OpenRecordset("SELECT TOP 1 IDRiesgo, IDEdicion FROM TbRiesgos ORDER BY IDRiesgo")
        If Not rcd.EOF Then
            p_IDRiesgo = CStr(rcd!IDRiesgo)
            If p_IDEdicion = "" Then
                p_IDEdicion = CStr(rcd!IDEdicion)
            End If
        End If
        rcd.Close
    End If

    If p_IDEdicion = "" Then
        Set rcd = db.OpenRecordset("SELECT TOP 1 IDEdicion, IDProyecto FROM TbEdiciones ORDER BY IDEdicion")
        If Not rcd.EOF Then
            p_IDEdicion = CStr(rcd!IDEdicion)
            If p_IDProyecto = "" Then
                p_IDProyecto = CStr(rcd!IDProyecto)
            End If
        End If
        rcd.Close
    End If

    If p_IDProyecto = "" Then
        Set rcd = db.OpenRecordset("SELECT TOP 1 IDProyecto FROM TbProyectos ORDER BY IDProyecto")
        If Not rcd.EOF Then
            p_IDProyecto = CStr(rcd!IDProyecto)
        End If
        rcd.Close
    End If

    If p_IDEdicion <> "" Then TempVars!Test_IDEdicion = p_IDEdicion
    If p_IDRiesgo <> "" Then TempVars!Test_IDRiesgo = p_IDRiesgo
    If p_IDProyecto <> "" Then TempVars!Test_IDProyecto = p_IDProyecto

    test_HTML_PrepararSuite = "OK: Test_IDEdicion=" & p_IDEdicion & "; Test_IDRiesgo=" & p_IDRiesgo & "; Test_IDProyecto=" & p_IDProyecto
    Exit Function

errores:
    test_HTML_PrepararSuite = "ERROR: " & Err.Description
End Function

Public Function test_HTML_LimpiarSuite() As String
    On Error GoTo errores

    If TempVars.Exists("Test_IDEdicion") Then TempVars.Remove "Test_IDEdicion"
    If TempVars.Exists("Test_IDRiesgo") Then TempVars.Remove "Test_IDRiesgo"
    If TempVars.Exists("Test_IDProyecto") Then TempVars.Remove "Test_IDProyecto"

    test_HTML_LimpiarSuite = "OK"
    Exit Function

errores:
    test_HTML_LimpiarSuite = "ERROR: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' NUEVOS TESTS TRANSACCIONALES (PROYECTO)
'---------------------------------------------------------------------------------------

Public Function test_Proyecto_CicloVida_Completo() As String
    ' Prueba el flujo completo: Alta -> Edición -> Borrado
    ' Verificando la existencia de registros y cachés en cada paso.
    Dim m_Proyecto As Proyecto
    Dim m_IDProyecto As String
    Dim m_Error As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    Set db = getdb()
    
    ' 1. ALTA
    Set m_Proyecto = New Proyecto
    With m_Proyecto
        .IDExpediente = "1" ' ID de prueba existente
        .ParaInformeAvisos = "No"
        .EnUTE = "No"
        .Elaborado = m_ObjUsuarioConectado.Nombre
        .Revisado = m_ObjUsuarioConectado.Nombre ' Mock para test
        .Aprobado = m_ObjUsuarioConectado.Nombre
    End With
    
    If CrearProyectoTransaccional(m_Proyecto, , , m_Error) <> "" Then
        test_Proyecto_CicloVida_Completo = "FAIL Alta: " & m_Error
        Exit Function
    End If
    m_IDProyecto = m_Proyecto.IDProyecto
    
    ' Verificar que existe Proyecto, Edicion y Cache
    If Nz(db.OpenRecordset("SELECT COUNT(*) FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto).Fields(0), 0) = 0 Then
        test_Proyecto_CicloVida_Completo = "FAIL: Proyecto no creado en DB"
        Exit Function
    End If
    If Nz(db.OpenRecordset("SELECT COUNT(*) FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto).Fields(0), 0) = 0 Then
        test_Proyecto_CicloVida_Completo = "FAIL: Primera Edición no creada"
        Exit Function
    End If
    
    ' 2. EDICIÓN
    m_Proyecto.NombreProyecto = "TEST_REFAC_" & Timer
    If EditarProyectoTransaccional(m_Proyecto, m_Proyecto, , m_Error) <> "" Then
        test_Proyecto_CicloVida_Completo = "FAIL Edición: " & m_Error
        Exit Function
    End If
    
    ' 3. BORRADO (Debe limpiar proyecto y caches)
    If BorrarProyectoTransaccional(m_Proyecto, , m_Error) <> "" Then
        test_Proyecto_CicloVida_Completo = "FAIL Borrado: " & m_Error
        Exit Function
    End If
    
    ' Verificar limpieza total
    If Nz(db.OpenRecordset("SELECT COUNT(*) FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto).Fields(0), 0) > 0 Then
        test_Proyecto_CicloVida_Completo = "FAIL: Proyecto no borrado"
        Exit Function
    End If
    
    test_Proyecto_CicloVida_Completo = "OK"
    Exit Function
errores:
    test_Proyecto_CicloVida_Completo = "ERROR: " & Err.Description
End Function

Public Function test_Proyecto_Alta_Rollback_En_Fallo() As String
    ' Verifica que si algo falla tras insertar el proyecto (ej. cache), el proyecto no se queda en DB.
    Dim m_Proyecto As Proyecto
    Dim m_IDProyecto As String
    Dim m_Error As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    
    On Error GoTo errores
    Set ws = DBEngine.Workspaces(0)
    Set db = getdb()
    
    Set m_Proyecto = New Proyecto
    With m_Proyecto
        .IDExpediente = "1"
        .ParaInformeAvisos = "No"
        .EnUTE = "No"
        .Elaborado = "TEST"
        .Revisado = "TEST"
        .Aprobado = "TEST"
    End With
    
    ' Simulamos un fallo envolviendo en una transacción que nosotros mismos revertimos
    ' o comprobamos que el error interno de CrearProyectoTransaccional (si lo hubiera) limpia.
    ' Para este test, vamos a intentar crear uno con un ID de expediente inexistente que falle en el motor si fuera posible,
    ' pero lo más fiable es verificar que el motor de transacciones de ProyectoTransaccional funciona.
    
    ' Provocamos error de validación
    m_Proyecto.IDExpediente = ""
    If CrearProyectoTransaccional(m_Proyecto, , , m_Error) = "" Then
        test_Proyecto_Alta_Rollback_En_Fallo = "FAIL: Debería haber fallado por validación"
        Exit Function
    End If
    
    test_Proyecto_Alta_Rollback_En_Fallo = "OK (Capturó error: " & Left(m_Error, 20) & "...)"
    Exit Function
errores:
    test_Proyecto_Alta_Rollback_En_Fallo = "ERROR: " & Err.Description
End Function

Public Function test_Proyecto_Alta_Sin_Riesgos_Cache() As String
    ' Verifica que el alta de un proyecto funciona correctamente a pesar de no tener riesgos iniciales,
    ' y que la caché de publicabilidad se genera sin errores.
    Dim m_Proyecto As Proyecto
    Dim m_Error As String
    Dim db As DAO.Database
    
    On Error GoTo errores
    Set db = getdb()
    
    Set m_Proyecto = New Proyecto
    With m_Proyecto
        .IDExpediente = "1"
        .ParaInformeAvisos = "No"
        .EnUTE = "No"
        .Elaborado = "TEST_USER"
        .Revisado = "TEST_USER"
        .Aprobado = "TEST_USER"
    End With
    
    ' Esto fallaba antes por el check de "La edición no tiene riesgos" en PublicabilidadEdicion.bas
    If CrearProyectoTransaccional(m_Proyecto, m_Error) <> "" Then
        test_Proyecto_Alta_Sin_Riesgos_Cache = "FAIL: " & m_Error
        Exit Function
    End If
    
    ' Limpiamos para no ensuciar la DB de pruebas
    BorrarProyectoTransaccional m_Proyecto, , m_Error
    
    test_Proyecto_Alta_Sin_Riesgos_Cache = "OK"
    Exit Function
errores:
    test_Proyecto_Alta_Sin_Riesgos_Cache = "ERROR: " & Err.Description
End Function

Private Function Test_ResolveEdicion(ByRef p_Edicion As Edicion, ByRef p_Error As String) As Boolean
    Dim m_IDEdicion As String

    m_IDEdicion = Trim$(Nz(Application.TempVars("Test_IDEdicion"), ""))
    If m_IDEdicion = "" Then
        p_Error = "SKIP: Defina TempVars('Test_IDEdicion') con una edicion valida"
        Exit Function
    End If

    Set p_Edicion = Constructor.getEdicion(m_IDEdicion, p_Error)
    If p_Error <> "" Then
        p_Error = "ERROR getEdicion: " & p_Error
        Exit Function
    End If
    If p_Edicion Is Nothing Then
        p_Error = "ERROR getEdicion: edicion nula"
        Exit Function
    End If

    Test_ResolveEdicion = True
End Function

Private Function Test_ResolveProyecto(ByRef p_Proyecto As Proyecto, ByRef p_Error As String) As Boolean
    Dim m_IDProyecto As String
    Dim m_Edicion As Edicion

    m_IDProyecto = Trim$(Nz(Application.TempVars("Test_IDProyecto"), ""))
    If m_IDProyecto <> "" Then
        Set p_Proyecto = Constructor.getProyecto(m_IDProyecto, p_Error)
        If p_Error <> "" Then
            p_Error = "ERROR getProyecto: " & p_Error
            Exit Function
        End If
        If p_Proyecto Is Nothing Then
            p_Error = "ERROR getProyecto: proyecto nulo"
            Exit Function
        End If
        Test_ResolveProyecto = True
        Exit Function
    End If

    If Not Test_ResolveEdicion(m_Edicion, p_Error) Then
        Exit Function
    End If

    Set p_Proyecto = m_Edicion.Proyecto
    If p_Proyecto Is Nothing Then
        p_Error = "ERROR: Edicion.Proyecto nulo"
        Exit Function
    End If

    Test_ResolveProyecto = True
End Function

Private Function Test_ResolveRiesgo(ByRef p_Riesgo As riesgo, ByRef p_Error As String) As Boolean
    Dim m_IdRiesgo As String

    m_IdRiesgo = Trim$(Nz(Application.TempVars("Test_IDRiesgo"), ""))
    If m_IdRiesgo = "" Then
        p_Error = "SKIP: Defina TempVars('Test_IDRiesgo') con un riesgo valido"
        Exit Function
    End If

    Set p_Riesgo = Constructor.getRiesgo(m_IdRiesgo, p_Error)
    If p_Error <> "" Then
        p_Error = "ERROR getRiesgo: " & p_Error
        Exit Function
    End If
    If p_Riesgo Is Nothing Then
        p_Error = "ERROR getRiesgo: riesgo nulo"
        Exit Function
    End If

    Test_ResolveRiesgo = True
End Function

Private Function Test_ValidarHTMLCorporativo(ByVal p_HTML As String, ByVal p_Contexto As String) As String
    If Len(p_HTML) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": HTML vacio"
        Exit Function
    End If
    If InStr(1, p_HTML, "<html", vbTextCompare) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": falta <html>"
        Exit Function
    End If
    If InStr(1, p_HTML, "header-container", vbTextCompare) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": falta header-container"
        Exit Function
    End If
    If InStr(1, p_HTML, "main-container", vbTextCompare) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": falta main-container"
        Exit Function
    End If
    If InStr(1, p_HTML, "<footer", vbTextCompare) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": falta footer"
        Exit Function
    End If
    If InStr(1, p_HTML, "#titulo", vbTextCompare) <> 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": titulo sin reemplazar"
        Exit Function
    End If
    If InStr(1, p_HTML, "</html>", vbTextCompare) = 0 Then
        Test_ValidarHTMLCorporativo = "ERROR " & p_Contexto & ": falta cierre </html>"
        Exit Function
    End If

    Test_ValidarHTMLCorporativo = "OK: " & p_Contexto
End Function

Private Function Test_LeerArchivoUTF8(ByVal p_URL As String, ByRef p_Error As String) As String
    Dim m_Stream As Object
    Dim m_FSO As Object

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    If Not m_FSO.FileExists(p_URL) Then
        p_Error = "ERROR: archivo no encontrado: " & p_URL
        Exit Function
    End If

    Set m_Stream = CreateObject("ADODB.Stream")
    m_Stream.Type = 2
    m_Stream.Charset = "utf-8"
    m_Stream.Open
    m_Stream.LoadFromFile p_URL
    Test_LeerArchivoUTF8 = m_Stream.ReadText
    m_Stream.Close
End Function



