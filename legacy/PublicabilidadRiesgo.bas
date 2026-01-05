Attribute VB_Name = "PublicabilidadRiesgo"
Option Compare Database
Option Explicit

Public Enum EnumPublicabilidadCheckEstado
    Cumple = 1
    NoCumple = 2
    NoAplica = 3
End Enum

Public Enum EnumPublicabilidadVeredicto
    Publicable = 1
    NoPublicable = 2
    NoAplica = 3
End Enum

Public Type tPublicabilidadRiesgoDatos
    CodigoRiesgo As String
    Descripcion As String
    EsEdicionActiva As EnumSiNo
    Estado As EnumRiesgoEstado
    Priorizacion As String
    RequiereRiesgoDeBiblioteca As EnumSiNo
    RiesgoParaRetipificar As EnumSiNo
    FechaRechazoAceptacionPorCalidad As String
    FechaRechazoRetiroPorCalidad As String
    FechaMaterializado As String
    RiesgoAltoOMuyAlto As EnumSiNo
    TienePMs As EnumSiNo
    TodosPMFinalizados As EnumSiNo
    AlgunPMActivo As EnumSiNo
    AlgunPMSinAcciones As EnumSiNo
    AlgunPCSinAcciones As EnumSiNo
    TienePCs As EnumSiNo
    TodosPCFinalizados As EnumSiNo
    AlgunPCActivo As EnumSiNo
End Type

Public Function EvaluarPublicabilidadRiesgo( _
                                            ByRef p_Datos As tPublicabilidadRiesgoDatos, _
                                            ByRef p_Checks As Scripting.Dictionary, _
                                            ByRef p_Veredicto As EnumPublicabilidadVeredicto, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSiNo
    Dim m_Publicable As EnumSiNo
    Dim m_Index As Long
    Dim m_Estado As EnumRiesgoEstado
    Dim m_EnAceptacion As EnumSiNo
    Dim m_EnRetirada As EnumSiNo
    Dim m_AplicaBloque As Boolean
    Dim m_TextoDetalle As String

    On Error GoTo errores

    p_Error = ""
    m_Publicable = EnumSiNo.Sí
    p_Veredicto = EnumPublicabilidadVeredicto.Publicable
    m_Index = 1

    Set p_Checks = New Scripting.Dictionary
    p_Checks.CompareMode = TextCompare

    m_Estado = p_Datos.Estado
    m_EnAceptacion = RiesgoEnAceptacion(m_Estado)
    m_EnRetirada = RiesgoEnRetirada(m_Estado)

    If p_Datos.EsEdicionActiva = EnumSiNo.No Then
        p_Veredicto = EnumPublicabilidadVeredicto.NoAplica
        AgregarCheck p_Checks, m_Index, "edicion_activa", "Edicion activa", EnumPublicabilidadCheckEstado.NoAplica, "Edicion no activa, riesgo ya publicado"
        AgregarChecksNoAplica p_Checks, m_Index
        EvaluarPublicabilidadRiesgo = EnumSiNo.Sí
        Exit Function
    End If

    AgregarCheck p_Checks, m_Index, "edicion_activa", "Edicion activa", EnumPublicabilidadCheckEstado.Cumple

    If m_Estado = EnumRiesgoEstado.Aceptado Then
        AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigacion con acciones", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoAplica

        If IsNumeric(p_Datos.Priorizacion) Then
            AgregarCheck p_Checks, m_Index, "priorizacion", "Priorizacion establecida", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "priorizacion", "Priorizacion establecida", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If

        AgregarCheck p_Checks, m_Index, "aceptacion_calidad", "Aceptacion aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigacion activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigacion activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigacion definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoAplica

        If m_Publicable = EnumSiNo.No Then
            p_Veredicto = EnumPublicabilidadVeredicto.NoPublicable
        End If
        EvaluarPublicabilidadRiesgo = m_Publicable
        Exit Function
    End If

    If p_Datos.TienePMs = EnumSiNo.Sí Then
        If p_Datos.AlgunPMSinAcciones = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigacion con acciones", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigacion con acciones", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_con_acciones", "Plan de mitigacion con acciones", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If p_Datos.TienePCs = EnumSiNo.Sí Then
        If p_Datos.AlgunPCSinAcciones = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pc_con_acciones", "Plan de contingencia con acciones", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoAplica
    ElseIf m_Estado <> EnumRiesgoEstado.Incompleto Then
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.Cumple
    Else
        AgregarCheck p_Checks, m_Index, "datos_generales", "Datos generales cumplimentados", EnumPublicabilidadCheckEstado.NoCumple
        m_Publicable = EnumSiNo.No
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "priorizacion", "Priorizacion establecida", EnumPublicabilidadCheckEstado.NoAplica
    ElseIf IsNumeric(p_Datos.Priorizacion) Then
        AgregarCheck p_Checks, m_Index, "priorizacion", "Priorizacion establecida", EnumPublicabilidadCheckEstado.Cumple
    Else
        AgregarCheck p_Checks, m_Index, "priorizacion", "Priorizacion establecida", EnumPublicabilidadCheckEstado.NoCumple
        m_Publicable = EnumSiNo.No
    End If

    If m_Estado = EnumRiesgoEstado.Aceptado Then
        AgregarCheck p_Checks, m_Index, "aceptacion_calidad", "Aceptacion aprobada por calidad", EnumPublicabilidadCheckEstado.Cumple
    ElseIf m_EnAceptacion = EnumSiNo.Sí Then
        If p_Datos.FechaRechazoAceptacionPorCalidad <> "" Then
            m_TextoDetalle = "Rechazada por calidad"
        Else
            m_TextoDetalle = "Pendiente de evaluacion por calidad"
        End If
        AgregarCheck p_Checks, m_Index, "aceptacion_calidad", "Aceptacion aprobada por calidad", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
        m_Publicable = EnumSiNo.No
    Else
        AgregarCheck p_Checks, m_Index, "aceptacion_calidad", "Aceptacion aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_Estado = EnumRiesgoEstado.Retirado Then
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.Cumple
    ElseIf m_EnRetirada = EnumSiNo.Sí Then
        If p_Datos.FechaRechazoRetiroPorCalidad <> "" Then
            m_TextoDetalle = "Rechazada por calidad"
        Else
            m_TextoDetalle = "Pendiente de evaluacion por calidad"
        End If
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
        m_Publicable = EnumSiNo.No
    Else
        AgregarCheck p_Checks, m_Index, "retirada_calidad", "Retirada aprobada por calidad", EnumPublicabilidadCheckEstado.NoAplica
    End If

    m_AplicaBloque = (m_Estado <> EnumRiesgoEstado.Aceptado And m_Estado <> EnumRiesgoEstado.Retirado)

    If m_AplicaBloque And p_Datos.RequiereRiesgoDeBiblioteca = EnumSiNo.Sí Then
        If p_Datos.RiesgoParaRetipificar = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        Else
            AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.Cumple
        End If
    Else
        AgregarCheck p_Checks, m_Index, "retipificacion", "Riesgo retipificado", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.AlgunPMActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigacion activo (materializado)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigacion activo (materializado)", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If
        If p_Datos.AlgunPCActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_activo_materializado", "Plan de mitigacion activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_activo_materializado", "Plan de contingencia activo (materializado)", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And p_Datos.RiesgoAltoOMuyAlto = EnumSiNo.Sí And Not IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.AlgunPMActivo = EnumSiNo.Sí Then
            AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigacion activo (alto/muy alto)", EnumPublicabilidadCheckEstado.Cumple
        Else
            AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigacion activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoCumple
            m_Publicable = EnumSiNo.No
        End If
        If p_Datos.TienePCs = EnumSiNo.Sí And p_Datos.TodosPCFinalizados = EnumSiNo.No Then
            AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.Cumple
        Else
            m_TextoDetalle = ""
            If p_Datos.TienePCs = EnumSiNo.No Then
                m_TextoDetalle = "Sin planes definidos"
            ElseIf p_Datos.TodosPCFinalizados = EnumSiNo.Sí Then
                m_TextoDetalle = "Todos los planes finalizados"
            End If
            AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
            m_Publicable = EnumSiNo.No
        End If
    Else
        AgregarCheck p_Checks, m_Index, "pm_activo_alto", "Plan de mitigacion activo (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
        AgregarCheck p_Checks, m_Index, "pc_definido_alto", "Plan de contingencia definido (alto/muy alto)", EnumPublicabilidadCheckEstado.NoAplica
    End If

    If m_AplicaBloque And p_Datos.RiesgoAltoOMuyAlto = EnumSiNo.No And Not IsDate(p_Datos.FechaMaterializado) Then
        If p_Datos.TienePMs = EnumSiNo.Sí And p_Datos.TodosPMFinalizados = EnumSiNo.No Then
            AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigacion definido (bajo/medio)", EnumPublicabilidadCheckEstado.Cumple
        Else
            m_TextoDetalle = ""
            If p_Datos.TienePMs = EnumSiNo.No Then
                m_TextoDetalle = "Sin planes definidos"
            ElseIf p_Datos.TodosPMFinalizados = EnumSiNo.Sí Then
                m_TextoDetalle = "Todos los planes finalizados"
            End If
            AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigacion definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoCumple, m_TextoDetalle
            m_Publicable = EnumSiNo.No
        End If

    Else
        AgregarCheck p_Checks, m_Index, "pm_definido_bajo", "Plan de mitigacion definido (bajo/medio)", EnumPublicabilidadCheckEstado.NoAplica

    End If

    If p_Veredicto = EnumPublicabilidadVeredicto.Publicable Then
        If m_Publicable = EnumSiNo.No Then
            p_Veredicto = EnumPublicabilidadVeredicto.NoPublicable
        End If
    End If

    EvaluarPublicabilidadRiesgo = m_Publicable
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EvaluarPublicabilidadRiesgo ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ConstruirDatosPublicabilidadRiesgo( _
                                                    ByRef p_Riesgo As Riesgo, _
                                                    ByRef p_Datos As tPublicabilidadRiesgoDatos, _
                                                    Optional p_db As DAO.Database = Nothing, _
                                                    Optional ByRef p_Error As String _
                                                    ) As EnumSiNo
    Dim m_ColPMs As Scripting.Dictionary
    Dim m_ColPCs As Scripting.Dictionary

    On Error GoTo errores

    p_Error = ""
    If p_Riesgo Is Nothing Then
        p_Error = "Se ha de indicar el riesgo"
        Err.Raise 1000
    End If

    p_Datos.CodigoRiesgo = p_Riesgo.CodigoRiesgo
    p_Datos.Descripcion = p_Riesgo.DescripcionParaLista
    If p_Riesgo.Edicion Is Nothing Then
        p_Error = "No se ha podido determinar la edicion"
        Err.Raise 1000
    End If
    p_Datos.EsEdicionActiva = p_Riesgo.Edicion.EsActivo
    p_Datos.Estado = p_Riesgo.EstadoEnum
    p_Datos.Priorizacion = p_Riesgo.Priorizacion
    p_Datos.RequiereRiesgoDeBiblioteca = p_Riesgo.RequiereRiesgoDeBibliotecaCalculado
    p_Datos.RiesgoParaRetipificar = p_Riesgo.RiesgoParaRetipificar
    p_Datos.FechaRechazoAceptacionPorCalidad = p_Riesgo.FechaRechazoAceptacionPorCalidad
    p_Datos.FechaRechazoRetiroPorCalidad = p_Riesgo.FechaRechazoRetiroPorCalidad
    p_Datos.FechaMaterializado = p_Riesgo.FechaMaterializado
    p_Datos.RiesgoAltoOMuyAlto = p_Riesgo.RiesgoAltoOMuyAlto
    p_Datos.TienePMs = p_Riesgo.TienePMs
    p_Datos.TodosPMFinalizados = p_Riesgo.TodosPMFinalizados
    p_Datos.AlgunPMSinAcciones = p_Riesgo.AlgunPMSinAcciones
    p_Datos.AlgunPCSinAcciones = p_Riesgo.AlgunPCSinAcciones
    p_Datos.TienePCs = p_Riesgo.TienePCs
    p_Datos.TodosPCFinalizados = p_Riesgo.TodosPCFinalizados

    If p_db Is Nothing Then
        Set m_ColPMs = p_Riesgo.ColPMs
    Else
        Set m_ColPMs = Constructor.getPMs(p_Riesgo.IDRiesgo, , p_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If

    p_Datos.AlgunPMActivo = CalcularAlgunPlanActivo(m_ColPMs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        Set m_ColPCs = p_Riesgo.ColPCs
    Else
        Set m_ColPCs = Constructor.getPCs(p_Riesgo.IDRiesgo, , p_db, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If

    p_Datos.AlgunPCActivo = CalcularAlgunPlanActivo(m_ColPCs, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    ConstruirDatosPublicabilidadRiesgo = EnumSiNo.Sí
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo ConstruirDatosPublicabilidadRiesgo ha devuelto el error: " & Err.Description
    End If
    ConstruirDatosPublicabilidadRiesgo = EnumSiNo.No
End Function

Public Function TextoVeredictoPublicabilidad(ByVal p_Veredicto As EnumPublicabilidadVeredicto) As String

    Select Case p_Veredicto
        Case EnumPublicabilidadVeredicto.Publicable
            TextoVeredictoPublicabilidad = "Publicable"
        Case EnumPublicabilidadVeredicto.NoPublicable
            TextoVeredictoPublicabilidad = "No publicable"
        Case EnumPublicabilidadVeredicto.NoAplica
            TextoVeredictoPublicabilidad = "No aplica"
        Case Else
            TextoVeredictoPublicabilidad = "Desconocido"
    End Select

End Function

Public Function TextoEstadoPublicabilidad(ByVal p_Estado As EnumPublicabilidadCheckEstado) As String

    Select Case p_Estado
        Case EnumPublicabilidadCheckEstado.Cumple
            TextoEstadoPublicabilidad = "cumple"
        Case EnumPublicabilidadCheckEstado.NoCumple
            TextoEstadoPublicabilidad = "no_cumple"
        Case EnumPublicabilidadCheckEstado.NoAplica
            TextoEstadoPublicabilidad = "no_aplica"
        Case Else
            TextoEstadoPublicabilidad = "desconocido"
    End Select

End Function

Private Sub AgregarCheck( _
                        ByRef p_Checks As Scripting.Dictionary, _
                        ByRef p_Index As Long, _
                        ByVal p_Id As String, _
                        ByVal p_Texto As String, _
                        ByVal p_Estado As EnumPublicabilidadCheckEstado, _
                        Optional ByVal p_Detalle As String = "" _
                        )
    Dim m_Check As Scripting.Dictionary

    Set m_Check = New Scripting.Dictionary
    m_Check.CompareMode = TextCompare
    m_Check.Add "id", p_Id
    m_Check.Add "texto", p_Texto
    m_Check.Add "estado", p_Estado
    If p_Detalle <> "" Then
        m_Check.Add "detalle", p_Detalle
    End If

    p_Checks.Add CStr(p_Index), m_Check
    p_Index = p_Index + 1
End Sub

Private Sub AgregarChecksNoAplica( _
                                ByRef p_Checks As Scripting.Dictionary, _
                                ByRef p_Index As Long _
                                )
    Dim m_List As Variant
    Dim m_Item As Variant

    m_List = Array( _
        Array("datos_generales", "Datos generales cumplimentados"), _
        Array("priorizacion", "Priorizacion establecida"), _
        Array("aceptacion_calidad", "Aceptacion aprobada por calidad"), _
        Array("retirada_calidad", "Retirada aprobada por calidad"), _
        Array("pm_con_acciones", "Plan de mitigacion con acciones"), _
        Array("pc_con_acciones", "Plan de contingencia con acciones"), _
        Array("retipificacion", "Riesgo retipificado"), _
        Array("pm_activo_materializado", "Plan de mitigacion activo (materializado)"), _
        Array("pc_activo_materializado", "Plan de contingencia activo (materializado)"), _
        Array("pm_activo_alto", "Plan de mitigacion activo (alto/muy alto)"), _
        Array("pc_definido_alto", "Plan de contingencia definido (alto/muy alto)"), _
        Array("pm_definido_bajo", "Plan de mitigacion definido (bajo/medio)") _
    )

    For Each m_Item In m_List
        AgregarCheck p_Checks, p_Index, m_Item(0), m_Item(1), EnumPublicabilidadCheckEstado.NoAplica
    Next
End Sub

Private Function CalcularAlgunPlanActivo( _
                                        ByRef p_ColPlanes As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    Dim m_Id As Variant
    Dim m_Plan As Object
    Dim m_Acciones As Scripting.Dictionary

    On Error GoTo errores

    If p_ColPlanes Is Nothing Then
        CalcularAlgunPlanActivo = EnumSiNo.No
        Exit Function
    End If

    For Each m_Id In p_ColPlanes
        Set m_Plan = p_ColPlanes(m_Id)
        Set m_Acciones = m_Plan.colAcciones
        If Not m_Acciones Is Nothing Then
            If AlgunaAccionActiva(m_Acciones) = EnumSiNo.Sí Then
                CalcularAlgunPlanActivo = EnumSiNo.Sí
                Exit Function
            End If
        End If
        Set m_Plan = Nothing
    Next

    CalcularAlgunPlanActivo = EnumSiNo.No
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CalcularAlgunPlanActivo ha devuelto el error: " & Err.Description
    End If
End Function

Private Function AlgunaAccionActiva( _
                                    ByRef p_Acciones As Scripting.Dictionary _
                                    ) As EnumSiNo
    Dim m_Id As Variant
    Dim m_Accion As Object

    If p_Acciones Is Nothing Then
        AlgunaAccionActiva = EnumSiNo.No
        Exit Function
    End If

    For Each m_Id In p_Acciones
        Set m_Accion = p_Acciones(m_Id)
        If IsDate(m_Accion.FechaFinPrevista) And Not IsDate(m_Accion.FechaFinReal) Then
            AlgunaAccionActiva = EnumSiNo.Sí
            Exit Function
        End If
        Set m_Accion = Nothing
    Next

    AlgunaAccionActiva = EnumSiNo.No
End Function
