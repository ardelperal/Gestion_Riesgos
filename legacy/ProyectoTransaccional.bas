Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ProyectoTransaccional
' Purpose   : Centraliza toda la lógica de persistencia de la clase Proyecto,
'             garantizando atomicidad absoluta mediante transacciones DAO y
'             manteniendo sincronizados todos los sistemas de caché.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : CrearProyectoTransaccional
' Purpose   : Realiza el Alta de un Proyecto.
'---------------------------------------------------------------------------------------
Public Function CrearProyectoTransaccional( _
                            p_Proyecto As Proyecto, _
                            Optional ByVal p_EsTecnico As EnumSiNo = EnumSiNo.No, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_FechaRef As Date
    Dim m_objEdicion As Edicion
    Dim m_motivosNoOK As String
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If

    m_EnTransaccion = False
    p_Error = ""
    
    ' 1. CÁLCULO DE PROPIEDADES
    With p_Proyecto
        If .Proyecto = "" Then .Proyecto = .ProyectoCalculado
        If .Juridica = "" Then .Juridica = .JuridicaCalculada
        If .NombreProyecto = "" Then .NombreProyecto = .NombreProyectoCalculado
        If .Cliente = "" Then .Cliente = .ClienteCalculado
        .FechaPrevistaCierre = .FechaPrevistaCierreCalculada
        If .CodigoDocumento = "" Then .CodigoDocumento = .CodigoDocumentoCalculado
        If .FechaFirmaContrato = "" Then .FechaFirmaContrato = .FechaFirmaContratoCalculada
        If .NombreUsuarioCalidad = "" Then .NombreUsuarioCalidad = .NombreUsuarioCalidadCalculado
        If .EnUTE <> "Sí" And .EnUTE <> "No" Then .EnUTE = .EnUTECalculado
        If .FechaMaxProximaPublicacion = "" Then .FechaMaxProximaPublicacion = .FechaMaxProximaPublicacionCalculada
        If .Ordinal = "" Then .Ordinal = .OrdinalCalculado
        .NombreParaNodo = .NombreParaNodoCalculado
    End With

    ' 2. VALIDACIÓN
    m_motivosNoOK = p_Proyecto.MotivoNoOK(p_ObjProyectoAlInicio:=Nothing, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_motivosNoOK <> "" Then
        p_Error = m_motivosNoOK
        Err.Raise 1000
    End If
    
    ' 3. INICIO DE TRANSACCIÓN
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If
    m_FechaRef = Date

    ' 4. INSERTAR EN TbProyectos
    If p_Proyecto.IDProyecto = "" Then p_Proyecto.IDProyecto = p_Proyecto.IDProyectoCalculado
    
    m_SQL = "TbProyectos"
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
        .Fields("IDProyecto") = p_Proyecto.IDProyecto
        If p_Proyecto.IDExpediente <> "" Then .Fields("IDExpediente") = p_Proyecto.IDExpediente
        If p_Proyecto.Proyecto <> "" Then .Fields("Proyecto") = p_Proyecto.Proyecto
        If p_Proyecto.Juridica <> "" Then .Fields("Juridica") = p_Proyecto.Juridica
        If p_Proyecto.NombreProyecto <> "" Then .Fields("NombreProyecto") = p_Proyecto.NombreProyecto
        If p_Proyecto.Cliente <> "" Then .Fields("Cliente") = p_Proyecto.Cliente
        If IsDate(p_Proyecto.FechaPrevistaCierre) Then .Fields("FechaPrevistaCierre") = p_Proyecto.FechaPrevistaCierre
        If IsDate(p_Proyecto.FechaCierre) Then .Fields("FechaCierre") = p_Proyecto.FechaCierre
        
        p_Proyecto.fechaRegistroInicial = CStr(m_FechaRef)
        .Fields("FechaRegistroInicial") = p_Proyecto.fechaRegistroInicial
        
        If p_Proyecto.Elaborado <> "" Then .Fields("Elaborado") = p_Proyecto.Elaborado
        If p_Proyecto.Revisado <> "" Then .Fields("Revisado") = p_Proyecto.Revisado
        If p_Proyecto.Aprobado <> "" Then .Fields("Aprobado") = p_Proyecto.Aprobado
        If p_Proyecto.CodigoDocumento <> "" Then .Fields("CodigoDocumento") = p_Proyecto.CodigoDocumento
        If p_Proyecto.ParaInformeAvisos <> "" Then .Fields("ParaInformeAvisos") = p_Proyecto.ParaInformeAvisos
        If IsDate(p_Proyecto.FechaFirmaContrato) Then .Fields("FechaFirmaContrato") = p_Proyecto.FechaFirmaContrato
        If p_Proyecto.NombreUsuarioCalidad <> "" Then .Fields("NombreUsuarioCalidad") = p_Proyecto.NombreUsuarioCalidad
        If p_Proyecto.EnUTE = "Sí" Or p_Proyecto.EnUTE = "No" Then .Fields("EnUTE") = p_Proyecto.EnUTE
        
        p_Proyecto.FechaMaxProximaPublicacion = p_Proyecto.FechaMaxProximaPublicacionCalculada
        If IsDate(p_Proyecto.FechaMaxProximaPublicacion) Then .Fields("FechaMaxProximaPublicacion") = p_Proyecto.FechaMaxProximaPublicacion
        
        p_Proyecto.RequiereRiesgoDeBiblioteca = "Sí"
        .Fields("RequiereRiesgoDeBiblioteca") = p_Proyecto.RequiereRiesgoDeBiblioteca
        
        If p_Proyecto.CorreoRAC <> "" Then .Fields("CorreoRAC") = p_Proyecto.CorreoRAC
        If IsNumeric(p_Proyecto.Ordinal) Then .Fields("Ordinal") = p_Proyecto.Ordinal
        If p_Proyecto.CadenaNombreAutorizados <> "" Then .Fields("CadenaNombreAutorizados") = p_Proyecto.CadenaNombreAutorizados
        If p_Proyecto.NombreParaNodo <> "" Then .Fields("NombreParaNodo") = p_Proyecto.NombreParaNodo
        
        .Update
    End With
    rcdDatos.Close

    ' 5. REGISTRAR PRIMERA EDICIÓN
    Set m_objEdicion = New Edicion
    With m_objEdicion
        .IDProyecto = p_Proyecto.IDProyecto
        .Aprobado = p_Proyecto.Aprobado
        .Elaborado = p_Proyecto.Elaborado
        .Revisado = p_Proyecto.Revisado
        .FechaEdicion = CStr(m_FechaRef)
        .IDEdicion = .IDEdicionCalculada
    End With
    
    Set rcdDatos = db.OpenRecordset("TbProyectosEdiciones")
    With rcdDatos
        .AddNew
        .Fields("IDEdicion") = m_objEdicion.IDEdicion
        m_objEdicion.Edicion = 1
        .Fields("Edicion") = m_objEdicion.Edicion
        .Fields("IDProyecto") = m_objEdicion.IDProyecto
        .Fields("Aprobado") = m_objEdicion.Aprobado
        .Fields("Elaborado") = m_objEdicion.Elaborado
        .Fields("Revisado") = m_objEdicion.Revisado
        .Fields("FechaEdicion") = m_objEdicion.FechaEdicion
        If IsDate(p_Proyecto.FechaMaxProximaPublicacion) Then .Fields("FechaMaxProximaPublicacion") = p_Proyecto.FechaMaxProximaPublicacion
        .Update
    End With
    rcdDatos.Close
    
    ' 6. INICIALIZAR CACHÉS
    CachePublicabilidad_RecalcularEdicionYResetear m_objEdicion, , p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' Inicializar caché de árbol (ahora con soporte de nodo raíz)
    CacheArbolRiesgosTx_RebuildEdicion m_objEdicion, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 7. ESTABLECER TAREAS
    If p_EsTecnico = EnumSiNo.Sí Then
        EstablecerTareasTecnico EnumSiNo.Sí, p_Error
    Else
        EstablecerTareasCalidad EnumSiNo.Sí, EnumSiNo.No, p_Error
    End If
    If p_Error <> "" Then Err.Raise 1000

    ' 8. FIN DE TRANSACCIÓN
    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    
    CrearProyectoTransaccional = ""
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en CrearProyectoTransaccional: " & Err.Description
    End If
    CrearProyectoTransaccional = p_Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : EditarProyectoTransaccional
' Purpose   : Realiza la edición de un Proyecto.
'---------------------------------------------------------------------------------------
Public Function EditarProyectoTransaccional( _
                            p_Proyecto As Proyecto, _
                            p_ProyectoAlInicio As Proyecto, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim rcdDatos As DAO.Recordset
    Dim m_objEdicion As Edicion
    Dim m_motivosNoOK As String
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If

    m_EnTransaccion = False
    
    ' Sincronizar propiedades calculadas
    With p_Proyecto
        If .Proyecto = "" Then .Proyecto = .ProyectoCalculado
        If .Juridica = "" Then .Juridica = .JuridicaCalculada
        If .NombreProyecto = "" Then .NombreProyecto = .NombreProyectoCalculado
        If .Cliente = "" Then .Cliente = .ClienteCalculado
        .FechaPrevistaCierre = .FechaPrevistaCierreCalculada
        If .CodigoDocumento = "" Then .CodigoDocumento = .CodigoDocumentoCalculado
        If .FechaFirmaContrato = "" Then .FechaFirmaContrato = .FechaFirmaContratoCalculada
        If .NombreUsuarioCalidad = "" Then .NombreUsuarioCalidad = .NombreUsuarioCalidadCalculado
        If .EnUTE <> "Sí" And .EnUTE <> "No" Then .EnUTE = .EnUTECalculado
        If .FechaMaxProximaPublicacion = "" Then .FechaMaxProximaPublicacion = .FechaMaxProximaPublicacionCalculada
        If .Ordinal = "" Then .Ordinal = .OrdinalCalculado
        .NombreParaNodo = .NombreParaNodoCalculado
    End With

    m_motivosNoOK = p_Proyecto.MotivoNoOK(p_ObjProyectoAlInicio:=p_ProyectoAlInicio, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_motivosNoOK <> "" Then
        p_Error = m_motivosNoOK
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    ' 1. ACTUALIZAR TbProyectos
    Set rcdDatos = db.OpenRecordset("SELECT * FROM TbProyectos WHERE IDProyecto=" & p_Proyecto.IDProyecto)
    With rcdDatos
        If .EOF Then
            p_Error = "No se ha podido obtener el registro del proyecto para editar"
            Err.Raise 1000
        End If
        .Edit
            .Fields("IDExpediente") = IIf(p_Proyecto.IDExpediente <> "", p_Proyecto.IDExpediente, Null)
            .Fields("Proyecto") = IIf(p_Proyecto.Proyecto <> "", p_Proyecto.Proyecto, Null)
            .Fields("Juridica") = IIf(p_Proyecto.Juridica <> "", p_Proyecto.Juridica, Null)
            .Fields("NombreProyecto") = IIf(p_Proyecto.NombreProyecto <> "", p_Proyecto.NombreProyecto, Null)
            .Fields("Cliente") = IIf(p_Proyecto.Cliente <> "", p_Proyecto.Cliente, Null)
            .Fields("FechaPrevistaCierre") = IIf(IsDate(p_Proyecto.FechaPrevistaCierre), p_Proyecto.FechaPrevistaCierre, Null)
            .Fields("FechaCierre") = IIf(IsDate(p_Proyecto.FechaCierre), p_Proyecto.FechaCierre, Null)
            .Fields("Elaborado") = IIf(p_Proyecto.Elaborado <> "", p_Proyecto.Elaborado, Null)
            .Fields("Revisado") = IIf(p_Proyecto.Revisado <> "", p_Proyecto.Revisado, Null)
            .Fields("Aprobado") = IIf(p_Proyecto.Aprobado <> "", p_Proyecto.Aprobado, Null)
            .Fields("CodigoDocumento") = IIf(p_Proyecto.CodigoDocumento <> "", p_Proyecto.CodigoDocumento, Null)
            .Fields("ParaInformeAvisos") = IIf(p_Proyecto.ParaInformeAvisos <> "", p_Proyecto.ParaInformeAvisos, Null)
            .Fields("FechaFirmaContrato") = IIf(IsDate(p_Proyecto.FechaFirmaContrato), p_Proyecto.FechaFirmaContrato, Null)
            .Fields("NombreUsuarioCalidad") = IIf(p_Proyecto.NombreUsuarioCalidad <> "", p_Proyecto.NombreUsuarioCalidad, Null)
            .Fields("EnUTE") = IIf(p_Proyecto.EnUTE = "Sí" Or p_Proyecto.EnUTE = "No", p_Proyecto.EnUTE, Null)
            .Fields("FechaMaxProximaPublicacion") = IIf(IsDate(p_Proyecto.FechaMaxProximaPublicacion), p_Proyecto.FechaMaxProximaPublicacion, Null)
            .Fields("CorreoRAC") = IIf(p_Proyecto.CorreoRAC <> "", p_Proyecto.CorreoRAC, Null)
            .Fields("Ordinal") = IIf(IsNumeric(p_Proyecto.Ordinal), p_Proyecto.Ordinal, Null)
            .Fields("CadenaNombreAutorizados") = IIf(p_Proyecto.CadenaNombreAutorizados <> "", p_Proyecto.CadenaNombreAutorizados, Null)
            .Fields("NombreParaNodo") = IIf(p_Proyecto.NombreParaNodo <> "", p_Proyecto.NombreParaNodo, Null)
        .Update
    End With
    rcdDatos.Close

    ' 2. ACTUALIZAR EDICIÓN ACTIVA
    Set m_objEdicion = p_Proyecto.EdicionActiva
    If Not m_objEdicion Is Nothing Then
        Set rcdDatos = db.OpenRecordset("SELECT * FROM TbProyectosEdiciones WHERE IDEdicion=" & m_objEdicion.IDEdicion)
        With rcdDatos
            If Not .EOF Then
                .Edit
                    .Fields("Aprobado") = p_Proyecto.Aprobado
                    .Fields("Elaborado") = p_Proyecto.Elaborado
                    .Fields("Revisado") = p_Proyecto.Revisado
                .Update
            End If
        End With
        rcdDatos.Close
        
        ' Sincronizar caches
        CachePublicabilidad_RecalcularEdicionYResetear m_objEdicion, , p_Error
        If p_Error <> "" Then Err.Raise 1000
        
        ' Actualizar el nodo raíz en el árbol (por si cambió el título del proyecto)
        CacheArbolRiesgosTx_ActualizarEdicion m_objEdicion, db, p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    
    EditarProyectoTransaccional = ""
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en EditarProyectoTransaccional: " & Err.Description
    End If
    EditarProyectoTransaccional = p_Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : BorrarProyectoTransaccional
' Purpose   : Borra un Proyecto y limpia TODAS sus cachés de forma atómica.
'---------------------------------------------------------------------------------------
Public Function BorrarProyectoTransaccional( _
                            p_Proyecto As Proyecto, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_IdProj As String
    
    On Error GoTo errores
    
    m_IdProj = p_Proyecto.IDProyecto
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    ' 1. Limpiar caches de todas las ediciones del proyecto
    db.Execute "DELETE FROM TbCacheArbolRiesgosMeta WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IdProj & ")"
    db.Execute "DELETE FROM TbCacheArbolRiesgosNodo WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IdProj & ")"
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IdProj & ")"
    db.Execute "DELETE FROM TbCacheControlCambiosMeta WHERE IDProyecto=" & m_IdProj
    db.Execute "DELETE FROM TbCacheControlCambiosRow WHERE IDProyecto=" & m_IdProj

    ' 2. Borrar proyecto y datos relacionados (asumiendo Cascade Delete en el motor DB para ediciones/riesgos)
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IdProj
    db.Execute "DELETE FROM TbUltimoProyecto WHERE IDProyecto=" & m_IdProj

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    
    BorrarProyectoTransaccional = ""
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en BorrarProyectoTransaccional: " & Err.Description
    BorrarProyectoTransaccional = p_Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistrarCorreoRACTransaccional
' Purpose   : Registra el Correo RAC y sincroniza caches.
'---------------------------------------------------------------------------------------
Public Function RegistrarCorreoRACTransaccional( _
                                    p_Proyecto As Proyecto, _
                                    p_CorreoRAC As String, _
                                    Optional ByVal p_db As DAO.Database, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    
    On Error GoTo errores
    
    If p_CorreoRAC = "" Then
        p_Error = "Se ha de indicar el correo del RAC"
        Err.Raise 1000
    End If

    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "UPDATE TbProyectos SET CorreoRAC = '" & p_CorreoRAC & "' WHERE IDProyecto=" & p_Proyecto.IDProyecto
    
    If Not p_Proyecto.EdicionActiva Is Nothing Then
        CachePublicabilidad_RecalcularEdicionYResetear p_Proyecto.EdicionActiva, , p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en RegistrarCorreoRACTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : EstablecerFechaMaxPublicacionTransaccional
' Purpose   : Establece la fecha máxima de publicación y sincroniza caches.
'---------------------------------------------------------------------------------------
Public Function EstablecerFechaMaxPublicacionTransaccional( _
                                                p_Proyecto As Proyecto, _
                                                Optional ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_Fecha As String
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    m_Fecha = p_Proyecto.FechaMaxProximaPublicacionCalculada
    If IsDate(m_Fecha) Then m_Fecha = "#" & Format(m_Fecha, "mm/dd/yyyy") & "#" Else m_Fecha = "Null"

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "UPDATE TbProyectos SET FechaMaxProximaPublicacion = " & m_Fecha & " WHERE IDProyecto=" & p_Proyecto.IDProyecto
    
    If Not p_Proyecto.EdicionActiva Is Nothing Then
        CachePublicabilidad_RecalcularEdicionYResetear p_Proyecto.EdicionActiva, , p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en EstablecerFechaMaxPublicacionTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : AbrirProyectoTransaccional
' Purpose   : Abre un proyecto cerrado y sincroniza caches.
'---------------------------------------------------------------------------------------
Public Function AbrirProyectoTransaccional( _
                                p_Proyecto As Proyecto, _
                                Optional ByVal p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    Set m_Edicion = p_Proyecto.EdicionUltima
    If m_Edicion Is Nothing Then
        p_Error = "No se ha podido obtener la última edición"
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    m_Edicion.EliminarPublicacion p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    m_Edicion.AbrirAcciones p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    db.Execute "UPDATE TbProyectos SET FechaCierre = Null WHERE IDProyecto=" & p_Proyecto.IDProyecto
    
    If Not p_Proyecto.EdicionActiva Is Nothing Then
        CachePublicabilidad_RecalcularEdicionYResetear p_Proyecto.EdicionActiva, , p_Error
        If p_Error <> "" Then Err.Raise 1000
        
        CacheArbolRiesgosTx_ActualizarEdicion p_Proyecto.EdicionActiva, db, p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en AbrirProyectoTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : FechaPrevistaCierreGrabarTransaccional
' Purpose   : Graba la fecha prevista de cierre.
'---------------------------------------------------------------------------------------
Public Function FechaPrevistaCierreGrabarTransaccional( _
                                            p_Proyecto As Proyecto, _
                                            Optional ByVal p_db As DAO.Database, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_Fecha As String
    
    On Error GoTo errores
    
    m_Fecha = p_Proyecto.FechaPrevistaCierreCalculada
    If Not IsDate(m_Fecha) Then Exit Function

    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "UPDATE TbProyectos SET FechaPrevistaCierre = #" & Format(m_Fecha, "mm/dd/yyyy") & "# WHERE IDProyecto=" & p_Proyecto.IDProyecto

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en FechaPrevistaCierreGrabarTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : FechaMaxProximaPublicacionGrabarTransaccional
' Purpose   : Graba la fecha máxima de próxima publicación.
'---------------------------------------------------------------------------------------
Public Function FechaMaxProximaPublicacionGrabarTransaccional( _
                                                p_Proyecto As Proyecto, _
                                                Optional ByVal p_db As DAO.Database, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_Fecha As String
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    
    m_Fecha = p_Proyecto.FechaMaxProximaPublicacionCalculada
    If IsDate(m_Fecha) Then m_Fecha = "#" & Format(m_Fecha, "mm/dd/yyyy") & "#" Else m_Fecha = "Null"

    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "UPDATE TbProyectos SET FechaMaxProximaPublicacion = " & m_Fecha & " WHERE IDProyecto=" & p_Proyecto.IDProyecto
    
    Set m_Edicion = p_Proyecto.EdicionActiva
    If Not m_Edicion Is Nothing Then
        db.Execute "UPDATE TbProyectosEdiciones SET FechaMaxProximaPublicacion = " & m_Fecha & " WHERE IDEdicion=" & m_Edicion.IDEdicion
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en FechaMaxProximaPublicacionGrabarTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistrarSuministradorTransaccional
' Purpose   : Registra un suministrador en un proyecto.
'---------------------------------------------------------------------------------------
Public Function RegistrarSuministradorTransaccional( _
                                        p_Proyecto As Proyecto, _
                                        p_IDSuministrador As String, _
                                        p_GestionCalidad As String, _
                                        Optional ByVal p_db As DAO.Database, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    Dim m_ProyectoSuministrador As ProyectoSuministrador
    Dim m_ProyectoSuministradorAlInicio As ProyectoSuministrador
    
    On Error GoTo errores
    
    If p_GestionCalidad <> "Sí" And p_GestionCalidad <> "No" Then
        p_Error = "Se ha de indicar si va a llevar la GestionCalidad o no"
        Err.Raise 1000
    End If

    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    Set m_ProyectoSuministrador = Constructor.getSuministradorEnProyecto(, p_Proyecto.IDProyecto, p_IDSuministrador, p_Error)
    If Not m_ProyectoSuministrador Is Nothing Then
        Set m_ProyectoSuministradorAlInicio = Constructor.getSuministradorEnProyecto(m_ProyectoSuministrador.ID, , , p_Error)
        m_ProyectoSuministrador.GestionCalidad = p_GestionCalidad
    Else
        Set m_ProyectoSuministrador = New ProyectoSuministrador
        With m_ProyectoSuministrador
            .IDProyecto = p_Proyecto.IDProyecto
            .IDSuministrador = p_IDSuministrador
            .GestionCalidad = p_GestionCalidad
        End With
    End If

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    m_ProyectoSuministrador.Registrar m_ProyectoSuministradorAlInicio, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en RegistrarSuministradorTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : CerrarProyectoTransaccional
' Purpose   : Cierra un Proyecto y sincroniza caches.
'---------------------------------------------------------------------------------------
Public Function CerrarProyectoTransaccional( _
                                p_Proyecto As Proyecto, _
                                Optional ByVal p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "UPDATE TbProyectos SET FechaCierre = #" & Format(p_Proyecto.FechaCierre, "mm/dd/yyyy") & "# WHERE IDProyecto=" & p_Proyecto.IDProyecto
    
    If Not p_Proyecto.EdicionActiva Is Nothing Then
        CachePublicabilidad_RecalcularEdicionYResetear p_Proyecto.EdicionActiva, , p_Error
        If p_Error <> "" Then Err.Raise 1000
        
        CacheArbolRiesgosTx_ActualizarEdicion p_Proyecto.EdicionActiva, db, p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en CerrarProyectoTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : BorrarCambiosProyectoTransaccional
' Purpose   : Borra los cambios de un Proyecto.
'---------------------------------------------------------------------------------------
Public Function BorrarCambiosProyectoTransaccional( _
                                p_Proyecto As Proyecto, _
                                Optional ByVal p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim m_EnTransaccion As Boolean
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then GoTo errores
    Else
        Set db = p_db
    End If
    
    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    db.Execute "DELETE FROM TbCambios WHERE IDProyecto=" & p_Proyecto.IDProyecto

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    p_Error = "Error en BorrarCambiosProyectoTransaccional: " & Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : GrabarCambiosEnProyectoTransaccional
' Purpose   : Graba los cambios de todas las ediciones de un Proyecto.
'---------------------------------------------------------------------------------------
Public Function GrabarCambiosEnProyectoTransaccional( _
                                            p_Proyecto As Proyecto, _
                                            Optional ByVal p_DesdeEdicion As String, _
                                            Optional ByVal p_db As DAO.Database, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim ws As DAO.Workspace
    Dim m_EnTransaccion As Boolean
    Dim m_objColEdiciones As Scripting.Dictionary
    Dim m_IDEdicion As Variant
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    
    If Not IsNumeric(p_DesdeEdicion) Then p_DesdeEdicion = "1"
    
    ' Obtenemos la colección de ediciones filtrada
    Set m_objColEdiciones = New Scripting.Dictionary
    m_objColEdiciones.CompareMode = TextCompare
    
    For Each m_IDEdicion In p_Proyecto.colEdiciones
        Set m_objEdicion = p_Proyecto.colEdiciones(m_IDEdicion)
        If CInt(m_objEdicion.Edicion) >= CInt(p_DesdeEdicion) Then
            m_objColEdiciones.Add CStr(m_IDEdicion), m_objEdicion
        End If
    Next
    
    If m_objColEdiciones.Count = 0 Then Exit Function

    If p_db Is Nothing Then
        ws.BeginTrans
        m_EnTransaccion = True
    End If

    For Each m_IDEdicion In m_objColEdiciones
        Set m_objEdicion = m_objColEdiciones(m_IDEdicion)
        m_objEdicion.GrabarCambiosEnEdicion p_Error
        If p_Error <> "" Then Err.Raise 1000
    Next

    If m_EnTransaccion Then
        ws.CommitTrans
        m_EnTransaccion = False
    End If
    
    Exit Function

errores:
    If m_EnTransaccion Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en GrabarCambiosEnProyectoTransaccional: " & Err.Description
    End If
End Function


