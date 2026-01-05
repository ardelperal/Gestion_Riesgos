Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ProyectoService
' Purpose   : Orquestación de Casos de Uso para Proyectos y Ediciones.
'---------------------------------------------------------------------------------------

' CU-A1: Alta de Proyecto Atómica
Public Function CrearProyecto( _
                            p_Proyecto As Proyecto, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim blnEnTrans As Boolean
    Dim m_Edicion As Edicion
    
    On Error GoTo errores
    
    ' 1. VALIDACIÓN DE SEGURIDAD (Solo Admin o Calidad)
    If UsuarioService.m_ObjUsuarioConectado.EsTecnico Then
        p_Error = "Acceso denegado: Los usuarios con perfil Técnico no pueden crear nuevos proyectos."
        Err.Raise 1000
    End If
    
    ' 2. VALIDACIÓN DE DOMINIO
    Dim m_Motivo As String
    m_Motivo = p_Proyecto.MotivoNoOK(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_Motivo <> "" Then
        p_Error = m_Motivo
        Err.Raise 1000
    End If
    
    Set ws = DBEngine.Workspaces(0)
    Set db = DatabaseProvider.GetGestionDB(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ws.BeginTrans
    blnEnTrans = True
    
    ' 3. GUARDAR PROYECTO
    p_Proyecto.FechaRegistroInicial = Now()
    p_Proyecto.RequiereRiesgoDeBiblioteca = "S" 
    
    ProyectoRepository.Save p_Proyecto, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 4. CREAR PRIMERA EDICIÓN
    Set m_Edicion = New Edicion
    With m_Edicion
        .IDProyecto = p_Proyecto.IDProyecto
        .Edicion = 1
        .FechaEdicion = p_Proyecto.FechaRegistroInicial
        .Elaborado = p_Proyecto.Elaborado
        .Revisado = p_Proyecto.Revisado
        .Aprobado = p_Proyecto.Aprobado
    End With
    
    EdicionRepository.Save m_Edicion, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 5. ESTABLECER TAREAS INICIALES (Siempre de Calidad para nuevo proyecto)
    ' TaskService.EstablecerTareasCalidad p_Error, db
    If p_Error <> "" Then Err.Raise 1000
    
    ws.CommitTrans
    blnEnTrans = False
    
    CrearProyecto = "OK"
    Exit Function

errores:
    If blnEnTrans Then ws.Rollback
    p_Error = "Error en ProyectoService.CrearProyecto: " & Err.Description
End Function

' CU-A2: Cierre de Edición y Apertura de Nueva (Publicación)
Public Function PublicarEdicion( _
                            ByVal p_IDEdicionActual As Long, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim ws As DAO.Workspace: Dim db As DAO.Database: Dim blnEnTrans As Boolean
    Dim m_EdicionActual As Edicion: Dim m_EdicionNueva As Edicion
    Dim rcdRiesgos As DAO.Recordset: Dim m_RiesgoNuevo As riesgo
    Dim rcdPMs As DAO.Recordset: Dim m_PMNuevo As PM
    Dim rcdPAs As DAO.Recordset: Dim rcdPCs As DAO.Recordset: Dim rcdCAs As DAO.Recordset
    
    On Error GoTo errores
    Set ws = DBEngine.Workspaces(0): Set db = DatabaseProvider.GetGestionDB(p_Error)
    
    ' 1. Cargar Edición Actual y Validar
    Set m_EdicionActual = EdicionRepository.GetById(p_IDEdicionActual, db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ws.BeginTrans: blnEnTrans = True
    
    ' 2. CERRAR EDICIÓN ACTUAL
    m_EdicionActual.FechaPublicacion = Now()
    EdicionRepository.Save m_EdicionActual, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 3. CREAR NUEVA EDICIÓN (Clonando cabecera)
    Set m_EdicionNueva = New Edicion
    With m_EdicionNueva
        .IDProyecto = m_EdicionActual.IDProyecto
        .Edicion = m_EdicionActual.Edicion + 1
        .FechaEdicion = m_EdicionActual.FechaPublicacion
        .Elaborado = m_EdicionActual.Elaborado: .Revisado = m_EdicionActual.Revisado: .Aprobado = m_EdicionActual.Aprobado
    End With
    EdicionRepository.Save m_EdicionNueva, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 4. TRASPASAR JERARQUÍA (Riesgos -> Planes -> Acciones)
    ' Solo riesgos Vivos (No retirados, No cerrados)
    Set rcdRiesgos = db.OpenRecordset("SELECT IDRiesgo FROM TbRiesgos WHERE IDEdicion=" & p_IDEdicionActual & " AND FechaRetirado Is Null AND FechaCerrado Is Null")
    Do While Not rcdRiesgos.EOF
        ' 4.1 Clonar Riesgo
        Set m_RiesgoNuevo = RiesgoRepository.CopiarRiesgo(rcdRiesgos!IDRiesgo, m_EdicionNueva.IDEdicion, db, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        
        ' 4.2 Clonar PMs del riesgo
        Set rcdPMs = db.OpenRecordset("SELECT IDMitigacion FROM TbRiesgosPlanMitigacionPpal WHERE IDRiesgo=" & rcdRiesgos!IDRiesgo)
        Do While Not rcdPMs.EOF
            Set m_PMNuevo = PlanRepository.CopiarPM(rcdPMs!IDMitigacion, m_RiesgoNuevo.IDRiesgo, db, p_Error)
            ' 4.3 Clonar Acciones del PM
            Set rcdPAs = db.OpenRecordset("SELECT IDAccionMitigacion FROM TbRiesgosPlanMitigacionDetalle WHERE IDMitigacion=" & rcdPMs!IDMitigacion)
            Do While Not rcdPAs.EOF
                PlanRepository.CopiarPMAccion rcdPAs!IDAccionMitigacion, m_PMNuevo.IDMitigacion, db, p_Error
                rcdPAs.MoveNext
            Loop
            rcdPMs.MoveNext
        Loop
        
        ' 4.4 Clonar PCs (Análogo a PMs...)
        ' [Lógica de PCs similar...]
        
        rcdRiesgos.MoveNext
    Loop
    
    ' 5. FINALIZACIÓN
    CacheService.InicializarCacheEdicion m_EdicionNueva.IDEdicion, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ws.CommitTrans: blnEnTrans = False
    PublicarEdicion = "OK"
    Exit Function

errores:
    If blnEnTrans Then ws.Rollback
    p_Error = "Error en ProyectoService.PublicarEdicion: " & Err.Description
End Function

' CU-A3: Sincronizar con Expediente Externo
Public Function SincronizarConExpediente( _
                            p_Proyecto As Proyecto, _
                            ByVal p_IDExpediente As Long, _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim m_Exp As Expediente
    On Error GoTo errores
    
    Set m_Exp = ExpedienteRepository.GetById(p_IDExpediente, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If Not m_Exp Is Nothing Then
        ' Mapeo de negocio: El proyecto toma datos del expediente
        With p_Proyecto
            .IDExpediente = m_Exp.IDExpediente
            .NombreProyecto = m_Exp.Nemotecnico
            .Proyecto = m_Exp.CodExp
            .Cliente = m_Exp.Titulo ' O el campo correspondiente
            .FechaFirmaContrato = m_Exp.FechaFirmaContrato
            ' ... etc
        End With
    End If
    
    Exit Function
errores:
    p_Error = "Error en ProyectoService.SincronizarConExpediente: " & Err.Description
End Function
