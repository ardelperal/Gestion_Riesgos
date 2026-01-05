Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : TaskService
' Purpose   : Gestión de tareas (validaciones, visados, etc.)
'---------------------------------------------------------------------------------------

Public Function InicializarTareasProyecto( _
                            ByVal p_IDProyecto As Long, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    ' Lógica Legacy: Un proyecto nuevo genera una tarea de "Revisión Inicial" para Calidad
    Set rcd = db.OpenRecordset("TbTareas")
    rcd.AddNew
        rcd!IDProyecto = p_IDProyecto
        rcd!TipoTarea = "REVISION_INICIAL"
        rcd!EstadoTarea = "PENDIENTE"
        rcd!FechaAccion = Now()
    rcd.Update
    rcd.Close
    
    InicializarTareasProyecto = "OK"
    Exit Function

errores:
    p_Error = "Error en TaskService.InicializarTareasProyecto: " & Err.Description
End Function
