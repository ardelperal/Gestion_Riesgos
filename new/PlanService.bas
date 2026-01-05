Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : PlanService
' Purpose   : Lógica de orquestación para Planes y Acciones.
'---------------------------------------------------------------------------------------

Public Function FinalizarAccionMitigacion( _
                            p_Accion As PMAccion, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim blnEnTrans As Boolean
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    Set db = DatabaseProvider.GetGestionDB(p_Error)
    
    ws.BeginTrans
    blnEnTrans = True
    
    ' 1. Marcar acción como finalizada
    p_Accion.Estado = "Finalizada"
    p_Accion.FechaFinReal = Now()
    
    PlanRepository.SavePMAccion p_Accion, db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 2. Lógica de Negocio: Si es la última acción, cerrar el Plan
    If p_Accion.EsUltimaAccion = "Sí" Then
        ' Deberíamos obtener el Plan padre y actualizarlo
        ' Dim m_PM As PM: Set m_PM = PlanRepository.GetPMById(p_Accion.IDMitigacion, db)
        ' m_PM.Estado = "Finalizado"
        ' PlanRepository.SavePM m_PM, db, p_Error
    End If
    
    ws.CommitTrans
    blnEnTrans = False
    
    FinalizarAccionMitigacion = "OK"
    Exit Function

errores:
    If blnEnTrans Then ws.Rollback
    p_Error = "Error en PlanService.FinalizarAccionMitigacion: " & Err.Description
End Function
