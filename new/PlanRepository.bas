Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : PlanRepository
' Purpose   : Gestión de persistencia para Planes de Mitigación y Contingencia.
'---------------------------------------------------------------------------------------

' --- MITIGACIÓN (PM) ---

Public Function SavePM(p_PM As PM, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    If p_PM.IDMitigacion = 0 Then
        Set rcd = db.OpenRecordset("TbRiesgosPlanMitigacionPpal")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion=" & p_PM.IDMitigacion)
        If rcd.EOF Then
            p_Error = "No se encontró el plan de mitigación"
            GoTo errores
        End If
        rcd.Edit
    End If
    
    With rcd
        If p_PM.IDMitigacion <> 0 Then !IDMitigacion = p_PM.IDMitigacion
        !IDRiesgo = IIf(p_PM.IDRiesgo = 0, Null, p_PM.IDRiesgo)
        !CodMitigacion = p_PM.CodMitigacion
        !DisparadorDelPlan = p_PM.DisparadorDelPlan
        !Estado = p_PM.Estado
        !FechaDeActivacion = p_PM.FechaDeActivacion
        !FechaDesactivacion = p_PM.FechaDesactivacion
        .Update
    End With
    SavePM = "OK"
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en PlanRepository.SavePM: " & Err.Description
End Function

Public Function SavePMAccion(p_Accion As PMAccion, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    
    If p_Accion.IDAccionMitigacion = 0 Then
        Set rcd = db.OpenRecordset("TbRiesgosPlanMitigacionDetalle")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgosPlanMitigacionDetalle WHERE IDAccionMitigacion=" & p_Accion.IDAccionMitigacion)
        rcd.Edit
    End If
    
    With rcd
        If p_Accion.IDAccionMitigacion <> 0 Then !IDAccionMitigacion = p_Accion.IDAccionMitigacion
        !IDMitigacion = p_Accion.IDMitigacion
        !CodAccion = p_Accion.CodAccion
        !Accion = p_Accion.Accion
        !ResponsableAccion = p_Accion.ResponsableAccion
        !Estado = p_Accion.Estado
        !EsUltimaAccion = p_Accion.EsUltimaAccion
        !FechaInicio = p_Accion.FechaInicio
        !FechaFinPrevista = p_Accion.FechaFinPrevista
        !FechaFinReal = p_Accion.FechaFinReal
        .Update
    End With
    SavePMAccion = "OK"
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en PlanRepository.SavePMAccion: " & Err.Description
End Function

' --- CONTINGENCIA (PC) ---

Public Function SavePC(p_PC As PC, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    
    If p_PC.IDContingencia = 0 Then
        Set rcd = db.OpenRecordset("TbRiesgosPlanContingenciaPpal")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgosPlanContingenciaPpal WHERE IDContingencia=" & p_PC.IDContingencia)
        rcd.Edit
    End If
    
    With rcd
        If p_PC.IDContingencia <> 0 Then !IDContingencia = p_PC.IDContingencia
        !IDRiesgo = p_PC.IDRiesgo
        !CodContingencia = p_PC.CodContingencia
        !DisparadorDelPlan = p_PC.DisparadorDelPlan
        !Estado = p_PC.Estado
        !FechaDeActivacion = p_PC.FechaDeActivacion
        !FechaDesactivacion = p_PC.FechaDesactivacion
        .Update
    End With
    SavePC = "OK"
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en PlanRepository.SavePC: " & Err.Description
End Function

Public Function CopiarPM(ByVal p_IDPMOld As Long, ByVal p_IDRiesgoNew As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As PM
    Dim db As DAO.Database: Dim rcdOld As DAO.Recordset: Dim rcdNew As DAO.Recordset: Dim fld As DAO.Field
    Dim m_NewID As Long
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    Set rcdOld = db.OpenRecordset("SELECT * FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion=" & p_IDPMOld)
    If rcdOld.EOF Then Exit Function
    m_NewID = DameID("TbRiesgosPlanMitigacionPpal", "IDMitigacion", db, p_Error)
    Set rcdNew = db.OpenRecordset("TbRiesgosPlanMitigacionPpal")
    rcdNew.AddNew
        For Each fld In rcdOld.Fields
            Select Case fld.Name
                Case "IDMitigacion": rcdNew!IDMitigacion = m_NewID
                Case "IDRiesgo": rcdNew!IDRiesgo = p_IDRiesgoNew
                Case Else: rcdNew.Fields(fld.Name).Value = fld.Value
            End Select
        Next
    rcdNew.Update
    Set m_PM = New PM: m_PM.IDMitigacion = m_NewID: Set CopiarPM = m_PM
    rcdOld.Close: rcdNew.Close: Exit Function
errores:
    p_Error = "Error en PlanRepository.CopiarPM: " & Err.Description
End Function

Public Function CopiarPMAccion(ByVal p_IDAccionOld As Long, ByVal p_IDPMNew As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database: Dim rcdOld As DAO.Recordset: Dim rcdNew As DAO.Recordset: Dim fld As DAO.Field
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    Set rcdOld = db.OpenRecordset("SELECT * FROM TbRiesgosPlanMitigacionDetalle WHERE IDAccionMitigacion=" & p_IDAccionOld)
    If rcdOld.EOF Then Exit Function
    Set rcdNew = db.OpenRecordset("TbRiesgosPlanMitigacionDetalle")
    rcdNew.AddNew
        For Each fld In rcdOld.Fields
            Select Case fld.Name
                Case "IDAccionMitigacion": rcdNew!IDAccionMitigacion = DameID("TbRiesgosPlanMitigacionDetalle", "IDAccionMitigacion", db, p_Error)
                Case "IDMitigacion": rcdNew!IDMitigacion = p_IDPMNew
                Case Else: rcdNew.Fields(fld.Name).Value = fld.Value
            End Select
        Next
    rcdNew.Update: CopiarPMAccion = "OK": rcdOld.Close: rcdNew.Close: Exit Function
errores:
    p_Error = "Error en PlanRepository.CopiarPMAccion: " & Err.Description
End Function
