Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : RiesgoRepository
' Purpose   : Capa de persistencia pura para la tabla TbRiesgos.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_IDRiesgo As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As riesgo
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Riesgo As riesgo
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo)
    If Not rcd.EOF Then
        Set m_Riesgo = New riesgo
        With m_Riesgo
            .IDRiesgo = Nz(rcd!IDRiesgo, 0)
            .IDEdicion = Nz(rcd!IDEdicion, 0)
            .CodigoUnico = Nz(rcd!CodigoUnico, "")
            .CodigoRiesgo = Nz(rcd!CodigoRiesgo, "")
            .FechaDetectado = rcd!FechaDetectado ' Variant acepta Null directamente
            .DetectadoPor = Nz(rcd!DetectadoPor, "")
            .EntidadDetecta = Nz(rcd!EntidadDetecta, "")
            .Plazo = Nz(rcd!Plazo, "")
            .Calidad = Nz(rcd!Calidad, "")
            .Coste = Nz(rcd!Coste, "")
            .ImpactoGlobal = Nz(rcd!ImpactoGlobal, "")
            .Vulnerabilidad = Nz(rcd!Vulnerabilidad, "")
            .Valoracion = Nz(rcd!Valoracion, "")
            .Mitigacion = Nz(rcd!Mitigacion, "")
            .Contingencia = Nz(rcd!Contingencia, "")
            .RequierePlanContingencia = Nz(rcd!RequierePlanContingencia, "")
            .Descripcion = Nz(rcd!Descripcion, "")
            .CausaRaiz = Nz(rcd!CausaRaiz, "")
            .Estado = Nz(rcd!Estado, "")
            .FechaEstado = rcd!FechaEstado
            .Priorizacion = Nz(rcd!Priorizacion, 0)
            .FechaMaterializado = rcd!FechaMaterializado
            .FechaRetirado = rcd!FechaRetirado
            .FechaCerrado = rcd!FechaCerrado
            .Origen = Nz(rcd!Origen, "")
        End With
        Set GetById = m_Riesgo
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en RiesgoRepository.GetById: " & Err.Description
End Function

Public Function Save(p_Riesgo As riesgo, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    If p_Riesgo.IDRiesgo = 0 Then
        Set rcd = db.OpenRecordset("TbRiesgos")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbRiesgos WHERE IDRiesgo=" & p_Riesgo.IDRiesgo)
        If rcd.EOF Then
            p_Error = "No se encontró el riesgo para actualizar"
            GoTo errores
        End If
        rcd.Edit
    End If
    
    With rcd
        If p_Riesgo.IDRiesgo <> 0 Then !IDRiesgo = p_Riesgo.IDRiesgo
        !IDEdicion = IIf(p_Riesgo.IDEdicion = 0, Null, p_Riesgo.IDEdicion)
        !CodigoUnico = p_Riesgo.CodigoUnico
        !CodigoRiesgo = p_Riesgo.CodigoRiesgo
        !FechaDetectado = p_Riesgo.FechaDetectado
        !DetectadoPor = p_Riesgo.DetectadoPor
        !EntidadDetecta = p_Riesgo.EntidadDetecta
        !Plazo = p_Riesgo.Plazo
        !Calidad = p_Riesgo.Calidad
        !Coste = p_Riesgo.Coste
        !ImpactoGlobal = p_Riesgo.ImpactoGlobal
        !Vulnerabilidad = p_Riesgo.Vulnerabilidad
        !Valoracion = p_Riesgo.Valoracion
        !Mitigacion = p_Riesgo.Mitigacion
        !Contingencia = p_Riesgo.Contingencia
        !RequierePlanContingencia = p_Riesgo.RequierePlanContingencia
        !Descripcion = p_Riesgo.Descripcion
        !CausaRaiz = p_Riesgo.CausaRaiz
        !Estado = p_Riesgo.Estado
        !FechaEstado = p_Riesgo.FechaEstado
        !Priorizacion = IIf(p_Riesgo.Priorizacion = 0, Null, p_Riesgo.Priorizacion)
        !FechaMaterializado = p_Riesgo.FechaMaterializado
        !FechaRetirado = p_Riesgo.FechaRetirado
        !FechaCerrado = p_Riesgo.FechaCerrado
        !Origen = p_Riesgo.Origen
        .Update
    End With
    
    Save = "OK"
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en RiesgoRepository.Save: " & Err.Description
End Function

Public Function Delete(ByVal p_IDRiesgo As String, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgo
    Delete = "OK"
    Exit Function

errores:
    p_Error = "Error en RiesgoRepository.Delete: " & Err.Description
End Function

Public Function CopiarRiesgo( _
                            ByVal p_IDRiesgoOld As Long, _
                            ByVal p_IDEdicionNew As Long, _
                            Optional ByVal p_db As DAO.Database, _
                            Optional ByRef p_Error As String _
                            ) As riesgo
    Dim db As DAO.Database
    Dim rcdOld As DAO.Recordset
    Dim rcdNew As DAO.Recordset
    Dim fld As DAO.Field
    Dim m_NewRiesgo As riesgo
    Dim m_NewID As Long
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcdOld = db.OpenRecordset("SELECT * FROM TbRiesgos WHERE IDRiesgo=" & p_IDRiesgoOld)
    If rcdOld.EOF Then Exit Function
    
    m_NewID = DameID("TbRiesgos", "IDRiesgo", db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcdNew = db.OpenRecordset("TbRiesgos")
    rcdNew.AddNew
        For Each fld In rcdOld.Fields
            Select Case fld.Name
                Case "IDRiesgo": rcdNew!IDRiesgo = m_NewID
                Case "IDEdicion": rcdNew!IDEdicion = p_IDEdicionNew
                Case "Priorizacion": ' Se mantiene o recalcula según servicio
                    rcdNew!Priorizacion = fld.Value
                Case Else: rcdNew.Fields(fld.Name).Value = fld.Value
            End Select
        Next
    rcdNew.Update
    
    Set m_NewRiesgo = GetById(m_NewID, db, p_Error)
    Set CopiarRiesgo = m_NewRiesgo
    
    rcdOld.Close
    rcdNew.Close
    Exit Function
errores:
    p_Error = "Error en RiesgoRepository.CopiarRiesgo: " & Err.Description
End Function
