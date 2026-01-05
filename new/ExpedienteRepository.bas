Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ExpedienteRepository
' Purpose   : Acceso a datos de la tabla TbExpedientes en la BD externa.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_IDExpediente As Long, Optional ByRef p_Error As String) As Expediente
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Exp As Expediente
    
    On Error GoTo errores
    Set db = getdbExpedientes() ' Función global en legacy/Variables Globales
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente)
    If Not rcd.EOF Then
        Set m_Exp = New Expediente
        With m_Exp
            .IDExpediente = Nz(rcd!IDExpediente, 0)
            .IDExpedientePadre = Nz(rcd!IDExpedientePadre, 0)
            .Nemotecnico = Nz(rcd!Nemotecnico, "")
            .Titulo = Nz(rcd!Titulo, "")
            .CodProyecto = Nz(rcd!CodProyecto, "")
            .CodExp = Nz(rcd!CodExp, "")
            .CodExpLargo = Nz(rcd!CodExpLargo, "")
            .CodigoActividad = Nz(rcd!CodigoActividad, "")
            .FechaInicioContrato = rcd!FechaInicioContrato
            .FechaFinContrato = rcd!FechaFinContrato
            .FechaFinGarantia = rcd!FechaFinGarantia
            .FechaFirmaContrato = rcd!FechaFirmaContrato
            .FechaCreacion = rcd!FechaCreacion
            .FechaUltimoCambio = rcd!FechaUltimoCambio
            .Ordinal = Nz(rcd!Ordinal, 0)
            .IDOrganoContratacion = Nz(rcd!IDOrganoContratacion, 0)
            .IDResponsableCalidad = Nz(rcd!IDResponsableCalidad, "")
            .Tipo = Nz(rcd!Tipo, "")
            .Estado = Nz(rcd!Estado, "")
            .Ambito = Nz(rcd!Ambito, "")
        End With
        Set GetById = m_Exp
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en ExpedienteRepository.GetById: " & Err.Description
End Function
