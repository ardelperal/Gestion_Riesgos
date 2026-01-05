Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : SuministradorRepository
' Purpose   : Gestión de persistencia para la tabla de Suministradores.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_ID As Long, Optional ByRef p_Error As String) As Suministrador
    Dim db As DAO.Database: Dim rcd As DAO.Recordset: Dim m_Obj As Suministrador
    On Error GoTo errores
    Set db = DatabaseProvider.GetGestionDB(p_Error)
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbSuministradores WHERE IDSuministrador=" & p_ID)
    If Not rcd.EOF Then
        Set m_Obj = New Suministrador
        With m_Obj
            .IDSuministrador = Nz(rcd!IDSuministrador, 0)
            .Nemotecnico = Nz(rcd!Nemotecnico, "")
            .NombreSuministrador = Nz(rcd!NombreSuministrador, "")
            .CIF = Nz(rcd!CIF, "")
            .Activo = Nz(rcd!Activo, "N")
        End With
        Set GetById = m_Obj
    End If
    rcd.Close
    Exit Function
errores:
    p_Error = "Error en SuministradorRepository.GetById: " & Err.Description
End Function
