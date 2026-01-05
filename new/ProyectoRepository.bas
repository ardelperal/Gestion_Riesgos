Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ProyectoRepository
' Purpose   : Capa de persistencia para la tabla TbProyectos.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_IDProyecto As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As Proyecto
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Proj As Proyecto
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbProyectos WHERE IDProyecto=" & p_IDProyecto)
    If Not rcd.EOF Then
        Set m_Proj = New Proyecto
        With m_Proj
            .IDProyecto = Nz(rcd!IDProyecto, 0)
            .IDExpediente = Nz(rcd!IDExpediente, 0)
            .Proyecto = Nz(rcd!Proyecto, "")
            .Juridica = Nz(rcd!Juridica, "")
            .NombreProyecto = Nz(rcd!NombreProyecto, "")
            .Cliente = Nz(rcd!Cliente, "")
            .CodigoDocumento = Nz(rcd!CodigoDocumento, "")
            .FechaPrevistaCierre = rcd!FechaPrevistaCierre
            .FechaCierre = rcd!FechaCierre
            .FechaRegistroInicial = rcd!FechaRegistroInicial
            .FechaFirmaContrato = rcd!FechaFirmaContrato
            .FechaMaxProximaPublicacion = rcd!FechaMaxProximaPublicacion
            .Elaborado = Nz(rcd!Elaborado, "")
            .Revisado = Nz(rcd!Revisado, "")
            .Aprobado = Nz(rcd!Aprobado, "")
            .NombreUsuarioCalidad = Nz(rcd!NombreUsuarioCalidad, "")
            .CorreoRAC = Nz(rcd!CorreoRAC, "")
            .ParaInformeAvisos = Nz(rcd!ParaInformeAvisos, "No")
            .EnUTE = Nz(rcd!EnUTE, "No")
            .RequiereRiesgoDeBiblioteca = Nz(rcd!RequiereRiesgoDeBiblioteca, "No")
            .RiesgosDeLaOferta = Nz(rcd!RiesgosDeLaOferta, "No")
            .RiesgosDelSubContratista = Nz(rcd!RiesgosDelSubContratista, "No")
            .Ordinal = Nz(rcd!Ordinal, 0)
            .CadenaNombreAutorizados = Nz(rcd!CadenaNombreAutorizados, "")
            .NombreParaNodo = Nz(rcd!NombreParaNodo, "")
        End With
        Set GetById = m_Proj
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en ProyectoRepository.GetById: " & Err.Description
End Function

Public Function Save(p_Proj As Proyecto, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = DatabaseProvider.GetGestionDB(p_Error) Else Set db = p_db
    If p_Error <> "" Then Err.Raise 1000
    
    If p_Proj.IDProyecto = 0 Then
        Set rcd = db.OpenRecordset("TbProyectos")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbProyectos WHERE IDProyecto=" & p_Proj.IDProyecto)
        If rcd.EOF Then
            p_Error = "No se encontró el proyecto para actualizar"
            GoTo errores
        End If
        rcd.Edit
    End If
    
    With rcd
        If p_Proj.IDProyecto <> 0 Then !IDProyecto = p_Proj.IDProyecto
        !IDExpediente = IIf(p_Proj.IDExpediente = 0, Null, p_Proj.IDExpediente)
        !Proyecto = p_Proj.Proyecto
        !Juridica = p_Proj.Juridica
        !NombreProyecto = p_Proj.NombreProyecto
        !Cliente = p_Proj.Cliente
        !CodigoDocumento = p_Proj.CodigoDocumento
        !FechaPrevistaCierre = p_Proj.FechaPrevistaCierre
        !FechaCierre = p_Proj.FechaCierre
        !FechaRegistroInicial = p_Proj.FechaRegistroInicial
        !FechaFirmaContrato = p_Proj.FechaFirmaContrato
        !FechaMaxProximaPublicacion = p_Proj.FechaMaxProximaPublicacion
        !Elaborado = p_Proj.Elaborado
        !Revisado = p_Proj.Revisado
        !Aprobado = p_Proj.Aprobado
        !NombreUsuarioCalidad = p_Proj.NombreUsuarioCalidad
        !CorreoRAC = p_Proj.CorreoRAC
        !ParaInformeAvisos = p_Proj.ParaInformeAvisos
        !EnUTE = p_Proj.EnUTE
        !RequiereRiesgoDeBiblioteca = p_Proj.RequiereRiesgoDeBiblioteca
        !RiesgosDeLaOferta = p_Proj.RiesgosDeLaOferta
        !RiesgosDelSubContratista = p_Proj.RiesgosDelSubContratista
        !Ordinal = p_Proj.Ordinal
        !CadenaNombreAutorizados = p_Proj.CadenaNombreAutorizados
        !NombreParaNodo = p_Proj.NombreParaNodo
        .Update
    End With
    
    Save = "OK"
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en ProyectoRepository.Save: " & Err.Description
End Function
