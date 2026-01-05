Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : EdicionRepository
' Purpose   : Capa de persistencia para la tabla TbProyectosEdiciones.
'---------------------------------------------------------------------------------------

Public Function GetById(ByVal p_IDEdicion As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As Edicion
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_Ed As Edicion
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = getdb() Else Set db = p_db
    
    Set rcd = db.OpenRecordset("SELECT * FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion)
    If Not rcd.EOF Then
        Set m_Ed = New Edicion
        With m_Ed
            .IDEdicion = Nz(rcd!IDEdicion, 0)
            .IDProyecto = Nz(rcd!IDProyecto, 0)
            .Edicion = Nz(rcd!Edicion, 0)
            .FechaEdicion = rcd!FechaEdicion
            .FechaPublicacion = rcd!FechaPublicacion
            .Elaborado = Nz(rcd!Elaborado, "")
            .Revisado = Nz(rcd!Revisado, "")
            .Aprobado = Nz(rcd!Aprobado, "")
            .EntregadoAClienteORAC = Nz(rcd!EntregadoAClienteORAC, "No")
            .Comentarios = Nz(rcd!Comentarios, "")
            .PermitidoImprimirExcel = Nz(rcd!PermitidoImprimirExcel, "No")
            .IDDocumentoAGEDO = Nz(rcd!IDDocumentoAGEDO, 0)
            .NombreArchivoInforme = Nz(rcd!NombreArchivoInforme, "")
            .FechaMaxProximaPublicacion = rcd!FechaMaxProximaPublicacion
            .FechaPreparadaParaPublicar = rcd!FechaPreparadaParaPublicar
            .UsuarioProponePublicar = Nz(rcd!UsuarioProponePublicar, "")
            .PropuestaRechazadaPorCalidadFecha = rcd!PropuestaRechazadaPorCalidadFecha
            .PropuestaRechazadaPorCalidadMotivo = Nz(rcd!PropuestaRechazadaPorCalidadMotivo, "")
            .UsuarioCalidadRechazaPropuesta = Nz(rcd!UsuarioCalidadRechazaPropuesta, "")
            .NotasCalidadParaPublicar = Nz(rcd!NotasCalidadParaPublicar, "")
            .FechaUltimoCambio = rcd!FechaUltimoCambio
            .UsuarioUltimoCambio = Nz(rcd!UsuarioUltimoCambio, "")
        End With
        Set GetById = m_Ed
    End If
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en EdicionRepository.GetById: " & Err.Description
End Function

Public Function Save(p_Ed As Edicion, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    If p_db Is Nothing Then Set db = getdb() Else Set db = p_db
    
    If p_Ed.IDEdicion = 0 Then
        Set rcd = db.OpenRecordset("TbProyectosEdiciones")
        rcd.AddNew
    Else
        Set rcd = db.OpenRecordset("SELECT * FROM TbProyectosEdiciones WHERE IDEdicion=" & p_Ed.IDEdicion)
        If rcd.EOF Then
            p_Error = "No se encontró la edición para actualizar"
            GoTo errores
        End If
        rcd.Edit
    End If
    
    With rcd
        If p_Ed.IDEdicion <> 0 Then !IDEdicion = p_Ed.IDEdicion
        !IDProyecto = IIf(p_Ed.IDProyecto = 0, Null, p_Ed.IDProyecto)
        !Edicion = p_Ed.Edicion
        !FechaEdicion = p_Ed.FechaEdicion
        !FechaPublicacion = p_Ed.FechaPublicacion
        !Elaborado = p_Ed.Elaborado
        !Revisado = p_Ed.Revisado
        !Aprobado = p_Ed.Aprobado
        !EntregadoAClienteORAC = p_Ed.EntregadoAClienteORAC
        !Comentarios = p_Ed.Comentarios
        !PermitidoImprimirExcel = p_Ed.PermitidoImprimirExcel
        !IDDocumentoAGEDO = IIf(p_Ed.IDDocumentoAGEDO = 0, Null, p_Ed.IDDocumentoAGEDO)
        !NombreArchivoInforme = p_Ed.NombreArchivoInforme
        !FechaMaxProximaPublicacion = p_Ed.FechaMaxProximaPublicacion
        !FechaPreparadaParaPublicar = p_Ed.FechaPreparadaParaPublicar
        !UsuarioProponePublicar = p_Ed.UsuarioProponePublicar
        !PropuestaRechazadaPorCalidadFecha = p_Ed.PropuestaRechazadaPorCalidadFecha
        !PropuestaRechazadaPorCalidadMotivo = p_Ed.PropuestaRechazadaPorCalidadMotivo
        !UsuarioCalidadRechazaPropuesta = p_Ed.UsuarioCalidadRechazaPropuesta
        !NotasCalidadParaPublicar = p_Ed.NotasCalidadParaPublicar
        !FechaUltimoCambio = p_Ed.FechaUltimoCambio
        !UsuarioUltimoCambio = p_Ed.UsuarioUltimoCambio
        .Update
    End With
    
    Save = "OK"
    rcd.Close
    Exit Function

errores:
    p_Error = "Error en EdicionRepository.Save: " & Err.Description
End Function
