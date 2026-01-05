Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : RiesgoService
' Purpose   : Orquestación de lógica de negocio para Riesgos.
'---------------------------------------------------------------------------------------

Public Function ProcesarRiesgoExterno( _
                            p_RiesgoExt As RiesgoExterno, _
                            Optional p_ObjRiesgoExternoAlInicio As RiesgoExterno, _
                            Optional ByRef p_Error As String _
                            ) As String

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim blnEnTrans As Boolean
    Dim m_Riesgo As riesgo
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    Set db = getdb() ' Supongo que esta función existe en legacy/Constructor o global
    
    ' 1. VALIDACIÓN DE NEGOCIO (Dominio)
    Dim m_MotivoNoOK As String
    m_MotivoNoOK = p_RiesgoExt.MotivoNoOK(p_ObjRiesgoExternoAlInicio, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_MotivoNoOK <> "" Then
        p_Error = m_MotivoNoOK
        Err.Raise 1000
    End If
    
    ws.BeginTrans
    blnEnTrans = True
    
    ' 2. LÓGICA DE PROMOCIÓN
    If p_RiesgoExt.Trasladar = "Sí" Then
        ' Obtener o crear Riesgo asociado
        If p_RiesgoExt.IDRiesgo = "" Then
            ' Alta de Riesgo nuevo (usando Repository)
            Set m_Riesgo = New riesgo
            ' Rellenar m_Riesgo desde p_RiesgoExt (Lógica que antes estaba en Riesgo.cls)
            With m_Riesgo
                .IDEdicion = p_RiesgoExt.IDEdicion
                .Descripcion = p_RiesgoExt.Descripcion
                .CausaRaiz = p_RiesgoExt.CausaRaiz
                .CodRiesgoBiblioteca = p_RiesgoExt.CodRiesgoBiblioteca
                .FechaDetectado = p_RiesgoExt.FechaDetectado
                .Estado = "Detectado" ' Valor inicial
                .Origen = "Oferta"
                ' ... resto de campos ...
            End With
            
            RiesgoRepository.Save m_Riesgo, db, p_Error
            If p_Error <> "" Then Err.Raise 1000
            p_RiesgoExt.IDRiesgo = m_Riesgo.IDRiesgo
        Else
            ' Actualizar Riesgo existente
            Set m_Riesgo = RiesgoRepository.GetById(p_RiesgoExt.IDRiesgo, db, p_Error)
            If p_Error <> "" Then Err.Raise 1000
            ' Sincronizar cambios...
            RiesgoRepository.Save m_Riesgo, db, p_Error
            If p_Error <> "" Then Err.Raise 1000
        End If
    Else
        ' Si antes se trasladaba y ahora no, borrar el riesgo asociado
        If p_RiesgoExt.IDRiesgo <> "" Then
            RiesgoRepository.Delete p_RiesgoExt.IDRiesgo, db, p_Error
            If p_Error <> "" Then Err.Raise 1000
            p_RiesgoExt.IDRiesgo = ""
        End If
    End If
    
    ' 3. PERSISTIR RIESGO EXTERNO
    ' (Aquí llamaríamos al RiesgoExternoRepository.Save una vez creado)
    ' RiesgoExternoRepository.Save p_RiesgoExt, db, p_Error
    
    ' 4. ACTUALIZAR CACHÉS (Service orquestador)
    ' CacheService.RefrescarRiesgo m_Riesgo, db, p_Error
    
    ws.CommitTrans
    blnEnTrans = False
    
    ProcesarRiesgoExterno = "OK"
    Exit Function

errores:
    If blnEnTrans Then ws.Rollback
    If Err.Number <> 1000 Then
        p_Error = "Error en RiesgoService.ProcesarRiesgoExterno: " & Err.Description
    End If
End Function
