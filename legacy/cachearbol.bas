Option Compare Database
Option Explicit
Public Function CacheArbolRiesgos_UsarCache(Optional ByRef p_Error As String) As EnumSiNo
    On Error GoTo errores
    CacheArbolRiesgos_UsarCache = EnumSiNo.Sí
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_UsarCache ha devuelto un error: " & vbNewLine & Err.Description
    End If
End Function
Private Function CacheArbolRiesgos_ExisteTabla(p_db As DAO.Database, p_NombreTabla As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo errores
    CacheArbolRiesgos_ExisteTabla = False
    For Each tdf In p_db.TableDefs
        If tdf.Name = p_NombreTabla Then
            CacheArbolRiesgos_ExisteTabla = True
            Exit Function
        End If
    Next
    Exit Function
errores:
    CacheArbolRiesgos_ExisteTabla = False
End Function
Private Function CacheArbolRiesgos_SqlText(p_Value As Variant) As String
    If IsNull(p_Value) Then
        CacheArbolRiesgos_SqlText = "NULL"
        Exit Function
    End If
    CacheArbolRiesgos_SqlText = "'" & Replace(CStr(p_Value), "'", "''") & "'"
End Function
Private Function CacheArbolRiesgos_SqlLong(p_Value As Variant) As String
    If IsNull(p_Value) Or p_Value = "" Then
        CacheArbolRiesgos_SqlLong = "NULL"
        Exit Function
    End If
    If Not IsNumeric(p_Value) Then
        CacheArbolRiesgos_SqlLong = "NULL"
        Exit Function
    End If
    CacheArbolRiesgos_SqlLong = CStr(CLng(p_Value))
End Function
Public Function CacheArbolRiesgos_AsegurarSchema(Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    On Error GoTo errores
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set db = p_db
    End If
    If Not CacheArbolRiesgos_ExisteTabla(db, "TbCacheArbolRiesgosMeta") Then
        db.Execute "CREATE TABLE TbCacheArbolRiesgosMeta (IDEdicion LONG NOT NULL, ActiveBuildId LONG NOT NULL, UpdatedAt DATETIME);"
        db.Execute "CREATE UNIQUE INDEX UX_TbCacheArbolRiesgosMeta ON TbCacheArbolRiesgosMeta (IDEdicion);"
    End If
    If Not CacheArbolRiesgos_ExisteTabla(db, "TbCacheArbolRiesgosNodo") Then
        db.Execute "CREATE TABLE TbCacheArbolRiesgosNodo (" & _
                    "IDEdicion LONG NOT NULL, " & _
                    "BuildId LONG NOT NULL, " & _
                    "NodeKey TEXT(255) NOT NULL, " & _
                    "ParentKey TEXT(255), " & _
                    "NodeType TEXT(20) NOT NULL, " & _
                    "IDRiesgo LONG, " & _
                    "IDMitigacion LONG, " & _
                    "IDContingencia LONG, " & _
                    "IDAccion LONG, " & _
                    "EsVisibleSinRetirados YESNO, " & _
                    "TextConDescripcion MEMO, " & _
                    "TextSinDescripcion MEMO, " & _
                    "IconName TEXT(255), " & _
                    "ForeColor LONG, " & _
                    "Depth INTEGER, " & _
                    "SortIndex LONG" & _
                    ");"
        db.Execute "CREATE UNIQUE INDEX UX_TbCacheArbolRiesgosNodo ON TbCacheArbolRiesgosNodo (IDEdicion, BuildId, NodeKey);"
        db.Execute "CREATE INDEX IX_TbCacheArbolRiesgosNodo_Parent ON TbCacheArbolRiesgosNodo (IDEdicion, BuildId, ParentKey, SortIndex);"
        db.Execute "CREATE INDEX IX_TbCacheArbolRiesgosNodo_Riesgo ON TbCacheArbolRiesgosNodo (IDEdicion, BuildId, IDRiesgo);"
    End If
    CacheArbolRiesgos_AsegurarSchema = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_AsegurarSchema ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function CacheArbolRiesgos_SqlYesNo(p_Value As Variant) As String
    CacheArbolRiesgos_SqlYesNo = IIf(CBool(p_Value), "True", "False")
End Function
Private Function CacheArbolRiesgos_InsertNodo( _
                                                p_db As DAO.Database, _
                                                p_IDEdicion As Long, _
                                                p_BuildId As Long, _
                                                p_NodeKey As String, _
                                                p_ParentKey As String, _
                                                p_NodeType As String, _
                                                p_IDRiesgo As Variant, _
                                                p_IDMitigacion As Variant, _
                                                p_IDContingencia As Variant, _
                                                p_IDAccion As Variant, _
                                                p_EsVisibleSinRetirados As Boolean, _
                                                p_TextConDescripcion As String, _
                                                p_TextSinDescripcion As String, _
                                                p_IconName As String, _
                                                p_ForeColor As Variant, _
                                                p_Depth As Integer, _
                                                p_SortIndex As Long, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_SQL As String
    On Error GoTo errores
    m_SQL = "INSERT INTO TbCacheArbolRiesgosNodo (IDEdicion, BuildId, NodeKey, ParentKey, NodeType, IDRiesgo, IDMitigacion, IDContingencia, IDAccion, EsVisibleSinRetirados, TextConDescripcion, TextSinDescripcion, IconName, ForeColor, Depth, SortIndex) VALUES (" & _
            p_IDEdicion & ", " & _
            p_BuildId & ", " & _
            CacheArbolRiesgos_SqlText(p_NodeKey) & ", " & _
            CacheArbolRiesgos_SqlText(p_ParentKey) & ", " & _
            CacheArbolRiesgos_SqlText(p_NodeType) & ", " & _
            CacheArbolRiesgos_SqlLong(p_IDRiesgo) & ", " & _
            CacheArbolRiesgos_SqlLong(p_IDMitigacion) & ", " & _
            CacheArbolRiesgos_SqlLong(p_IDContingencia) & ", " & _
            CacheArbolRiesgos_SqlLong(p_IDAccion) & ", " & _
            CacheArbolRiesgos_SqlYesNo(p_EsVisibleSinRetirados) & ", " & _
            CacheArbolRiesgos_SqlText(Left(p_TextConDescripcion, 255)) & ", " & _
            CacheArbolRiesgos_SqlText(Left(p_TextSinDescripcion, 255)) & ", " & _
            CacheArbolRiesgos_SqlText(p_IconName) & ", " & _
            CacheArbolRiesgos_SqlLong(p_ForeColor) & ", " & _
            p_Depth & ", " & _
            p_SortIndex & ");"
    p_db.Execute m_SQL
    CacheArbolRiesgos_InsertNodo = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_InsertNodo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function CacheArbolRiesgos_GetActiveBuildId(p_IDEdicion As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As Long
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    
    CacheArbolRiesgos_AsegurarSchema db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Set rcd = db.OpenRecordset("SELECT ActiveBuildId FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & p_IDEdicion)
    If rcd.EOF Then
        CacheArbolRiesgos_GetActiveBuildId = 0
    Else
        CacheArbolRiesgos_GetActiveBuildId = rcd!ActiveBuildId
    End If
    rcd.Close
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_GetActiveBuildId ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_InvalidarEdicion(p_IDEdicion As Long, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim blnTransaccionPropia As Boolean
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    Set wksLocal = DBEngine.Workspaces(0)
    
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If
    
    CacheArbolRiesgos_AsegurarSchema db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    db.Execute "DELETE * FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & p_IDEdicion
    db.Execute "DELETE * FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_IDEdicion
    
    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If
    
    CacheArbolRiesgos_InvalidarEdicion = "OK"
    Exit Function
errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_InvalidarEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_RebuildEdicion(p_Edicion As Edicion, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim m_NewBuildId As Long
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_IdRiesgo As Variant
    Dim m_Riesgo As riesgo
    Dim m_SortIndex As Long
    Dim m_ColRiesgosActivos As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_Edicion Is Nothing Then Exit Function
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    Set wksLocal = DBEngine.Workspaces(0)
    
    CacheArbolRiesgos_AsegurarSchema db, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    wksLocal.BeginTrans
    
    ' Clear existing cache for this edition
    db.Execute "DELETE * FROM TbCacheArbolRiesgosMeta WHERE IDEdicion=" & p_Edicion.IDEdicion
    db.Execute "DELETE * FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_Edicion.IDEdicion
    
    m_NewBuildId = 1 ' Simple versioning for now, could be incremented
    
    db.Execute "INSERT INTO TbCacheArbolRiesgosMeta (IDEdicion, ActiveBuildId, UpdatedAt) VALUES (" & p_Edicion.IDEdicion & ", " & m_NewBuildId & ", Now())"
    
    ' Insert root node for Edition
    CacheArbolRiesgos_InsertNodo db, CLng(p_Edicion.IDEdicion), m_NewBuildId, "EDICION|" & p_Edicion.IDEdicion, "", "EDICION", Null, Null, Null, Null, True, p_Edicion.NombreParaNodo, p_Edicion.NombreParaNodo, FSO.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32), Null, 0, 0, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Set m_ColRiesgos = p_Edicion.colRiesgos
    Set m_ColRiesgosActivos = p_Edicion.colRiesgosActivos
    
    If Not m_ColRiesgos Is Nothing Then
        m_SortIndex = 0
        For Each m_IdRiesgo In m_ColRiesgos
            m_SortIndex = m_SortIndex + 1
            Set m_Riesgo = m_ColRiesgos(m_IdRiesgo)
            
            CacheArbolRiesgos_InsertarRiesgoCompleto db, CLng(p_Edicion.IDEdicion), m_NewBuildId, m_Riesgo, m_SortIndex, m_ColRiesgosActivos, p_Error
            If p_Error <> "" Then Err.Raise 1000
        Next
    End If
    
    wksLocal.CommitTrans
    
    CacheArbolRiesgos_RebuildEdicion = "OK"
    Exit Function
errores:
    wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_RebuildEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function CacheArbolRiesgos_InsertarRiesgoCompleto( _
                                                p_db As DAO.Database, _
                                                p_IDEdicion As Long, _
                                                p_BuildId As Long, _
                                                p_Riesgo As riesgo, _
                                                p_SortIndex As Long, _
                                                p_ColRiesgosActivos As Scripting.Dictionary, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_IdRiesgo As Variant
    Dim m_EsVisibleSinRetirados As Boolean
    Dim m_Color As String
    Dim m_ForeColor As Variant
    Dim m_TextConDescripcion As String
    Dim m_TextSinDescripcion As String
    Dim m_IconName As String
    Dim m_HayError As String
    Dim m_ColPMs As Scripting.Dictionary
    Dim m_PM As PM
    Dim m_IDPM As Variant
    Dim m_SortPlan As Long
    Dim m_ColAcciones As Scripting.Dictionary
    Dim m_SortAccion As Long
    Dim m_PMA As PMAccion
    Dim m_IdPMA As Variant
    Dim m_ColPCs As Scripting.Dictionary
    Dim m_PC As PC
    Dim m_IDPC As Variant
    Dim m_PCA As PCAccion
    Dim m_IdPCA As Variant
    Dim m_KeyEdicion As String
    
    On Error GoTo errores
    
    m_IdRiesgo = p_Riesgo.IDRiesgo
    m_KeyEdicion = "EDICION|" & p_IDEdicion
    
    m_EsVisibleSinRetirados = False
    If Not p_ColRiesgosActivos Is Nothing Then
        If p_ColRiesgosActivos.Exists(m_IdRiesgo) Then
            m_EsVisibleSinRetirados = True
        End If
    Else
        m_EsVisibleSinRetirados = True
    End If
    
    m_TextConDescripcion = p_Riesgo.NombreNodoDesc
    If m_TextConDescripcion = "" Then
        m_TextConDescripcion = p_Riesgo.NombreNodoDescCalculado
        If p_Riesgo.Error <> "" Then
            p_Error = p_Riesgo.Error
            Err.Raise 1000
        End If
    End If
    m_TextSinDescripcion = p_Riesgo.NombreNodoEstado
    If m_TextSinDescripcion = "" Then
        m_TextSinDescripcion = p_Riesgo.NombreNodoEstadoCalculado
        If p_Riesgo.Error <> "" Then
            p_Error = p_Riesgo.Error
            Err.Raise 1000
        End If
    End If
    m_IconName = p_Riesgo.NombreIcono
    If m_IconName = "" Then
        m_IconName = p_Riesgo.NombreIconoCalculado(p_db:=p_db, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_Riesgo.EsActivo = EnumSiNo.No Then
        m_Color = "Negro"
    Else
        m_HayError = p_Riesgo.CalcularHayErrorEnRiesgo(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        
        If Left$(UCase$(Trim$(m_HayError)), 1) = "S" Then
            m_Color = "Rojo"
        Else
            m_Color = "Negro"
        End If
    End If
    If m_Color = "Negro" Then
        m_ForeColor = vbBlack
    ElseIf m_Color = "Rojo" Then
        m_ForeColor = vbRed
    Else
        m_ForeColor = Null
    End If
    
    CacheArbolRiesgos_InsertNodo p_db, p_IDEdicion, p_BuildId, "RIESGO|" & m_IdRiesgo, m_KeyEdicion, "RIESGO", m_IdRiesgo, Null, Null, Null, m_EsVisibleSinRetirados, m_TextConDescripcion, m_TextSinDescripcion, m_IconName, m_ForeColor, 1, p_SortIndex, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    Set m_ColPMs = p_Riesgo.ColPMs
    If Not m_ColPMs Is Nothing Then
        m_SortPlan = 0
        For Each m_IDPM In m_ColPMs
            m_SortPlan = m_SortPlan + 1
            Set m_PM = m_ColPMs(m_IDPM)
            m_IconName = m_PM.NombreIconoCalculado
            CacheArbolRiesgos_InsertNodo p_db, p_IDEdicion, p_BuildId, "MITIGACION|" & m_IDPM, "RIESGO|" & m_IdRiesgo, "MITIGACION", m_IdRiesgo, m_IDPM, Null, Null, m_EsVisibleSinRetirados, m_IconName, m_IconName, m_IconName, Null, 2, m_SortPlan, p_Error
            If p_Error <> "" Then Err.Raise 1000
            
            Set m_ColAcciones = m_PM.colAcciones
            If Not m_ColAcciones Is Nothing Then
                m_SortAccion = 0
                For Each m_IdPMA In m_ColAcciones
                    m_SortAccion = m_SortAccion + 1
                    Set m_PMA = m_ColAcciones(m_IdPMA)
                    m_IconName = m_PMA.NombreIconoCalculado
                    CacheArbolRiesgos_InsertNodo p_db, p_IDEdicion, p_BuildId, "ACCIONMITIGACION|" & m_IdPMA, "MITIGACION|" & m_IDPM, "ACCIONMITIGACION", m_IdRiesgo, m_IDPM, Null, m_IdPMA, m_EsVisibleSinRetirados, m_IconName, m_IconName, m_IconName, Null, 3, m_SortAccion, p_Error
                    If p_Error <> "" Then Err.Raise 1000
                Next
            End If
        Next
    End If
    
    Set m_ColPCs = p_Riesgo.ColPCs
    If Not m_ColPCs Is Nothing Then
        m_SortPlan = 0
        For Each m_IDPC In m_ColPCs
            m_SortPlan = m_SortPlan + 1
            Set m_PC = m_ColPCs(m_IDPC)
            m_IconName = m_PC.NombreIconoCalculado
            CacheArbolRiesgos_InsertNodo p_db, p_IDEdicion, p_BuildId, "CONTINGENCIA|" & m_IDPC, "RIESGO|" & m_IdRiesgo, "CONTINGENCIA", m_IdRiesgo, Null, m_IDPC, Null, m_EsVisibleSinRetirados, m_IconName, m_IconName, m_IconName, Null, 2, m_SortPlan, p_Error
            If p_Error <> "" Then Err.Raise 1000
            
            Set m_ColAcciones = m_PC.colAcciones
            If Not m_ColAcciones Is Nothing Then
                m_SortAccion = 0
                For Each m_IdPCA In m_ColAcciones
                    m_SortAccion = m_SortAccion + 1
                    Set m_PCA = m_ColAcciones(m_IdPCA)
                    m_IconName = m_PCA.NombreIconoCalculado
                    CacheArbolRiesgos_InsertNodo p_db, p_IDEdicion, p_BuildId, "ACCIONCONTINGENCIA|" & m_IdPCA, "CONTINGENCIA|" & m_IDPC, "ACCIONCONTINGENCIA", m_IdRiesgo, Null, m_IDPC, m_IdPCA, m_EsVisibleSinRetirados, m_IconName, m_IconName, m_IconName, Null, 3, m_SortAccion, p_Error
                    If p_Error <> "" Then Err.Raise 1000
                Next
            End If
        Next
    End If
    
    CacheArbolRiesgos_InsertarRiesgoCompleto = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_InsertarRiesgoCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_ActualizarRiesgo(p_Riesgo As riesgo, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim m_ActiveBuildId As Long
    Dim m_SortIndex As Long
    Dim rcd As DAO.Recordset
    Dim m_ColRiesgosActivos As Scripting.Dictionary
    Dim m_Edicion As Edicion
    Dim blnTransaccionPropia As Boolean
    
    On Error GoTo errores
    
    If p_Riesgo Is Nothing Then Exit Function
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    
    Set wksLocal = DBEngine.Workspaces(0)
    
    m_ActiveBuildId = CacheArbolRiesgos_GetActiveBuildId(CLng(p_Riesgo.IDEdicion), db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If m_ActiveBuildId = 0 Then
        ' No cache exists, rebuild all
        Set m_Edicion = getEdicion(p_Riesgo.IDEdicion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        
        CacheArbolRiesgos_ActualizarRiesgo = CacheArbolRiesgos_RebuildEdicion(m_Edicion, db, p_Error)
        Exit Function
    End If
    
    ' Find sort index
    Set rcd = db.OpenRecordset("SELECT SortIndex FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_Riesgo.IDEdicion & " AND BuildId=" & m_ActiveBuildId & " AND IDRiesgo=" & p_Riesgo.IDRiesgo & " AND NodeType='RIESGO'")
    If rcd.EOF Then
        ' Risk not in cache, rebuild all
        rcd.Close
        Set m_Edicion = getEdicion(p_Riesgo.IDEdicion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
        CacheArbolRiesgos_ActualizarRiesgo = CacheArbolRiesgos_RebuildEdicion(m_Edicion, db, p_Error)
        Exit Function
    End If
    m_SortIndex = rcd!SortIndex
    rcd.Close
    
    Set m_Edicion = getEdicion(p_Riesgo.IDEdicion, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Set m_ColRiesgosActivos = m_Edicion.colRiesgosActivos
    
    ' Solo iniciamos transacción si no nos han pasado una base de datos (se asume que si la pasan, el llamador gestiona la trans.)
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If
    
    ' Delete old nodes for this risk
    db.Execute "DELETE * FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_Riesgo.IDEdicion & " AND BuildId=" & m_ActiveBuildId & " AND IDRiesgo=" & p_Riesgo.IDRiesgo
    
    ' Insert new nodes
    CacheArbolRiesgos_InsertarRiesgoCompleto db, CLng(p_Riesgo.IDEdicion), m_ActiveBuildId, p_Riesgo, m_SortIndex, m_ColRiesgosActivos, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' Update UpdatedAt in Meta
    db.Execute "UPDATE TbCacheArbolRiesgosMeta SET UpdatedAt=Now() WHERE IDEdicion=" & p_Riesgo.IDEdicion
    
    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If
    
    CacheArbolRiesgos_ActualizarRiesgo = "OK"
    Exit Function
errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_ActualizarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_BorrarRiesgo(p_Riesgo As riesgo, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim m_ActiveBuildId As Long
    Dim blnTransaccionPropia As Boolean
    
    On Error GoTo errores
    
    If p_Riesgo Is Nothing Then Exit Function
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    
    Set wksLocal = DBEngine.Workspaces(0)
    
    m_ActiveBuildId = CacheArbolRiesgos_GetActiveBuildId(CLng(p_Riesgo.IDEdicion), db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If m_ActiveBuildId = 0 Then
        ' Cache not active, nothing to do
        CacheArbolRiesgos_BorrarRiesgo = "OK"
        Exit Function
    End If
    
    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If
    
    ' Delete nodes for this risk
    db.Execute "DELETE * FROM TbCacheArbolRiesgosNodo WHERE IDEdicion=" & p_Riesgo.IDEdicion & " AND BuildId=" & m_ActiveBuildId & " AND IDRiesgo=" & p_Riesgo.IDRiesgo
    
    ' Update UpdatedAt in Meta
    db.Execute "UPDATE TbCacheArbolRiesgosMeta SET UpdatedAt=Now() WHERE IDEdicion=" & p_Riesgo.IDEdicion
    
    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If
    
    CacheArbolRiesgos_BorrarRiesgo = "OK"
    Exit Function
errores:
    If blnTransaccionPropia Then wksLocal.Rollback
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_BorrarRiesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgos_ActualizarEdicion(p_Edicion As Edicion, Optional ByVal p_db As DAO.Database, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim m_ActiveBuildId As Long
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_Edicion Is Nothing Then Exit Function
    
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    
    m_ActiveBuildId = CacheArbolRiesgos_GetActiveBuildId(CLng(p_Edicion.IDEdicion), db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_ActiveBuildId = 0 Then
        ' No cache, do nothing or full rebuild if you prefer. For now, just exit.
        Exit Function
    End If
    
    ' Update the root node (EDICION|...)
    m_SQL = "UPDATE TbCacheArbolRiesgosNodo SET " & _
            "TextConDescripcion = " & CacheArbolRiesgos_SqlText(p_Edicion.NombreParaNodo) & ", " & _
            "TextSinDescripcion = " & CacheArbolRiesgos_SqlText(p_Edicion.NombreParaNodo) & ", " & _
            "IconName = " & CacheArbolRiesgos_SqlText(FSO.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)) & " " & _
            "WHERE IDEdicion = " & p_Edicion.IDEdicion & " AND BuildId = " & m_ActiveBuildId & " AND NodeKey = " & CacheArbolRiesgos_SqlText("EDICION|" & p_Edicion.IDEdicion) & ";"
    db.Execute m_SQL
    
    db.Execute "UPDATE TbCacheArbolRiesgosMeta SET UpdatedAt=Now() WHERE IDEdicion=" & p_Edicion.IDEdicion
    
    CacheArbolRiesgos_ActualizarEdicion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_ActualizarEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_CargarEnTreeView( _
                                                p_Arbol As MSComctlLib.TreeView, _
                                                p_ListImages As MSComctlLib.ImageList, _
                                                p_Edicion As Edicion, _
                                                p_VerRetirados As String, _
                                                p_VerDescripcion As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim m_BuildId As Long
    Dim m_Key As String
    Dim m_ParentKey As String
    Dim m_Text As String
    Dim m_Icon As String
    Dim m_Color As Variant
    Dim m_VisibleSinRetirados As Boolean
    Dim m_BlnVerRetirados As Boolean
    Dim m_BlnVerDescripcion As Boolean
    
    On Error GoTo errores
    
    CacheArbolRiesgos_CargarEnTreeView = EnumSiNo.No
    
    If p_Arbol Is Nothing Or p_Edicion Is Nothing Then Exit Function
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    m_BuildId = CacheArbolRiesgos_GetActiveBuildId(CLng(p_Edicion.IDEdicion), db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_BuildId <= 0 Then Exit Function
    
    m_BlnVerRetirados = (Left$(UCase$(Trim$(Nz(p_VerRetirados, "No"))), 1) = "S")
    m_BlnVerDescripcion = (Left$(UCase$(Trim$(Nz(p_VerDescripcion, "Sí"))), 1) = "S")
    
    ' Query nodes ordered by depth and sort index
    m_SQL = "SELECT * FROM TbCacheArbolRiesgosNodo " & _
            "WHERE IDEdicion=" & p_Edicion.IDEdicion & " AND BuildId=" & m_BuildId & " " & _
            "ORDER BY Depth, SortIndex;"
    
    Set rcd = db.OpenRecordset(m_SQL)
    p_Arbol.Nodes.Clear
    
    Do While Not rcd.EOF
        m_VisibleSinRetirados = rcd!EsVisibleSinRetirados
        
        ' Filter logic
        If m_BlnVerRetirados Or m_VisibleSinRetirados Then
            m_Key = rcd!NodeKey
            m_ParentKey = Nz(rcd!ParentKey, "")
            m_Icon = Nz(rcd!IconName, "")
            m_Color = rcd!ForeColor
            
            If m_BlnVerDescripcion Then
                m_Text = Nz(rcd!TextConDescripcion, "")
            Else
                m_Text = Nz(rcd!TextSinDescripcion, "")
            End If
            
            Dim m_Node As MSComctlLib.Node
            If m_ParentKey = "" Then
                Set m_Node = p_Arbol.Nodes.Add(, , m_Key, m_Text, m_Icon)
            Else
                On Error Resume Next
                Set m_Node = p_Arbol.Nodes.Add(m_ParentKey, tvwChild, m_Key, m_Text, m_Icon)
                On Error GoTo errores
            End If
            
            If Not m_Node Is Nothing Then
                If Not IsNull(m_Color) Then
                    If CLng(m_Color) <> 0 Then m_Node.ForeColor = CLng(m_Color)
                End If
            End If
        End If
        rcd.MoveNext
    Loop
    
    rcd.Close
    CacheArbolRiesgos_CargarEnTreeView = EnumSiNo.Sí
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en CacheArbolRiesgos_CargarEnTreeView: " & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_AplicarFiltroRetiradosEnTreeView( _
                                                                p_Arbol As MSComctlLib.TreeView, _
                                                                p_Edicion As Edicion, _
                                                                p_VerRetirados As String, _
                                                                p_VerDescripcion As String, _
                                                                p_ListImages As MSComctlLib.ImageList, _
                                                                Optional ByRef p_Error As String _
                                                                ) As EnumSiNo
    ' Si cambia el filtro de retirados, lo más seguro es recargar desde caché
    ' ya que algunos nodos pueden aparecer o desaparecer por completo.
    CacheArbolRiesgos_AplicarFiltroRetiradosEnTreeView = CacheArbolRiesgos_CargarEnTreeView(p_Arbol, p_ListImages, p_Edicion, p_VerRetirados, p_VerDescripcion, p_Error)
End Function
Public Function CacheArbolRiesgos_ActualizarTitulosRiesgosEnTreeView( _
                                                                    p_Arbol As MSComctlLib.TreeView, _
                                                                    p_Edicion As Edicion, _
                                                                    p_VerDescripcion As String, _
                                                                    Optional ByVal p_db As DAO.Database = Nothing, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As EnumSiNo
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim m_BuildId As Long
    Dim m_VerDescripcionTxt As String
    Dim m_BlnVerDescripcion As Boolean
    Dim m_Key As String
    Dim m_Text As String
    Dim m_ForeColor As Variant
    Dim m_Node As MSComctlLib.Node
    On Error GoTo errores
    CacheArbolRiesgos_ActualizarTitulosRiesgosEnTreeView = EnumSiNo.No
    p_Error = ""
    If p_Arbol Is Nothing Or p_Edicion Is Nothing Then Exit Function
    If p_Arbol.Nodes.Count = 0 Then Exit Function
    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set db = p_db
    End If
    m_BuildId = CacheArbolRiesgos_GetActiveBuildId(CLng(p_Edicion.IDEdicion), db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_BuildId <= 0 Then Exit Function
    m_VerDescripcionTxt = Nz(p_VerDescripcion, "Sí")
    m_BlnVerDescripcion = (Left$(UCase$(Trim$(m_VerDescripcionTxt)), 1) = "S")
    m_SQL = "SELECT NodeKey, TextConDescripcion, TextSinDescripcion, ForeColor " & _
            "FROM TbCacheArbolRiesgosNodo " & _
            "WHERE IDEdicion=" & CLng(p_Edicion.IDEdicion) & " AND BuildId=" & CLng(m_BuildId) & " AND NodeType='RIESGO';"
    Set rcd = db.OpenRecordset(m_SQL)
    Do While Not rcd.EOF
        m_Key = Nz(rcd.Fields("NodeKey"), "")
        If m_Key <> "" Then
            If m_BlnVerDescripcion Then
                m_Text = Nz(rcd.Fields("TextConDescripcion"), "")
            Else
                m_Text = Nz(rcd.Fields("TextSinDescripcion"), "")
            End If
            m_ForeColor = rcd.Fields("ForeColor")
            On Error Resume Next
            Set m_Node = p_Arbol.Nodes(m_Key)
            On Error GoTo errores
            If Not m_Node Is Nothing Then
                If m_Text <> "" Then m_Node.Text = m_Text
                If Not IsNull(m_ForeColor) Then
                    If Nz(m_ForeColor, 0) <> 0 Then m_Node.ForeColor = CLng(m_ForeColor)
                End If
            End If
            Set m_Node = Nothing
        End If
        rcd.MoveNext
    Loop
    rcd.Close
    CacheArbolRiesgos_ActualizarTitulosRiesgosEnTreeView = EnumSiNo.Sí
    Exit Function
errores:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_ActualizarTitulosRiesgosEnTreeView ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function CacheArbolRiesgos_ConstruirColPriorizacion( _
                                                            p_IDEdicion As Long, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Scripting.Dictionary
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_ColPriorizacion As Scripting.Dictionary
    Dim m_Id As Variant
    Dim m_Riesgo As riesgo

    On Error GoTo errores

    If p_IDEdicion <= 0 Then
        Exit Function
    End If

    Set m_ColRiesgos = Constructor.getRiesgosPorEdicion(CStr(p_IDEdicion), EnumSiNo.No, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColRiesgos Is Nothing Then
        Exit Function
    End If

    Set m_ColPriorizacion = New Scripting.Dictionary
    m_ColPriorizacion.CompareMode = TextCompare

    For Each m_Id In m_ColRiesgos
        Set m_Riesgo = m_ColRiesgos(m_Id)
        If Not m_Riesgo Is Nothing Then
            If Not m_ColPriorizacion.Exists(CStr(m_Riesgo.IDRiesgo)) Then
                m_ColPriorizacion.Add CStr(m_Riesgo.IDRiesgo), m_Riesgo.Priorizacion
            End If
        End If
    Next

    Set CacheArbolRiesgos_ConstruirColPriorizacion = m_ColPriorizacion
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo CacheArbolRiesgos_ConstruirColPriorizacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function CacheArbolRiesgos_ActualizarOrdenRiesgos( _
                                                            p_Edicion As Edicion, _
                                                            p_ColPriorizacion As Scripting.Dictionary, _
                                                            Optional ByVal p_db As DAO.Database, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim db As DAO.Database
    Dim wksLocal As DAO.Workspace
    Dim m_BuildId As Long
    Dim m_Id As Variant
    Dim m_Pri As Variant
    Dim m_SQL As String
    Dim m_Ids() As Variant

    Dim m_Prioridades() As Variant

    Dim m_Index As Long
    Dim blnTransaccionPropia As Boolean

    On Error GoTo errores

    CacheArbolRiesgos_ActualizarOrdenRiesgos = "OK"
    p_Error = ""

    If p_Edicion Is Nothing Then
        Exit Function
    End If
    If p_ColPriorizacion Is Nothing Then
        Exit Function
    End If

    If CacheArbolRiesgos_UsarCache(p_Error) = EnumSiNo.No Then
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If

    CacheArbolRiesgos_AsegurarSchema p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If p_db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set db = p_db
    End If
    Set wksLocal = DBEngine.Workspaces(0)

    m_BuildId = CacheArbolRiesgos_GetActiveBuildId(p_IDEdicion:=CLng(p_Edicion.IDEdicion), p_db:=db, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_BuildId <= 0 Then
        Exit Function
    End If

    If p_db Is Nothing Then
        wksLocal.BeginTrans
        blnTransaccionPropia = True
    End If

    If p_ColPriorizacion.Count > 0 Then

        ReDim m_Ids(1 To p_ColPriorizacion.Count)

        ReDim m_Prioridades(1 To p_ColPriorizacion.Count)

        m_Index = 0

        For Each m_Id In p_ColPriorizacion

            m_Index = m_Index + 1

            m_Ids(m_Index) = m_Id

            m_Prioridades(m_Index) = p_ColPriorizacion(m_Id)

        Next

        m_Ids = CacheArbolRiesgos_OrdenarIdsPorPrioridad(m_Ids, m_Prioridades)

        For m_Index = LBound(m_Ids) To UBound(m_Ids)

            m_SQL = "UPDATE TbCacheArbolRiesgosNodo SET SortIndex=" & _
                    CStr(m_Index - LBound(m_Ids) + 1) & _
                    " WHERE IDEdicion=" & CLng(p_Edicion.IDEdicion) & _
                    " AND BuildId=" & CLng(m_BuildId) & _
                    " AND NodeKey=" & CacheArbolRiesgos_SqlText("RIESGO|" & m_Ids(m_Index)) & ";"

            db.Execute m_SQL

        Next

    End If

    If blnTransaccionPropia Then
        wksLocal.CommitTrans
    End If
    Exit Function
errores:
    If blnTransaccionPropia Then
        On Error Resume Next
        wksLocal.Rollback
        On Error GoTo 0
    End If
    If Err.Number <> 1000 Then
        p_Error = "El método CacheArbolRiesgos_ActualizarOrdenRiesgos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function CacheArbolRiesgos_OrdenarIdsPorPrioridad( _
                                                            p_Ids As Variant, _
                                                            p_Prioridades As Variant _
                                                            ) As Variant
    Dim i As Long
    Dim j As Long
    Dim m_TempId As Variant
    Dim m_TempPri As Variant

    If IsEmpty(p_Ids) Then
        Exit Function
    End If

    For i = LBound(p_Ids) To UBound(p_Ids) - 1
        For j = i + 1 To UBound(p_Ids)
            If CacheArbolRiesgos_CompareOrden(p_Ids(i), p_Prioridades(i), p_Ids(j), p_Prioridades(j)) > 0 Then
                m_TempId = p_Ids(i)
                p_Ids(i) = p_Ids(j)
                p_Ids(j) = m_TempId

                m_TempPri = p_Prioridades(i)
                p_Prioridades(i) = p_Prioridades(j)
                p_Prioridades(j) = m_TempPri
            End If
        Next
    Next

    CacheArbolRiesgos_OrdenarIdsPorPrioridad = p_Ids
End Function
Private Function CacheArbolRiesgos_CompareOrden( _
                                                    p_IdA As Variant, _
                                                    p_PriA As Variant, _
                                                    p_IdB As Variant, _
                                                    p_PriB As Variant _
                                                    ) As Long
    Dim m_TieneA As Boolean
    Dim m_TieneB As Boolean
    Dim m_IdALong As Long
    Dim m_IdBLong As Long

    m_TieneA = (Not IsNull(p_PriA)) And IsNumeric(p_PriA)
    m_TieneB = (Not IsNull(p_PriB)) And IsNumeric(p_PriB)

    m_IdALong = CLng(p_IdA)
    m_IdBLong = CLng(p_IdB)

    If m_TieneA And Not m_TieneB Then
        CacheArbolRiesgos_CompareOrden = -1
        Exit Function
    End If
    If Not m_TieneA And m_TieneB Then
        CacheArbolRiesgos_CompareOrden = 1
        Exit Function
    End If

    If m_TieneA And m_TieneB Then
        If CLng(p_PriA) < CLng(p_PriB) Then
            CacheArbolRiesgos_CompareOrden = -1
            Exit Function
        ElseIf CLng(p_PriA) > CLng(p_PriB) Then
            CacheArbolRiesgos_CompareOrden = 1
            Exit Function
        End If
    End If

    If m_IdALong < m_IdBLong Then
        CacheArbolRiesgos_CompareOrden = -1
    ElseIf m_IdALong > m_IdBLong Then
        CacheArbolRiesgos_CompareOrden = 1
    Else
        CacheArbolRiesgos_CompareOrden = 0
    End If
End Function

