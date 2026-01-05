Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : DatabaseProvider
' Purpose   : Centraliza el acceso a las bases de datos respetando la configuración
'             establecida por la función EVE (Local vs Remoto) y las TempVars.
'---------------------------------------------------------------------------------------

Public Function GetGestionDB(Optional ByRef p_Error As String) As DAO.Database
    Dim m_URL As String
    Dim m_NombreDatos As String
    Dim db As DAO.Database
    Dim wks As DAO.Workspace
    
    On Error GoTo errores
    
    m_NombreDatos = "Gestion_Riesgos_Datos.accdb"
    
    ' RESPETO TOTAL A LA LÓGICA EVE
    If Application.TempVars("EnPruebas") = "Sí" Then
        If Application.TempVars("DatosEnLocal") = "Sí" Then
            ' m_URLRutaAplicacionesLocal se rellena en EVE
            m_URL = Application.TempVars("URLRutaAplicacionesLocal") & "000datoslocal\" & m_NombreDatos
        Else
            ' m_URLRutaAplicacionRemota se rellena en EVE
            m_URL = Application.TempVars("URLRutaAplicacionRemota") & m_NombreDatos
        End If
    Else
        If Application.TempVars("DatosEnLocal") = "Sí" Then
            m_URL = Application.TempVars("URLRutaAplicacionesLocal") & "000datoslocal\" & m_NombreDatos
        Else
            m_URL = Application.TempVars("URLRutaAplicacionRemota") & m_NombreDatos
        End If
    End If
    
    Set wks = DBEngine.Workspaces(0)
    ' Usamos la clave de paso que tienes definida en el legacy
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=dpddpd")
    Set GetGestionDB = db
    
    Exit Function

errores:
    p_Error = "Error en DatabaseProvider.GetGestionDB: " & Err.Description
End Function

Public Function GetExpedientesDB(Optional ByRef p_Error As String) As DAO.Database
    Dim m_URL As String
    Dim db As DAO.Database
    Dim wks As DAO.Workspace
    
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = Application.TempVars("URLRutaAplicacionesLocal") & "000datoslocal\Expedientes_datos.accdb"
    Else
        m_URL = Application.TempVars("URLRutaAplicacionesRemotas") & "EXPEDIENTES\Expedientes_datos.accdb"
    End If
    
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=dpddpd")
    Set GetExpedientesDB = db
    Exit Function

errores:
    p_Error = "Error en DatabaseProvider.GetExpedientesDB: " & Err.Description
End Function
