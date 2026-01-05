Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : InfrastructureService
' Purpose   : Lógica de detección de rutas del sistema (OneDrive, Red, etc.)
'---------------------------------------------------------------------------------------

Public Function GetRutaAplicacionesLocal(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim m_Ruta As String
    
    On Error GoTo errores
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Intentar OneDrive Telefónica
    m_Ruta = GetDirectorioOneDriveTelefonicaApps(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If m_Ruta <> "" And FSO.FolderExists(m_Ruta) Then
        GetRutaAplicacionesLocal = m_Ruta
        Set FSO = Nothing
        Exit Function
    End If
    
    ' 2. Intentar OneDrive Estándar
    m_Ruta = GetDirectorioOneDrive(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    GetRutaAplicacionesLocal = m_Ruta
    Set FSO = Nothing
    Exit Function

errores:
    p_Error = "Error en GetRutaAplicacionesLocal: " & Err.Description
End Function

Private Function GetDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim FSO As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    On Error GoTo errores
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set carpetaRaiz = FSO.GetFolder("C:\")
    
    ' Buscar carpeta OneDrive en C:\
    For Each subCarpeta In carpetaRaiz.SubFolders
        If InStr(1, subCarpeta.Name, "OneDrive", vbTextCompare) > 0 Then
            GetDirectorioOneDrive = subCarpeta.Path
            Exit Function
        End If
    Next
    
    ' Si no está en C:\, usar variable de entorno
    GetDirectorioOneDrive = Environ("OneDrive")
    Exit Function
errores:
    p_Error = "Error en GetDirectorioOneDrive: " & Err.Description
End Function

Private Function GetDirectorioOneDriveTelefonicaApps(Optional ByRef p_Error As String) As String
    Dim m_RutaBase As String
    On Error GoTo errores
    
    m_RutaBase = GetDirectorioOneDrive(p_Error)
    If m_RutaBase <> "" Then
        GetDirectorioOneDriveTelefonicaApps = m_RutaBase & "\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    End If
    Exit Function
errores:
    p_Error = "Error en GetDirectorioOneDriveTelefonicaApps: " & Err.Description
End Function
