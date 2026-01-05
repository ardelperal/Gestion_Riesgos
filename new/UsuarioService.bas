Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : UsuarioService
' Purpose   : Lógica de negocio para perfiles y permisos de usuario.
'---------------------------------------------------------------------------------------

Public m_ObjUsuarioConectado As Usuario

Public Function GetUsuarioActual(Optional ByVal p_CorreoManual As String, Optional ByRef p_Error As String) As Usuario
    Dim m_UsuarioRed As String
    Dim m_Usr As Usuario
    Dim objNetwork As Object
    
    On Error GoTo errores
    
    ' 1. Obtener ID de Red o Correo
    If p_CorreoManual <> "" Then
        ' Lógica para buscar por correo si fuera necesario
    Else
        Set objNetwork = CreateObject("Wscript.Network")
        m_UsuarioRed = objNetwork.UserName
        ' Normalización de usuarios de desarrollo
        If m_UsuarioRed = "adm1" Or m_UsuarioRed = "Local1" Then m_UsuarioRed = "adm"
    End If
    
    ' 2. Cargar datos básicos desde el Repositorio
    Set m_Usr = UsuarioRepository.GetByUsuarioRed(m_UsuarioRed, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If m_Usr Is Nothing Then
        p_Error = "Usuario '" & m_UsuarioRed & "' no reconocido en el sistema."
        Err.Raise 1000
    End If
    
    ' 3. Determinar Perfil mediante consulta a la Lanzadera
    Dim m_IDApp As Integer
    m_IDApp = CInt(m_ObjEntorno.IDAplicacion)
    
    m_Usr.EsAdministrador = UsuarioRepository.EsAdministrador(m_Usr.UsuarioRed, m_IDApp, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    m_Usr.EsCalidad = UsuarioRepository.EsUsuarioCalidad(m_Usr.UsuarioRed, m_IDApp, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ' Si no es admin ni calidad, por defecto es Técnico
    If Not m_Usr.EsAdministrador And Not m_Usr.EsCalidad Then
        m_Usr.EsTecnico = True
    End If
    
    ' 4. Cargar Permisos detallados si existen
    ' Set m_Usr.Permisos = UsuarioRepository.GetPermisos(m_Usr.ID, m_IDApp, p_Error)
    
    Set GetUsuarioActual = m_Usr
    Set m_ObjUsuarioConectado = m_Usr
    
    Exit Function

errores:
    p_Error = "Error en UsuarioService.GetUsuarioActual: " & Err.Description
End Function
