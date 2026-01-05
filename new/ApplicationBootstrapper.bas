Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : ApplicationBootstrapper
' Purpose   : Punto de entrada único del sistema. Inicializa el entorno y la seguridad.
'---------------------------------------------------------------------------------------

Public m_ObjEntorno As Entorno ' Objeto Global de Configuración

Public Function EVE( _
                    Optional ByRef p_CorreoUsuario As String, _
                    Optional ByRef p_Error As String _
                    ) As String


    Dim m_NombreCarpeta As String
    On Error GoTo errores
    
    ' 1. INICIALIZACIÓN DE VARIABLES DE ESTADO (TempVars para compatibilidad global)
    Application.TempVars.RemoveAll
    Application.TempVars("CadenaJerarquicaModelo") = "nuevo"
    Application.TempVars("JPMesesAvisoEntreEdiciones") = 3
    Application.TempVars("JPDiasPreviosParaElAviso") = 15
    Application.TempVars("CalDiaInicialMesAviso") = 2
    Application.TempVars("GenerarInformeTipo") = "Excel"
    
    ' Entorno Local/Remoto (Ajustar según necesidad del desarrollador)
    Application.TempVars("DatosEnLocal") = "No" 
    Application.TempVars("EnDesarrollo") = "No"
    Application.TempVars("EnPruebas") = "No"
    
    ' 2. INSTANCIACIÓN DEL OBJETO ENTORNO (Dominio)
    Set m_ObjEntorno = New Entorno
    With m_ObjEntorno
        .DatosEnLocal = Application.TempVars("DatosEnLocal")
        .EnDesarrollo = Application.TempVars("EnDesarrollo")
        .EnPruebas = Application.TempVars("EnPruebas")
        .IDAplicacion = IIf(.EnPruebas = "Sí", "51", "5")
        
        ' Rutas Remotas (Fieles al legacy)
        .RutaAplicacionesRemotas = "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
        m_NombreCarpeta = IIf(.EnPruebas = "Sí", "GESTION RIESGOS PRUEBA", "GESTION RIESGOS")
        .RutaAplicacionRemota = .RutaAplicacionesRemotas & m_NombreCarpeta & "\"
        
        ' Rutas Locales
        If .DatosEnLocal = "Sí" Then
            .RutaAplicacionesLocal = InfrastructureService.GetRutaAplicacionesLocal(p_Error)
            If p_Error <> "" Then Err.Raise 1000
            
            If .RutaAplicacionesLocal <> "" Then
                .RutaAplicacionLocal = .RutaAplicacionesLocal & m_NombreCarpeta & "\"
            End If
        End If
        
        ' Guardar rutas en TempVars para que DatabaseProvider las vea
        Application.TempVars("URLRutaAplicacionesLocal") = .RutaAplicacionesLocal
        Application.TempVars("URLRutaAplicacionRemota") = .RutaAplicacionRemota
        Application.TempVars("URLRutaAplicacionesRemotas") = .RutaAplicacionesRemotas
        
        ' Parámetros de Negocio
        .JPMesesAvisoEntreEdiciones = Application.TempVars("JPMesesAvisoEntreEdiciones")
        .JPDiasPreviosParaElAviso = Application.TempVars("JPDiasPreviosParaElAviso")
        .CalDiaInicialMesAviso = Application.TempVars("CalDiaInicialMesAviso")
        .GenerarInformeTipo = Application.TempVars("GenerarInformeTipo")
    End With
    
    ' 3. CARGA DE USUARIO CONECTADO
    ' El Servicio de Usuarios centraliza la autenticación y permisos
    Set UsuarioService.m_ObjUsuarioConectado = UsuarioService.GetUsuarioActual(p_CorreoUsuario, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ' Sincronizar banderas heredadas si es necesario para compatibilidad UI
    Application.TempVars("EsAdministrador") = UsuarioService.m_ObjUsuarioConectado.EsAdministrador
    Application.TempVars("EsCalidad") = UsuarioService.m_ObjUsuarioConectado.EsCalidad
    
    EVE = "OK"
    Exit Function

errores:
    p_Error = "Error crítico en el arranque (EVE): " & Err.Description
    EVE = "ERROR"
End Function
