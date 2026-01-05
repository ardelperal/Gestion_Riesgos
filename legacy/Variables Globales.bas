Option Compare Database
Option Explicit

#If Win64 = 1 Then
    
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#Else
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
            ByVal lpApplicationName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As String, _
            ByVal nSize As Long, _
            ByVal lpFileName As String) As Long
    Public Declare Function Ejecutar Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hWnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    Public Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, lpExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    Public Declare  Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare  Function GetIpAddrTable Lib "Iphlpapi" ( _
            pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
#End If



Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&


Public pregunta As Long
Public lbl As Label
Public FSO As New FileSystemObject
Public Const SubRedOficina As String = "10.14.7"
Public m_ObjEntorno As Entorno

Public m_ObjProyectoActivo As Proyecto
Public m_ObjProyectoAlInicio As Proyecto


Public m_ObjEdicionActiva As Edicion
Public m_ObjRiesgoActivo As riesgo
Public m_EstadoRiesgoActivo As EnumRiesgoEstado
Public m_EsAlta As EnumSiNo
Public m_ObjRiesgoAlInicio As riesgo
Public blnEdicionActiva As Boolean
'Public blnPermitidoEditar As Boolean
Public m_ObjRiesgoExtActivo As RiesgoExterno
Public m_ObjPMActivo As PM
Public m_ObjPCActivo As PC
Public m_ObjPMAccionActiva As PMAccion
Public m_ObjPCAccionActiva As PCAccion
Public m_ObjNCActiva As NC
Public m_ObjSuministradorActivo As Suministrador
Public m_ObjRiesgoMaterializadoActivo As RiesgoMaterializacion
Public m_ObjRiesgoBibliotecaActivo As RiesgoBiblioteca



Public m_ObjTareasCalidad As TareasCalidad
Public m_ObjTareasTecnico As TareasTecnico
Public wks As DAO.Workspace
Private db As DAO.Database
Private db1 As DAO.Database
Public IDAplicacion As String
Public m_ObjUsuarioConectadoInicialmente As Usuario
Public m_ObjUsuarioConectado As Usuario
Public EsAdministrador As EnumSiNo
Public EsAdministradorConectadoInicialmente As EnumSiNo
Public EsCalidad As EnumSiNo
Public EsTecnico As EnumSiNo



Public t1 As Single
Public t2 As Single
Public varItem As Variant
Public m_EnOficina As EnumSiNo


Public m_ObjUsuarioParaTareas As Usuario

Public blnPermitidoEscribir As Boolean
Public m_ObjAnexoEvicenciaUTE As Anexo
Public m_URLInforme As String
Public m_URLHTMLActivo As String
Public m_URLRutaAplicacionesLocal As String
Public m_URLRutaAplicacionesRemotas As String
Public m_URLRutaAplicacionLocal As String
Public m_URLRutaAplicacionRemota As String
Public m_ListaUsuarios As String
Public m_ObjUltimoProyecto As UltimoProyecto

Public Const USE_INDICADORES_V2 As Boolean = True '/ False


Public Function getNombreUsuarioConectado(Optional ByRef p_Error As String) As String
    
    Dim m_UsuarioMaquina As Usuario
    
    On Error GoTo errores
    
    If Not m_ObjUsuarioConectado Is Nothing Then
        getNombreUsuarioConectado = m_ObjUsuarioConectado.Nombre
        Exit Function
    End If
    
    Set m_UsuarioMaquina = Constructor.getUsuarioConectadoPorMaquina(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioMaquina Is Nothing Then
        getNombreUsuarioConectado = "Desconocido"
        Exit Function
    End If
    getNombreUsuarioConectado = m_UsuarioMaquina.Nombre
    Exit Function
errores:
    getNombreUsuarioConectado = "Desconocido"
End Function
Public Function ReiniciarLasVariables(Optional ByRef p_Error As String) As String
    
    
    
    On Error GoTo errores
   
    Set m_ObjProyectoActivo = Nothing
    Set m_ObjProyectoAlInicio = Nothing
    Set m_ObjEdicionActiva = Nothing
    Set m_ObjRiesgoActivo = Nothing
    m_EsAlta = Empty
    Set m_ObjRiesgoAlInicio = Nothing
    blnEdicionActiva = False
    
    
    
    Set m_ObjRiesgoExtActivo = Nothing
    Set m_ObjPMActivo = Nothing
    Set m_ObjPMAccionActiva = Nothing
    Set m_ObjPCActivo = Nothing
    Set m_ObjPCAccionActiva = Nothing
    
    Set m_ObjNCActiva = Nothing
    Set m_ObjSuministradorActivo = Nothing
    Set m_ObjRiesgoMaterializadoActivo = Nothing
    Set m_ObjRiesgoBibliotecaActivo = Nothing
    
    Set m_ObjTareasCalidad = Nothing
    Set m_ObjTareasTecnico = Nothing
    Set m_ObjUsuarioConectado = Nothing
    EsAdministrador = Empty
    EsCalidad = Empty
    EsTecnico = Empty
    
    Set m_ObjUsuarioParaTareas = Nothing
    blnPermitidoEscribir = False
    Set m_ObjAnexoEvicenciaUTE = Nothing
    m_URLInforme = ""
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo ReiniciarLasVariables ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function EVE( _
                    Optional ByRef p_CorreoUsuario As String, _
                    Optional ByRef p_Error As String _
                    ) As String

    Dim m_NombreCarpeta As String
    Dim m_NombreCampo As Variant
    Dim m_valor As String
    Dim m_Objeto As Object
    Dim ti As Single
    Dim tf As Single
    Dim m_UsuarioLogeadoEnOrdenador As String
    
    Dim m_TipoCampo As String
    Dim m_ValorCampo As String
   
    Dim m_ValorCampoTruncado As String
    Dim m_Command As String 'se obtienen cuando se abre la base de datos con parámteros
    Dim m_ComandoResultante As String
    Dim m_Linea As String
    Dim m_ClaveValor As String
    Dim objNetwork As Object
    Dim m_GenerarInformeTipo As String
    
    Dim intNumeroErrores As Integer
    Dim m_CadenaCamposConError As String
    On Error GoTo errores
    
    ti = Timer
    
    If p_CorreoUsuario <> "" Then
        ReiniciarLasVariables p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    m_Linea = "Vaciando variables de estado....."
    Debug.Print m_Linea
    Avance m_Linea
    
    
    'Debug.Print m_TextoWin64
    Application.TempVars.RemoveAll
    Application.TempVars("CadenaJerarquicaModelo") = "nuevo"
    'Application.TempVars("CadenaJerarquicaModelo") = "antiguo"
    Application.TempVars("JPMesesAvisoEntreEdiciones") = 3
    Application.TempVars("JPDiasPreviosParaElAviso") = 15
    Application.TempVars("CalDiaInicialMesAviso") = 2
    Application.TempVars("GenerarInformeTipo") = "Excel"
    m_GenerarInformeTipo = Replace(UCase$(Trim$(CStr(Application.TempVars("GenerarInformeTipo")))), " ", "")
    Application.TempVars("GenerarInformeEnWord") = IIf(m_GenerarInformeTipo = "WORD" Or m_GenerarInformeTipo = "DOCX", "Sí", "No")
    VBA.DoEvents
    'Debug.Print "JPMesesAvisoEntreEdiciones", Application.TempVars("JPMesesAvisoEntreEdiciones")
    'Debug.Print "JPDiasPreviosParaElAviso", Application.TempVars("JPDiasPreviosParaElAviso")
    'Debug.Print "CalDiaInicialMesAviso", Application.TempVars("CalDiaInicialMesAviso")
    VBA.DoEvents
    
    Application.TempVars("DatosEnLocal") = "No"
    'Application.TempVars("DatosEnLocal") = "Sí"
    Application.TempVars("EnDesarrollo") = "No"
    Application.TempVars("EnPruebas") = "No"
    'Application.TempVars("EnPruebas") = "Sí"
        
    If Application.TempVars("EnPruebas") = "Sí" Then
        IDAplicacion = "51"
    Else
        IDAplicacion = "5"
    End If
   VBA.DoEvents
   'Debug.Print "DatosEnLocal", Application.TempVars("DatosEnLocal")
   'Debug.Print "EnDesarrollo", Application.TempVars("EnDesarrollo")
   'Debug.Print "EnPruebas", Application.TempVars("EnPruebas")
   VBA.DoEvents
    Set m_ObjEntorno = New Entorno
    m_URLRutaAplicacionesRemotas = "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_NombreCarpeta = "GESTION RIESGOS PRUEBA"
    Else
        m_NombreCarpeta = "GESTION RIESGOS"
    End If
    
    m_URLRutaAplicacionRemota = m_URLRutaAplicacionesRemotas & m_NombreCarpeta & "\"
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URLRutaAplicacionesLocal = getRutaAplicacionesLocal(p_Error)
        If m_URLRutaAplicacionesLocal <> "" Then
            m_URLRutaAplicacionLocal = m_URLRutaAplicacionesLocal & m_NombreCarpeta & "\"
        End If
    Else
       m_URLRutaAplicacionesLocal = ""
       m_URLRutaAplicacionLocal = ""
    End If
    m_Command = Nz(VBA.Command, "")
   ' m_Command = "beatriz.novalgutierrez@telefonica.com"
   ' m_Command = "marta.garridovaamonde@telefonica.com"
    'm_Command = "rosamaria.fuentesherrero@telefonica.com"
    'm_Command = "felix.sanchezpimentel@telefonica.com"
    'm_Command = "sergio.garciamontalvo@telefonica.com"
    'm_Command = "juan.jerezgarcia@telefonica.com"
    'm_Command = "javier.amousanos@telefonica.com"
    'm_Command = "jose.perezdionisio@telefonica.com"
    'm_Command = "carlos.alonsocarmona@telefonica.com"
    'm_Command = "juliobenedicto.vicariomancho@telefonica.com"
    'm_Command = "mario.martinabad@telefonica.com"
    'm_Command = "anamaria.rubiocanales@telefonica.com"
    'm_Command = "natalia.casangarcia@telefonica.com"
    'm_Command = "fernando.lazarodiaz@telefonica.com"
    'm_Command = "blanca.aguadovicente@telefonica.com"
    t1 = Timer
    If p_CorreoUsuario <> "" Then
        m_Command = p_CorreoUsuario
    End If
    If m_Command <> "" Then
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, , , m_Command, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Set objNetwork = CreateObject("Wscript.Network")
        m_UsuarioLogeadoEnOrdenador = objNetwork.UserName
        If m_UsuarioLogeadoEnOrdenador = "Local1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        If m_UsuarioLogeadoEnOrdenador = "adm1" Then m_UsuarioLogeadoEnOrdenador = "adm"
        Set m_ObjUsuarioConectado = Constructor.getUsuario(, m_UsuarioLogeadoEnOrdenador, , , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set objNetwork = Nothing
    End If
    If m_ObjUsuarioConectado Is Nothing Then
        p_Error = "No se ha podido determinar el usuario que está usando la herramienta"
        Err.Raise 1000
    End If
    If m_ObjUsuarioConectadoInicialmente Is Nothing Then
        Set m_ObjUsuarioConectadoInicialmente = m_ObjUsuarioConectado
    End If
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(m_ObjUsuarioConectadoInicialmente.UsuarioRed) Then
        EsAdministradorConectadoInicialmente = EnumSiNo.Sí
    Else
        EsAdministradorConectadoInicialmente = EnumSiNo.No
    End If
    EsAdministrador = m_ObjEntorno.UsuarioConectadoEsAdministrador
    If EsAdministrador <> EnumSiNo.Sí Then
        
        EsCalidad = m_ObjEntorno.UsuarioConectadoEsDeCalidad
        If EsCalidad <> EnumSiNo.Sí Then
            EsTecnico = EnumSiNo.Sí
        Else
            EsTecnico = EnumSiNo.No
        End If
    End If
    
    
    t2 = Timer
    If EsAdministrador = EnumSiNo.Sí Then
        VBA.DoEvents
        'Debug.Print "EsAdministrador:Sí"
        VBA.DoEvents
        EsCalidad = EnumSiNo.No
        EsTecnico = EnumSiNo.No
    End If
    If EsCalidad = EnumSiNo.Sí Then
        VBA.DoEvents
        'Debug.Print "EsCalidad:Sí"
        VBA.DoEvents
        EsAdministrador = EnumSiNo.No
        EsTecnico = EnumSiNo.No
    Else
        VBA.DoEvents
        'Debug.Print "EsTecnico:Sí"
        VBA.DoEvents
    End If
    
    t2 = Timer
    
    VBA.DoEvents
    'Debug.Print "CargarUsuario: " & m_ObjUsuarioConectado.Nombre & vbTab & "T:" & t2 - t1
    VBA.DoEvents
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_EnOficina = EnumSiNo.No
    Else
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
  
    
    For Each m_NombreCampo In m_ObjEntorno.ColItems.keys
        'Debug.Print m_nombreCampo
        'If CStr(m_nombreCampo) = "ColUsuarios" Then Stop
        Avance m_NombreCampo
        m_TipoCampo = m_ObjEntorno.ColItems(m_NombreCampo)
        If m_TipoCampo = "o" Then
            Set m_Objeto = m_ObjEntorno.getPropiedad(m_NombreCampo, p_Error)
            If p_Error <> "" Then
                If m_CadenaCamposConError = "" Then
                    m_CadenaCamposConError = m_NombreCampo
                Else
                    m_CadenaCamposConError = m_CadenaCamposConError & vbNewLine & m_NombreCampo
                End If
                intNumeroErrores = intNumeroErrores + 1
                p_Error = ""
            End If
            
        Else
            m_ValorCampo = m_ObjEntorno.getPropiedad(m_NombreCampo, p_Error)
            If p_Error <> "" Then
                If m_CadenaCamposConError = "" Then
                    m_CadenaCamposConError = m_NombreCampo
                Else
                    m_CadenaCamposConError = m_CadenaCamposConError & vbNewLine & m_NombreCampo
                End If
                intNumeroErrores = intNumeroErrores + 1
                p_Error = ""
            End If
            m_ValorCampoTruncado = Left(m_ValorCampo, 10) & " ..."
            
        End If
    Next
    t1 = Timer
    Avance "Cargando Perfiles"
    
    
    
    
    If m_ObjEntorno.VerSoloRiesgosNoRetirados = Empty Then
        m_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí
    End If
    If m_ObjEntorno.VerRiesgosDescripcion = Empty Then
        m_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí
    End If
    Set m_ObjUsuarioParaTareas = Nothing
    
   
    
    tf = Timer
    VBA.DoEvents
    Debug.Print "EVE en ......." & tf - ti
    If intNumeroErrores > 0 Then
        p_Error = "Se han producido los siguientes Errores: " & vbNewLine & m_CadenaCamposConError
        Err.Raise 1000
    End If
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo EVE ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
    Debug.Print p_Error
    
End Function

Public Function getdbLanzadera( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Lanzadera_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "0Lanzadera\Lanzadera_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbLanzadera = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbLanzadera ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdb(Optional ByRef p_Error As String, Optional p_db As DAO.Database = Nothing) As DAO.Database
    If Not p_db Is Nothing Then
        Set getdb = p_db
        Exit Function
    End If
    
    Dim m_URL As String
    Dim m_NombreDatos As String
    On Error GoTo errores
    
    m_NombreDatos = "Gestion_Riesgos_Datos.accdb"
    
    
    If Application.TempVars("EnPruebas") = "Sí" Then
        If Application.TempVars("DatosEnLocal") = "Sí" Then
            m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\" & m_NombreDatos
        Else
            m_URL = m_URLRutaAplicacionRemota & m_NombreDatos
        End If
    Else
        If Application.TempVars("DatosEnLocal") = "Sí" Then
            m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\" & m_NombreDatos
        Else
            m_URL = m_URLRutaAplicacionRemota & m_NombreDatos
        End If
    End If
    
    

    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdb = db
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdb ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdbExpedientes(Optional ByRef p_Error As String) As DAO.Database
    Dim m_URL As String
    On Error GoTo errores
   
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Expedientes_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "EXPEDIENTES\Expedientes_datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbExpedientes = db
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbExpedientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdbNC( _
                        Optional ByRef p_Error As String _
                        ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\NoConformidades_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "No Conformidades\NoConformidades_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
    
   
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbNC = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbNC ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdbAGEDO( _
                            Optional ByRef p_Error As String _
                            ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\AGEDO20_Datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "AGEDO\AGEDO20_Datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
    
     
   
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbAGEDO = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbAGEDO ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdbAgedys( _
                            Optional ByRef p_Error As String _
                            ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\AGEDYS_DATOS.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "Proyectos\AGEDYS_DATOS.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbAgedys = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbAgedys ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function getdbCorreo( _
                            Optional ByRef p_Error As String _
                            ) As DAO.Database
    
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Correos_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "00Recursos\Correos_datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
   
    
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbCorreo = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbCorreo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getdbControlCambios( _
                                Optional ByRef p_Error As String _
                                ) As DAO.Database
        
    Dim m_URL As String
    On Error GoTo errores
    
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_URL = m_URLRutaAplicacionesLocal & "000datoslocal\Control_Cambios_datos.accdb"
    ElseIf Application.TempVars("DatosEnLocal") = "No" Then
        m_URL = m_URLRutaAplicacionesRemotas & "CONTROL CAMBIOS\Control_Cambios_datos.accdb"
    Else
        p_Error = "No se conoce el origen de los datos"
        Err.Raise 1000
    End If
    
    
    Set wks = DBEngine.Workspaces(0)
    Set db1 = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbControlCambios = db1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getdbControlCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function LeerIni(Key As String, Default As Variant) As String
    Dim bufer As String * 256, Len_Value As Long
    
    
    Len_Value = GetPrivateProfileString(FSO.GetBaseName(CurrentDb().Name), _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         m_ObjEntorno.URLAchivoIni)
    LeerIni = Left$(bufer, CLng(Len_Value))
    
End Function



Public Function GetIPAddresses(Optional FilterLocalhost As Boolean = False) As String

    Dim Ret As Long
    Dim Buffer() As Byte
    Dim IPTableRow As IPINFO
    Dim Count As Long
    Dim BufferRequired As Long
    Dim StructSize As Long
    Dim NumIPAddresses As Long
    Dim IPAddress As String

  
        
    Call GetIpAddrTable(ByVal 0&, BufferRequired, 1)

    If BufferRequired > 0 Then
        
        ReDim Buffer(0 To BufferRequired - 1) As Byte
        
        If GetIpAddrTable(Buffer(0), BufferRequired, 1) = 0 Then
        
            'We've successfully obtained the IP Address details...
            'First 4 bytes is a long indicating the number of entries in the table
            StructSize = LenB(IPTableRow)
            CopyMemory NumIPAddresses, Buffer(0), 4
        
            While Count < NumIPAddresses
            
                'Buffer contains the IPINFO structures (after initial 4 byte long)
                CopyMemory IPTableRow, Buffer(4 + (Count * StructSize)), StructSize
                    
                IPAddress = IPAddressToString(IPTableRow.dwAddr)
                    
                If Not ((IPAddress = "127.0.0.1") _
                        And FilterLocalhost) Then
                            
                    'Replace this with whatever you want to do with the IP Address...
                    GetIPAddresses = GetIPAddresses & IPAddress & ";     "
                        
                End If
                
                Count = Count + 1
                
            Wend
            
        End If
            
    End If
 
    Exit Function



End Function
    
Private Function IPAddressToString(EncodedAddress As Long) As String
        
    Dim IPBytes(3) As Byte
    Dim Count As Long
        
    'Converts a long IP Address to a string formatted 255.255.255.255
    'Note: Could use inet_ntoa instead
        
    CopyMemory IPBytes(0), EncodedAddress, 4 ' IP Address is stored in four bytes (255.255.255.255)
        
    'Convert the 4 byte values to a formatted string
    While Count < 4
        
        IPAddressToString = IPAddressToString & _
                                CStr(IPBytes(Count)) & _
                                IIf(Count < 3, ".", "")

        Count = Count + 1
            
    Wend
        
End Function

